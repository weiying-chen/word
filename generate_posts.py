#!/usr/bin/env python3

from __future__ import annotations

import argparse
import re
from pathlib import Path

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.shared import Inches, Pt


PERSON_LINE_RE = re.compile(r"^\d+\.\s*(\S+)")
PROGRAM_SECTION_RE = re.compile(r"^節目.*則")
STOP_SECTION_RE = re.compile(r"^(?:-+|FB小編文|本周節日)")
TRANSLATOR_TAG_RE = re.compile(r"\s*[A-Za-z]+/[A-Za-z]+\s*$")


def iter_non_empty_paragraphs(doc: Document) -> list[str]:
    lines: list[str] = []
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if text:
            lines.append(text)
    return lines


def normalize_title(title_line: str) -> str:
    title = title_line.replace(" - ", " ").strip()
    title = TRANSLATOR_TAG_RE.sub("", title).strip()
    title = title.replace("/", "")
    return title


def _clean_title_for_display(title_line: str) -> str:
    return TRANSLATOR_TAG_RE.sub("", title_line).strip()


def _extract_reference(lines: list[str], start_idx: int) -> tuple[str, str]:
    for idx in range(start_idx, len(lines)):
        line = lines[idx]
        if PERSON_LINE_RE.match(line) or STOP_SECTION_RE.match(line):
            break
        if line.strip() == "搭配":
            ref_url = lines[idx + 1] if idx + 1 < len(lines) else ""
            ref_title = lines[idx + 2] if idx + 2 < len(lines) else ""
            return ref_url, ref_title
    return "", ""


def extract_post_entries(schedule_path: Path) -> list[dict[str, str]]:
    doc = Document(str(schedule_path))
    lines = iter_non_empty_paragraphs(doc)
    entries: list[dict[str, str]] = []
    in_program_section = False

    for idx, line in enumerate(lines):
        if not in_program_section:
            if PROGRAM_SECTION_RE.match(line):
                in_program_section = True
            continue

        if STOP_SECTION_RE.match(line):
            break

        match = PERSON_LINE_RE.match(line)
        if not match:
            continue

        person = match.group(1).strip().lower()
        if idx + 1 >= len(lines):
            continue
        title_line = lines[idx + 1]
        url_line = lines[idx + 2] if idx + 2 < len(lines) else ""
        if not url_line.startswith("http"):
            url_line = ""
        if person == "alex":
            ref_url, ref_title = _extract_reference(lines, idx + 1)
            entries.append(
                {
                    "filename_title": normalize_title(title_line),
                    "header_title": _clean_title_for_display(title_line),
                    "header_url": url_line,
                    "video_url": url_line,
                    "video_title": _clean_title_for_display(title_line),
                    "ref_url": ref_url,
                    "ref_title": ref_title,
                }
            )

    return entries


def extract_post_titles(schedule_path: Path) -> list[str]:
    return [entry["filename_title"] for entry in extract_post_entries(schedule_path)]


def get_default_tab_stop_inches(doc: Document) -> float:
    settings = doc.part.settings.element
    node = settings.find(qn("w:defaultTabStop"))
    if node is None:
        return 0.5
    value = node.get(qn("w:val"))
    if not value:
        return 0.5
    return int(value) / 1440


def clear_paragraph(paragraph) -> None:
    for run in paragraph.runs:
        run._element.getparent().remove(run._element)


def add_highlighted_run(paragraph, text: str) -> None:
    run = paragraph.add_run(text)
    run.font.size = Pt(10)
    run.font.highlight_color = WD_COLOR_INDEX.TURQUOISE


def add_highlighted_hyperlink(paragraph, text: str, url: str) -> None:
    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    h_run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")
    r_style = OxmlElement("w:rStyle")
    r_style.set(qn("w:val"), "Hyperlink")
    r_pr.append(r_style)
    r_color = OxmlElement("w:color")
    r_color.set(qn("w:val"), "0563C1")
    r_pr.append(r_color)
    r_highlight = OxmlElement("w:highlight")
    r_highlight.set(qn("w:val"), "cyan")
    r_pr.append(r_highlight)
    r_sz = OxmlElement("w:sz")
    r_sz.set(qn("w:val"), "20")
    r_pr.append(r_sz)
    r_u = OxmlElement("w:u")
    r_u.set(qn("w:val"), "single")
    r_pr.append(r_u)
    h_run.append(r_pr)

    t = OxmlElement("w:t")
    t.text = text
    h_run.append(t)

    hyperlink.append(h_run)
    paragraph._p.append(hyperlink)


def replace_placeholders(
    doc: Document, mapping: dict[str, str], indent_inches: float
) -> None:
    highlight_keys = {"{{REF_TITLE}}", "{{VIDEO_TITLE}}"}
    hyperlink_keys = {"{{HEADER_URL}}", "{{REF_URL}}", "{{VIDEO_URL}}"}
    indent_keys = {
        "{{REF_URL}}",
        "{{REF_TITLE}}",
        "{{REF_SUMMARY_ZH}}",
        "{{REF_TITLE_EN}}",
        "{{REF_SUMMARY_EN}}",
        "{{VIDEO_URL}}",
        "{{VIDEO_TITLE}}",
        "{{VIDEO_DESC_EN}}",
        "{{VIDEO_DESC_ZH}}",
    }

    for paragraph in doc.paragraphs:
        text = paragraph.text
        for placeholder, value in mapping.items():
            if placeholder not in text:
                continue
            if placeholder in indent_keys:
                paragraph.paragraph_format.left_indent = Inches(indent_inches)
                paragraph.paragraph_format.first_line_indent = 0
            if placeholder in hyperlink_keys and value:
                clear_paragraph(paragraph)
                add_highlighted_hyperlink(paragraph, value, value)
                text = paragraph.text
                continue
            if placeholder in highlight_keys and value:
                clear_paragraph(paragraph)
                add_highlighted_run(paragraph, value)
                text = paragraph.text
                continue
            text = text.replace(placeholder, value)
        if text != paragraph.text:
            paragraph.text = text


def make_unique_path(base: Path) -> Path:
    if not base.exists():
        return base
    stem = base.stem
    suffix = base.suffix
    counter = 2
    while True:
        candidate = base.with_name(f"{stem}_{counter}{suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def generate_docs(
    schedule_path: Path,
    template_path: Path,
    output_dir: Path,
    filename_prefix: str,
    filename_suffix: str,
) -> list[Path]:
    entries = extract_post_entries(schedule_path)
    output_paths: list[Path] = []
    for entry in entries:
        filename = f"{filename_prefix}{entry['filename_title']}{filename_suffix}.docx"
        output_path = make_unique_path(output_dir / filename)
        doc = Document(str(template_path))
        default_tab_stop = get_default_tab_stop_inches(doc)
        replace_placeholders(
            doc,
            {
                "{{HEADER_TITLE}}": entry["header_title"],
                "{{HEADER_URL}}": entry["header_url"],
                "{{REF_URL}}": entry["ref_url"],
                "{{REF_TITLE}}": entry["ref_title"],
                "{{VIDEO_URL}}": entry["video_url"],
                "{{VIDEO_TITLE}}": entry["video_title"],
            },
            default_tab_stop,
        )
        doc.save(str(output_path))
        output_paths.append(output_path)
    return output_paths


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generate empty post docs from alex entries in a schedule DOCX."
    )
    parser.add_argument(
        "--schedule",
        default="260302排程_ev_k.docx",
        help="Path to the schedule DOCX.",
    )
    parser.add_argument(
        "--template",
        default="templates/post_template.docx",
        help="Path to the base DOCX template with shared styles.",
    )
    parser.add_argument(
        "--output-dir",
        default="outputs",
        help="Directory to write generated DOCX files.",
    )
    parser.add_argument(
        "--prefix",
        default="日期未定_",
        help="Filename prefix.",
    )
    parser.add_argument(
        "--suffix",
        default="_al",
        help="Filename suffix (without extension).",
    )
    args = parser.parse_args()

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    generate_docs(
        schedule_path=Path(args.schedule),
        template_path=Path(args.template),
        output_dir=output_dir,
        filename_prefix=args.prefix,
        filename_suffix=args.suffix,
    )


if __name__ == "__main__":
    main()
