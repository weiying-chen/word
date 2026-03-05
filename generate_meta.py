#!/usr/bin/env python3

from __future__ import annotations

import argparse
import re
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from docx.shared import Inches

from generate_news import parse_input as parse_news_input


TITLE_PLACEHOLDER = "{{TITLE_EN}}"
PEOPLE_PLACEHOLDER = "{{PEOPLE}}"
OVERVIEW_PLACEHOLDER = "{{OVERVIEW_EN}}"
META_TITLE_EN_KEY = "META_TITLE_EN"
META_OVERVIEW_EN_KEY = "META_OVERVIEW_EN"
META_PEOPLE_KEY = "META_PEOPLE"
CJK_RE = re.compile(r"[\u4e00-\u9fff]")
EN_NAME_PAREN_RE = re.compile(r"^\(\s*\d+\s+([A-Za-z][A-Za-z.\s'-]*)\s*\)$")


def _contains_cjk(text: str) -> bool:
    return bool(CJK_RE.search(text))


def _clean_super_line(text: str) -> str:
    cleaned = text.strip()
    if cleaned.endswith("//"):
        cleaned = cleaned[:-2].rstrip()
    return cleaned


def _parse_super(lines: list[str]) -> dict:
    role_zh = ""
    name_zh = ""
    quotes_zh: list[str] = []
    if lines:
        header = lines[0]
        if "│" in header:
            role_zh, name_zh = [part.strip() for part in header.split("│", 1)]
        else:
            name_zh = header.strip()
        if len(lines) > 1:
            quotes_zh = [line for line in lines[1:] if line]
    return {
        "role_zh": role_zh,
        "name_zh": name_zh,
        "quotes_zh": quotes_zh,
    }


def _parse_meta_people_blocks(text: str) -> list[dict[str, str]]:
    if not text.strip():
        return []

    blocks: list[list[str]] = []
    current: list[str] = []
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            if current:
                blocks.append(current)
                current = []
            continue
        current.append(line)
    if current:
        blocks.append(current)

    entries: list[dict[str, str]] = []
    for block in blocks:
        if not block:
            continue
        label_zh = block[0].strip()
        name_zh = ""
        role_zh = ""
        if "｜" in label_zh:
            role_zh, name_zh = [part.strip() for part in label_zh.split("｜", 1)]
        else:
            name_zh = label_zh.strip()

        entries.append(
            {
                "label_zh": label_zh,
                "name_zh": name_zh,
                "role_zh": role_zh,
                "name_en": block[1].strip() if len(block) > 1 else "",
                "role_en": block[2].strip() if len(block) > 2 else "",
                "org_en": block[3].strip() if len(block) > 3 else "",
            }
        )
    return entries


def _merge_meta_people_overrides(
    people: list[dict[str, str]],
    overrides: list[dict[str, str]],
) -> list[dict[str, str]]:
    if not overrides:
        return people

    merged = [dict(person) for person in people]
    for person in merged:
        role_zh = person.get("role_zh", "").strip()
        name_zh = person.get("name_zh", "").strip()
        label_zh = f"{role_zh}｜{name_zh}" if role_zh and name_zh else (role_zh or name_zh)

        match = next(
            (
                entry
                for entry in overrides
                if entry.get("label_zh", "").strip()
                and entry.get("label_zh", "").strip() == label_zh
            ),
            None,
        )
        if match is None and name_zh:
            match = next(
                (
                    entry
                    for entry in overrides
                    if entry.get("name_zh", "").strip() == name_zh
                ),
                None,
            )
        if match is None:
            continue

        for key in ("name_en", "role_en", "org_en"):
            value = match.get(key, "").strip()
            if value:
                person[key] = value

    return merged


def parse_input(path: Path) -> dict[str, object]:
    if path.suffix.lower() != ".txt":
        raise ValueError(f"Unsupported input format: {path}")

    data = parse_news_input(path)
    body_lines = data.get("BODY", "").splitlines()

    narration_zh: list[str] = []
    supers: list[dict[str, object]] = []
    report_zh: list[str] = []
    english_names: list[str] = []
    super_lines: list[str] = []
    in_super = False
    in_report = False

    for raw in body_lines:
        line = raw.strip()
        en_match = EN_NAME_PAREN_RE.match(line)
        if en_match:
            english_names.append(en_match.group(1).strip())
            continue
        if line.startswith("/*SUPER"):
            in_super = True
            super_lines = []
            continue
        if line.startswith("/*REPORT"):
            in_report = True
            continue
        if in_report:
            if line.startswith("*/"):
                in_report = False
            else:
                report_zh.append(_clean_super_line(line))
            continue
        if in_super:
            if line.startswith("*/"):
                supers.append(_parse_super(super_lines))
                in_super = False
            else:
                super_lines.append(_clean_super_line(line))
            continue

        if not _contains_cjk(line):
            continue
        if re.fullmatch(r"[\d_]+", line):
            continue
        if re.fullmatch(r"\(?\s*\d+.*\)?", line) and not _contains_cjk(line):
            continue
        if re.fullmatch(r"\(\s*NS\s*\)", line):
            continue
        if re.fullmatch(r"\(\s*\d+\s+[A-Za-z]+\s*\)", line):
            continue

        line = re.sub(r"^\(\s*[^)]*\)\s*", "", line).strip()
        if not line:
            continue
        narration_zh.append(line)

    people: list[dict[str, str]] = [
        {
            "name_zh": str(s.get("name_zh", "")),
            "name_en": "",
            "role_zh": str(s.get("role_zh", "")),
            "role_en": "",
            "org_en": "",
        }
        for s in supers
    ]
    for idx, en_name in enumerate(english_names):
        if idx >= len(people):
            break
        people[idx]["name_en"] = en_name

    summary = data.get("SUMMARY", "").splitlines()
    meta_people_text = data.get(META_PEOPLE_KEY, "")
    overrides = _parse_meta_people_blocks(meta_people_text)
    people = _merge_meta_people_overrides(people, overrides)
    return {
        "title_zh": data.get("TITLE", ""),
        "summary_zh": summary[0] if summary else "",
        "narration_zh": narration_zh,
        "supers_zh": supers,
        "report_zh": report_zh,
        "people": people,
        "title_en": data.get(META_TITLE_EN_KEY, ""),
        "overview_en": data.get(META_OVERVIEW_EN_KEY, ""),
    }


def remove_paragraph(paragraph: Paragraph) -> None:
    element = paragraph._element
    element.getparent().remove(element)


def insert_paragraph_after(paragraph: Paragraph, text: str) -> Paragraph:
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    new_para.add_run(text)
    return new_para


def find_paragraph_by_text(doc: Document, text: str) -> Paragraph | None:
    for paragraph in doc.paragraphs:
        if paragraph.text.strip() == text:
            return paragraph
    return None


def replace_multiline(paragraph: Paragraph, lines: list[str]) -> None:
    paragraph.text = ""
    if not lines:
        remove_paragraph(paragraph)
        return
    paragraph.add_run(lines[0])
    current = paragraph
    for line in lines[1:]:
        current = insert_paragraph_after(current, line)


def build_people_lines(people: list[dict]) -> list[str]:
    lines: list[str] = []
    for idx, person in enumerate(people):
        label_zh = person.get("label_zh")
        if not label_zh:
            role_zh = person.get("role_zh", "").strip()
            name_zh = person.get("name_zh", "").strip()
            if role_zh and name_zh:
                label_zh = f"{role_zh}｜{name_zh}"
            else:
                label_zh = role_zh or name_zh
        lines.append(label_zh or "")
        name_zh = person.get("name_zh", "").strip()
        name_en = person.get("name_en", "").strip()
        if not name_en:
            placeholder_key = name_zh or "NAME_EN"
            name_en = f"{{{{{placeholder_key}}}}}"
        lines.append(name_en)
        role_zh = person.get("role_zh", "").strip()
        role_en = person.get("role_en", "").strip()
        if not role_en:
            placeholder_key = role_zh or "ROLE_EN"
            role_en = f"{{{{{placeholder_key}}}}}"
        lines.append(role_en)
        org_en = person.get("org_en", "").strip()
        if org_en:
            lines.append(org_en)
        if idx < len(people) - 1:
            lines.append("")
    return lines


def apply_default_margins(doc: Document) -> None:
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.25)
        section.right_margin = Inches(1.25)


def default_output_path(source_docx: Path, output_dir: Path) -> Path:
    stem = source_docx.stem
    if stem.endswith("_final"):
        stem = stem[: -len("_final")]
    return output_dir / f"{stem}_標題職銜_final.docx"


def generate_meta(template_path: Path, input_path: Path, output_path: Path) -> None:
    data = parse_input(input_path)
    doc = Document(str(template_path))
    apply_default_margins(doc)

    title_placeholder = find_paragraph_by_text(doc, TITLE_PLACEHOLDER)
    if title_placeholder:
        title_placeholder.text = data.get("title_en", "")

    people_placeholder = find_paragraph_by_text(doc, PEOPLE_PLACEHOLDER)
    if people_placeholder:
        replace_multiline(people_placeholder, build_people_lines(data.get("people", [])))

    overview_placeholder = find_paragraph_by_text(doc, OVERVIEW_PLACEHOLDER)
    if overview_placeholder:
        overview_placeholder.text = data.get("overview_en", "")

    doc.save(str(output_path))


def main() -> None:
    parser = argparse.ArgumentParser(description="Render meta.docx from shared news txt data.")
    parser.add_argument(
        "--template",
        default="templates/meta_template.docx",
        help="Path to the meta DOCX template.",
    )
    parser.add_argument(
        "--input",
        default="news_input.txt",
        help="Path to the shared news txt input.",
    )
    parser.add_argument(
        "--output",
        default="",
        help="Path to write the rendered meta DOCX.",
    )
    parser.add_argument(
        "--source-docx",
        required=True,
        help="Original source DOCX for naming the output.",
    )
    args = parser.parse_args()

    output_path = Path(args.output) if args.output else default_output_path(
        Path(args.source_docx), Path("output")
    )

    generate_meta(Path(args.template), Path(args.input), output_path)


if __name__ == "__main__":
    main()
