#!/usr/bin/env python3

from __future__ import annotations

import argparse
import re
from pathlib import Path

from docx import Document


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


def _clean_header_title(title_line: str) -> str:
    return TRANSLATOR_TAG_RE.sub("", title_line).strip()


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
            entries.append(
                {
                    "filename_title": normalize_title(title_line),
                    "header_title": _clean_header_title(title_line),
                    "header_url": url_line,
                }
            )

    return entries


def extract_post_titles(schedule_path: Path) -> list[str]:
    return [entry["filename_title"] for entry in extract_post_entries(schedule_path)]


def replace_placeholders(doc: Document, mapping: dict[str, str]) -> None:
    for paragraph in doc.paragraphs:
        text = paragraph.text
        for placeholder, value in mapping.items():
            if placeholder in text:
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
        replace_placeholders(
            doc,
            {
                "{{HEADER_TITLE}}": entry["header_title"],
                "{{HEADER_URL}}": entry["header_url"],
            },
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
