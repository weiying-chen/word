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


def extract_post_titles(schedule_path: Path) -> list[str]:
    doc = Document(str(schedule_path))
    lines = iter_non_empty_paragraphs(doc)
    titles: list[str] = []
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
        if person == "alex":
            titles.append(normalize_title(title_line))

    return titles


def clear_document(doc: Document) -> None:
    for paragraph in list(doc.paragraphs):
        element = paragraph._element
        element.getparent().remove(element)
    doc.add_paragraph("")


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
    titles = extract_post_titles(schedule_path)
    output_paths: list[Path] = []
    for title in titles:
        filename = f"{filename_prefix}{title}{filename_suffix}.docx"
        output_path = make_unique_path(output_dir / filename)
        doc = Document(str(template_path))
        clear_document(doc)
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
