#!/usr/bin/env python3

from __future__ import annotations

import argparse
from pathlib import Path

from docx import Document
from docx.enum.text import WD_COLOR_INDEX

from docx_utils import apply_highlight_to_runs


ALLOWED_KEYS = {"NAME", "MONTH"}
GOAL_LABEL_TEXT = "本月精進目標:"


def resolve_template_path(template_path: Path) -> Path:
    if template_path.is_absolute() or template_path.exists():
        return template_path
    return Path(__file__).resolve().parent / template_path


def parse_input(path: Path) -> dict[str, str]:
    if path.suffix.lower() != ".txt":
        raise ValueError(f"Unsupported input format: {path}")

    data: dict[str, str] = {}
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        if ":" not in raw_line:
            continue
        key, value = raw_line.split(":", 1)
        normalized_key = key.strip().upper()
        if normalized_key not in ALLOWED_KEYS:
            continue
        data[normalized_key] = value.strip()

    for key in ALLOWED_KEYS:
        data.setdefault(key, "")
    return data


def replace_placeholders(doc: Document, mapping: dict[str, str]) -> None:
    for paragraph in doc.paragraphs:
        text = paragraph.text
        if not text:
            continue
        updated = text
        for key, value in mapping.items():
            updated = updated.replace(f"{{{{{key}}}}}", value)
        if updated != text:
            paragraph.text = updated


def apply_review_highlights(doc: Document) -> None:
    for paragraph in doc.paragraphs:
        if paragraph.text.strip() == GOAL_LABEL_TEXT:
            apply_highlight_to_runs(
                paragraph,
                highlight_color=WD_COLOR_INDEX.YELLOW,
            )


def generate_review(template_path: Path, input_path: Path, output_path: Path) -> None:
    data = parse_input(input_path)
    doc = Document(str(resolve_template_path(template_path)))
    replace_placeholders(doc, data)
    apply_review_highlights(doc)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Render review DOCX from template and text input."
    )
    parser.add_argument(
        "--template",
        default="templates/review_template.docx",
        help="Path to the review DOCX template.",
    )
    parser.add_argument(
        "--source-txt",
        default="review.txt",
        help="Path to the review text input.",
    )
    parser.add_argument(
        "--output",
        default="output/review_output.docx",
        help="Path to write the rendered review DOCX.",
    )
    args = parser.parse_args()

    generate_review(Path(args.template), Path(args.source_txt), Path(args.output))


if __name__ == "__main__":
    main()
