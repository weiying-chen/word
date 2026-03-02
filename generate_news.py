#!/usr/bin/env python3

from __future__ import annotations

import argparse
import re
import warnings
from pathlib import Path

from docx import Document
from docx.enum.text import WD_COLOR_INDEX

from docx_utils import add_hyperlink
from generate_subs import (
    apply_default_margins,
    ensure_hyperlink_style,
    fix_docx_namespaces,
    normalize_input_text,
    remove_paragraph,
)


PLACEHOLDER_KEYS = [
    "TITLE",
    "TITLE_URL",
    "SUMMARY",
    "META_TITLE_EN",
    "META_OVERVIEW_EN",
    "SUPER_PEOPLE",
    "BODY",
]
PLACEHOLDER_KEY_SET = set(PLACEHOLDER_KEYS)
SHOT_ID_RE = re.compile(r"^\d+_\d+$")
FIXED_MARKER = "<"
DEFAULT_TEMPLATE = Path("templates/news_template.docx")


def _decode_input_text(path: Path) -> tuple[str, str, bool]:
    raw = path.read_bytes()
    tried_encodings: list[str] = []

    if raw.startswith(b"\xef\xbb\xbf"):
        return raw.decode("utf-8-sig"), "utf-8-sig", False

    if raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff"):
        return raw.decode("utf-16"), "utf-16", False

    for encoding in ("utf-8", "big5", "cp950", "gb18030", "cp1252"):
        tried_encodings.append(encoding)
        try:
            return raw.decode(encoding), encoding, encoding != "utf-8"
        except UnicodeDecodeError:
            continue

    tried = ", ".join(tried_encodings)
    raise UnicodeError(
        f"Unable to decode input file '{path}' with supported encodings: {tried}"
    )


def parse_input(path: Path) -> dict[str, str]:
    data: dict[str, str] = {}
    text, encoding_used, used_fallback = _decode_input_text(path)
    if used_fallback:
        warnings.warn(
            f"Using fallback encoding '{encoding_used}' for {path}; rewriting as UTF-8.",
            stacklevel=2,
        )
        path.write_text(text, encoding="utf-8")

    lines = text.splitlines()
    idx = 0
    while idx < len(lines):
        raw_line = lines[idx]
        if ":" not in raw_line:
            idx += 1
            continue

        key, value = raw_line.split(":", 1)
        key = key.lstrip("\ufeff").strip().upper()
        value = value.lstrip()

        if key not in PLACEHOLDER_KEY_SET:
            idx += 1
            continue

        if key in {"SUMMARY", "META_OVERVIEW_EN", "SUPER_PEOPLE", "BODY"}:
            collected: list[str] = []
            if value:
                collected.append(value)
            idx += 1
            while idx < len(lines):
                next_line = lines[idx]
                if ":" in next_line:
                    next_key = next_line.split(":", 1)[0].strip().upper()
                    if next_key in PLACEHOLDER_KEY_SET:
                        break
                collected.append(next_line)
                idx += 1
            data[key] = normalize_input_text("\n".join(collected).rstrip())
            continue

        data[key] = normalize_input_text(value)
        idx += 1

    data.setdefault("SUMMARY", "")
    data.setdefault("META_OVERVIEW_EN", "")
    data.setdefault("BODY", "")
    return data


def _add_plain_paragraph(doc: Document, text: str) -> None:
    paragraph = doc.add_paragraph("")
    if not text:
        return
    run = paragraph.add_run(text)
    run.font.highlight_color = (
        WD_COLOR_INDEX.BRIGHT_GREEN
        if SHOT_ID_RE.match(text.strip())
        else WD_COLOR_INDEX.WHITE
    )


def _render_multiline_block(doc: Document, text: str) -> None:
    for line in text.splitlines():
        _add_plain_paragraph(doc, line)


def _new_document_from_template(template_path: Path) -> Document:
    doc = Document(str(template_path))
    for paragraph in list(doc.paragraphs):
        remove_paragraph(paragraph)
    return doc


def default_output_path(source_docx: Path, output_dir: Path) -> Path:
    stem = source_docx.stem
    if not stem.endswith("_final"):
        stem = f"{stem}_final"
    return output_dir / f"{stem}.docx"


def generate_news(
    input_path: Path,
    output_path: Path,
    template_path: Path = DEFAULT_TEMPLATE,
) -> None:
    data = parse_input(input_path)
    doc = _new_document_from_template(template_path)
    apply_default_margins(doc)
    ensure_hyperlink_style(doc)

    title = data.get("TITLE", "").strip()
    title_url = data.get("TITLE_URL", "").strip()
    if title:
        title_paragraph = doc.add_paragraph("")
        if title_url:
            add_hyperlink(title_paragraph, title, title_url)
        else:
            title_paragraph.add_run(title)
        doc.add_paragraph("")

    summary = data.get("SUMMARY", "")
    if summary:
        _render_multiline_block(doc, summary)

    super_people = data.get("SUPER_PEOPLE", "")
    if super_people:
        doc.add_paragraph("")
        _render_multiline_block(doc, super_people)

    if summary or super_people or data.get("BODY", ""):
        _add_plain_paragraph(doc, FIXED_MARKER)
        doc.add_paragraph("")

    body = data.get("BODY", "")
    if body:
        _render_multiline_block(doc, body)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    fix_docx_namespaces(output_path)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Render a newsroom DOCX from a structured text input."
    )
    parser.add_argument(
        "--input",
        default="news_input.txt",
        help="Path to the news input text file.",
    )
    parser.add_argument(
        "--output",
        default="",
        help="Path to write the generated DOCX.",
    )
    parser.add_argument(
        "--source-docx",
        required=True,
        help="Original source DOCX for naming the output.",
    )
    parser.add_argument(
        "--template",
        default=str(DEFAULT_TEMPLATE),
        help="Path to the base DOCX template for styles.",
    )
    args = parser.parse_args()

    output_path = Path(args.output) if args.output else default_output_path(
        Path(args.source_docx), Path("output")
    )

    generate_news(Path(args.input), output_path, Path(args.template))


if __name__ == "__main__":
    main()
