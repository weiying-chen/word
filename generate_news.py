#!/usr/bin/env python3

from __future__ import annotations

import argparse
import re
import warnings
from pathlib import Path

from docx import Document
from docx.enum.text import WD_COLOR_INDEX

from generate_subs import fix_docx_namespaces, normalize_input_text, remove_paragraph


SHOT_ID_RE = re.compile(r"^\d+_\d+$")
BODY_LABEL_LINE_RE = re.compile(r"^\s*(BODY|字幕)\s*[:：]\s*$")
BODY_INLINE_LINE_RE = re.compile(r"^\s*(BODY|字幕)\s*[:：]\s*(.*)$")
MARKER_TEXT = "<"


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
    text, encoding_used, used_fallback = _decode_input_text(path)
    if used_fallback:
        warnings.warn(
            f"Using fallback encoding '{encoding_used}' for {path}; rewriting as UTF-8.",
            stacklevel=2,
        )
        path.write_text(text, encoding="utf-8")

    lines = text.splitlines()

    for idx, raw_line in enumerate(lines):
        stripped = raw_line.strip()

        inline = BODY_INLINE_LINE_RE.match(stripped)
        if inline and not BODY_LABEL_LINE_RE.match(stripped):
            collected = [inline.group(2)] if inline.group(2) else []
            collected.extend(lines[idx + 1 :])
            return {"BODY": normalize_input_text("\n".join(collected).rstrip())}

        if BODY_LABEL_LINE_RE.match(stripped):
            return {
                "BODY": normalize_input_text("\n".join(lines[idx + 1 :]).rstrip())
            }

    return {"BODY": normalize_input_text(text.rstrip())}


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


def default_output_path(source_docx: Path, output_dir: Path) -> Path:
    stem = source_docx.stem
    if not stem.endswith("_final"):
        stem = f"{stem}_final"
    return output_dir / f"{stem}.docx"


def _marker_index(doc: Document) -> int:
    for idx, paragraph in enumerate(doc.paragraphs):
        if paragraph.text.strip() == MARKER_TEXT:
            return idx
    raise ValueError("source.docx must contain a '<' marker paragraph.")


def _trim_existing_body(doc: Document, marker_idx: int) -> None:
    for paragraph in list(doc.paragraphs[marker_idx + 1 :]):
        remove_paragraph(paragraph)


def generate_news(
    source_docx_path: Path,
    input_path: Path,
    output_path: Path,
) -> None:
    data = parse_input(input_path)
    generate_news_from_data(source_docx_path, data, output_path)


def generate_news_from_data(
    source_docx_path: Path,
    data: dict[str, str],
    output_path: Path,
) -> None:
    doc = Document(str(source_docx_path))
    marker_idx = _marker_index(doc)
    _trim_existing_body(doc, marker_idx)

    body = data.get("BODY", "")
    if body:
        doc.add_paragraph("")
        _render_multiline_block(doc, body)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    fix_docx_namespaces(output_path)


def generate_news_from_sources(
    source_docx_path: Path,
    source_txt_path: Path,
    output_path: Path,
) -> None:
    generate_news(source_docx_path, source_txt_path, output_path)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Render a newsroom DOCX by preserving the DOCX header and replacing the body from text."
    )
    parser.add_argument(
        "--source-txt",
        default="source.txt",
        help="Path to the body text source.",
    )
    parser.add_argument(
        "--output",
        default="",
        help="Path to write the generated DOCX.",
    )
    parser.add_argument(
        "--source-docx",
        required=True,
        help="Original source DOCX whose header is preserved.",
    )
    args = parser.parse_args()

    output_path = Path(args.output) if args.output else default_output_path(
        Path(args.source_docx), Path("output")
    )
    generate_news(Path(args.source_docx), Path(args.source_txt), output_path)


if __name__ == "__main__":
    main()
