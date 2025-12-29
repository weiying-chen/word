#!/usr/bin/env python3

from __future__ import annotations

import argparse
from pathlib import Path

from docx import Document


LABEL_MAP = {
    "建議YT標題": "YT_TITLE_SUGGESTED",
    "建議標題": "TITLE_SUGGESTED",
    "選圖": "THUMBNAIL",
}


def parse_source_docx(path: Path) -> tuple[str, str, str, str]:
    doc = Document(str(path))
    non_empty = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    if len(non_empty) < 4:
        raise ValueError("source.docx must include title, url, summary, time range.")
    return non_empty[0], non_empty[1], non_empty[2], non_empty[3]


def parse_source_txt(path: Path) -> tuple[dict[str, str], str]:
    lines = path.read_text(encoding="utf-8").splitlines()
    fields: dict[str, str] = {}
    body_lines: list[str] = []
    idx = 0

    while idx < len(lines):
        line = lines[idx].strip()
        if line.endswith("："):
            label = line[:-1]
            idx += 1
            collected = []
            while idx < len(lines) and lines[idx].strip() != "":
                collected.append(lines[idx])
                idx += 1
            if label in LABEL_MAP:
                fields[LABEL_MAP[label]] = "\n".join(collected).strip()
            while idx < len(lines) and lines[idx].strip() == "":
                idx += 1
            if label == "字幕":
                body_lines = lines[idx:]
                break
        else:
            idx += 1

    while body_lines and body_lines[0].strip() == "":
        body_lines.pop(0)

    body = "\n".join(body_lines).rstrip()
    return fields, body


def write_input(
    output_path: Path,
    title: str,
    url: str,
    summary: str,
    time_range: str,
    fields: dict[str, str],
    body: str,
) -> None:
    output_path.write_text(
        "\n".join(
            [
                f"TITLE: {title}",
                f"URL: {url}",
                f"SUMMARY: {summary}",
                "",
                f"YT_TITLE_SUGGESTED: {fields.get('YT_TITLE_SUGGESTED', '')}",
                f"TITLE_SUGGESTED: {fields.get('TITLE_SUGGESTED', '')}",
                f"THUMBNAIL: {fields.get('THUMBNAIL', '')}",
                "",
                f"TIME_RANGE: {time_range}",
                "",
                "BODY:",
                body,
                "",
            ]
        ),
        encoding="utf-8",
    )


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Build input.txt from source.docx and source.txt."
    )
    parser.add_argument("--source-docx", default="source.docx")
    parser.add_argument("--source-txt", default="source.txt")
    parser.add_argument("--output", default="input.txt")
    args = parser.parse_args()

    title, url, summary, time_range = parse_source_docx(Path(args.source_docx))
    fields, body = parse_source_txt(Path(args.source_txt))
    write_input(Path(args.output), title, url, summary, time_range, fields, body)


if __name__ == "__main__":
    main()
