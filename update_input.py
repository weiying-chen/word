#!/usr/bin/env python3

from __future__ import annotations

import argparse
import re
from pathlib import Path

from docx import Document


LABEL_MAP = {
    # Chinese source labels
    "建議YT標題": "YT_TITLE_SUGGESTED",
    "建議標題": "TITLE_SUGGESTED",
    "簡介": "INTRO",
    "選圖": "THUMBNAIL",
    "SUPER_PEOPLE": "SUPER_PEOPLE",
    # English source labels (news-native)
    "YT_TITLE_SUGGESTED": "YT_TITLE_SUGGESTED",
    "TITLE_SUGGESTED": "TITLE_SUGGESTED",
    "INTRO": "INTRO",
    "THUMBNAIL": "THUMBNAIL",
    "TITLE_TEXT": "TITLE_TEXT",
    "TITLE_URL": "TITLE_URL",
    "SUMMARY": "SUMMARY",
    "BODY": "BODY",
}

BODY_LABELS = {"字幕", "BODY"}
LABEL_LINE_RE = re.compile(r"^\s*([^:：]+)\s*[:：]\s*$")
INLINE_LINE_RE = re.compile(r"^\s*([^:：]+)\s*[:：]\s*(.*)$")


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

    def parse_label_line(raw: str) -> str | None:
        m = LABEL_LINE_RE.match(raw.strip())
        if not m:
            return None
        return m.group(1).strip()

    def is_label_line(raw: str) -> bool:
        label = parse_label_line(raw)
        if not label:
            return False
        return label in LABEL_MAP or label in BODY_LABELS

    while idx < len(lines):
        raw = lines[idx]

        # Inline labels: KEY: value
        inline = INLINE_LINE_RE.match(raw.strip())
        if inline and not LABEL_LINE_RE.match(raw.strip()):
            label = inline.group(1).strip()
            value = inline.group(2)
            mapped = LABEL_MAP.get(label)
            if mapped:
                fields[mapped] = value.strip("\n")
            idx += 1
            continue

        label = parse_label_line(raw)
        if not label:
            idx += 1
            continue

        if label in BODY_LABELS:
            body_lines = lines[idx + 1 :]
            break

        idx += 1
        collected = []
        while idx < len(lines) and not is_label_line(lines[idx]):
            collected.append(lines[idx])
            idx += 1

        mapped = LABEL_MAP.get(label)
        if mapped:
            fields[mapped] = "\n".join(collected).strip("\n")

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
    intro = fields.get("INTRO", "")
    intro_lines = intro.splitlines() if intro else [""]
    super_people = fields.get("SUPER_PEOPLE", "")
    super_people_lines = super_people.splitlines() if super_people else [""]
    output_path.write_text(
        "\n".join(
            [
                f"TITLE: {title}",
                f"URL: {url}",
                f"SUMMARY: {summary}",
                "",
                f"YT_TITLE_SUGGESTED: {fields.get('YT_TITLE_SUGGESTED', '')}",
                f"TITLE_SUGGESTED: {fields.get('TITLE_SUGGESTED', '')}",
                "INTRO:",
                *intro_lines,
                f"THUMBNAIL: {fields.get('THUMBNAIL', '')}",
                "",
                f"TIME_RANGE: {time_range}",
                "",
                "SUPER_PEOPLE:",
                *super_people_lines,
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
