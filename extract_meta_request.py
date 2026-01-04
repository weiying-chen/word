#!/usr/bin/env python3

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from docx import Document


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


def extract_from_docx(path: Path) -> dict:
    doc = Document(str(path))
    lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

    narration_zh: list[str] = []
    supers: list[dict] = []
    report_zh: list[str] = []
    english_names: list[str] = []
    super_lines: list[str] = []
    in_super = False
    in_report = False

    for raw in lines:
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

    title_zh = narration_zh[0] if narration_zh else ""
    summary_zh = narration_zh[1] if len(narration_zh) > 1 else ""

    people = [
        {
            "name_zh": s.get("name_zh", ""),
            "name_en": "",
            "role_zh": s.get("role_zh", ""),
            "role_en": "",
        }
        for s in supers
    ]
    for idx, en_name in enumerate(english_names):
        if idx >= len(people):
            break
        people[idx]["name_en"] = en_name

    return {
        "title_zh": title_zh,
        "summary_zh": summary_zh,
        "narration_zh": narration_zh,
        "supers_zh": supers,
        "report_zh": report_zh,
        "people": people,
        "title_en": "",
        "overview_en": "",
    }


def build_prompt(payload: dict) -> str:
    return "\n".join(
        [
            "You are given JSON. Fill only the *_en fields with English translations/writing.",
            "Do not change the structure or any *_zh fields.",
            "Output JSON only, no extra text.",
            "",
            json.dumps(payload, ensure_ascii=False, indent=2),
            "",
        ]
    )


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extract meta request JSON and a GPT prompt from main.docx."
    )
    parser.add_argument("--source", default="main.docx", help="Path to main.docx.")
    parser.add_argument(
        "--json-out",
        default="",
        help="Optional path to write JSON (leave empty to skip).",
    )
    parser.add_argument(
        "--prompt-out", default="meta_prompt.txt", help="Where to write prompt text."
    )
    args = parser.parse_args()

    payload = extract_from_docx(Path(args.source))

    if args.json_out:
        json_out = Path(args.json_out)
        json_out.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )

    prompt_out = Path(args.prompt_out)
    prompt_out.write_text(build_prompt(payload), encoding="utf-8")


if __name__ == "__main__":
    main()
