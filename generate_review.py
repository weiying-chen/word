#!/usr/bin/env python3

from __future__ import annotations

import argparse
import json
from copy import deepcopy
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.shared import Pt

from docx_utils import apply_highlight_to_runs, clear_paragraph
from style_tokens import REVIEW_TEXT_SIZE_PT


ALLOWED_KEYS = {"NAME"}
GOAL_LABEL_TEXT = "本月精進目標:"
MONTH_KEY = "MONTH"
ASSIGNMENTS_KEY = "assignments"
HEADER_FONT_SIZE_PT = REVIEW_TEXT_SIZE_PT


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


def parse_export_month(payload: dict) -> str:
    value = str(payload.get("exportMonth", "")).strip()
    if not value:
        return ""

    if len(value) == 7 and value[4] == "-":
        year_text, month_text = value.split("-", 1)
        if year_text.isdigit() and month_text.isdigit():
            return f"{int(year_text)}年{int(month_text)}月"
    return value


def parse_assignments_payload(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def _format_month_day(deadline_iso: str) -> str:
    if not deadline_iso:
        return ""
    text = deadline_iso.replace("Z", "+00:00")
    dt = datetime.fromisoformat(text)
    return f"{dt.month}/{dt.day}"


def _format_work_minutes(work_minutes: int | str | None) -> str:
    if work_minutes in (None, ""):
        return ""
    minutes = int(work_minutes)
    hours, remain = divmod(minutes, 60)
    parts: list[str] = []
    if hours:
        parts.append(f"{hours}時")
    if remain:
        parts.append(f"{remain}分")
    return "".join(parts) if parts else "0分"


def _set_cell_lines(cell, lines: list[str]) -> None:
    paragraph = cell.paragraphs[0]
    clear_paragraph(paragraph)
    if lines:
        for idx, line in enumerate(lines):
            run = paragraph.add_run(line)
            if idx < len(lines) - 1:
                run.add_break()
    for extra in list(cell.paragraphs[1:]):
        extra._element.getparent().remove(extra._element)


def _find_regular_translation_rows(table) -> list[int]:
    rows: list[int] = []
    for idx in range(1, len(table.rows)):
        first = table.cell(idx, 0).text.strip()
        second = table.cell(idx, 1).text.strip()
        if first == "日期" and "審稿" in second:
            break
        rows.append(idx)
    return rows


def _remove_row(table, row_idx: int) -> None:
    table.rows[row_idx]._tr.getparent().remove(table.rows[row_idx]._tr)


def _insert_cloned_row_before(table, source_row_idx: int, before_row_idx: int) -> None:
    cloned = deepcopy(table.rows[source_row_idx]._tr)
    table.rows[before_row_idx]._tr.addprevious(cloned)


def _find_regular_translation_block(table) -> tuple[int, int]:
    start_idx = None
    next_section_idx = None
    for idx in range(len(table.rows)):
        first = table.cell(idx, 0).text.strip()
        second = table.cell(idx, 1).text.strip()
        if start_idx is None and first == "日期" and "字幕翻譯" in second:
            start_idx = idx + 1
            continue
        if start_idx is not None and first == "日期" and "審稿" in second:
            next_section_idx = idx
            break
    if start_idx is None or next_section_idx is None or start_idx >= next_section_idx:
        raise ValueError("Unable to locate regular translation row block in template.")
    return start_idx, next_section_idx


def _ensure_regular_translation_row_count(table, desired_count: int) -> list[int]:
    start_idx, next_section_idx = _find_regular_translation_block(table)
    current_count = next_section_idx - start_idx
    target = max(1, desired_count)

    while current_count < target:
        _insert_cloned_row_before(table, start_idx, next_section_idx)
        next_section_idx += 1
        current_count += 1

    while current_count > target:
        _remove_row(table, next_section_idx - 1)
        next_section_idx -= 1
        current_count -= 1

    return list(range(start_idx, start_idx + target))


def fill_regular_translation_table(doc: Document, assignments: list[dict]) -> None:
    if not doc.tables:
        return
    table = doc.tables[0]
    row_indexes = _ensure_regular_translation_row_count(table, len(assignments))
    for slot, row_idx in enumerate(row_indexes):
        assignment = assignments[slot] if slot < len(assignments) else None
        if not assignment:
            _set_cell_lines(table.cell(row_idx, 0), [])
            _set_cell_lines(table.cell(row_idx, 1), [])
            _set_cell_lines(table.cell(row_idx, 2), [])
            _set_cell_lines(table.cell(row_idx, 3), [])
            continue

        _set_cell_lines(
            table.cell(row_idx, 0),
            [_format_month_day(str(assignment.get("deadlineIso", "")).strip())],
        )
        item_lines = [f"{slot + 1}.", str(assignment.get("title", "")).strip()]
        work_text = _format_work_minutes(assignment.get("workMinutes"))
        if work_text:
            item_lines.append(f"實際作業時間:{work_text}")
        _set_cell_lines(table.cell(row_idx, 1), [line for line in item_lines if line])

        comments = assignment.get("comments", [])
        comment_lines = [f"• {str(c).strip()}" for c in comments if str(c).strip()]
        _set_cell_lines(table.cell(row_idx, 2), comment_lines)
        _set_cell_lines(table.cell(row_idx, 3), [])


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


def apply_header_font_size(doc: Document, size_pt: int = HEADER_FONT_SIZE_PT) -> None:
    for idx, paragraph in enumerate(doc.paragraphs[:4]):
        if not paragraph.runs:
            paragraph.add_run(paragraph.text)
        for run in paragraph.runs:
            run.font.size = Pt(size_pt)


def generate_review(
    template_path: Path,
    input_path: Path,
    output_path: Path,
    assignments_path: Path,
) -> None:
    data = parse_input(input_path)
    payload = parse_assignments_payload(assignments_path)
    data[MONTH_KEY] = parse_export_month(payload)
    doc = Document(str(resolve_template_path(template_path)))
    replace_placeholders(doc, data)
    apply_header_font_size(doc)
    apply_review_highlights(doc)
    fill_regular_translation_table(doc, payload.get(ASSIGNMENTS_KEY, []))
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
    parser.add_argument(
        "--assignments-json",
        default="assignments.json",
        help="Path to assignments JSON that provides exportMonth.",
    )
    args = parser.parse_args()

    generate_review(
        Path(args.template),
        Path(args.source_txt),
        Path(args.output),
        Path(args.assignments_json),
    )


if __name__ == "__main__":
    main()
