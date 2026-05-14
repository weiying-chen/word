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
TASKS_KEY = "tasks"
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


def _format_year_month_text(value: str) -> str:
    text = value.strip()
    if not text:
        return ""
    if len(text) == 7 and text[4] == "-":
        year_text, month_text = text.split("-", 1)
        if year_text.isdigit() and month_text.isdigit():
            return f"{int(year_text)}年{int(month_text)}月"
    return text


def parse_tasks_payload(path: Path) -> list[dict]:
    raw = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(raw, list):
        return raw
    if isinstance(raw, dict):
        return raw.get(TASKS_KEY, [])
    return []


def derive_month_from_tasks(path: Path, tasks: list[dict]) -> str:
    raw = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(raw, dict):
        export_month = _format_year_month_text(str(raw.get("exportMonth", "")))
        if export_month:
            return export_month

    deadlines: list[datetime] = []
    for task in tasks:
        deadline_text = str(task.get("deadline", "")).strip()
        if not deadline_text:
            deadline_text = str(task.get("deadlineIso", "")).strip()
        if not deadline_text:
            continue
        try:
            deadlines.append(datetime.fromisoformat(deadline_text.replace("Z", "+00:00")))
        except ValueError:
            continue

    if not deadlines:
        return ""

    latest = max(deadlines)
    return f"{latest.year}年{latest.month}月"


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


def _format_content_seconds(content_seconds: int | str | None) -> str:
    if content_seconds in (None, ""):
        return ""
    total = int(content_seconds)
    minutes, seconds = divmod(total, 60)
    parts: list[str] = []
    if minutes:
        parts.append(f"{minutes}分")
    if seconds:
        parts.append(f"{seconds}秒")
    return "".join(parts) if parts else "0秒"


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


def fill_regular_translation_table(doc: Document, tasks: list[dict]) -> None:
    if not doc.tables:
        return
    table = doc.tables[0]
    row_indexes = _ensure_regular_translation_row_count(table, len(tasks))
    for slot, row_idx in enumerate(row_indexes):
        task = tasks[slot] if slot < len(tasks) else None
        if not task:
            _set_cell_lines(table.cell(row_idx, 0), [])
            _set_cell_lines(table.cell(row_idx, 1), [])
            _set_cell_lines(table.cell(row_idx, 2), [])
            _set_cell_lines(table.cell(row_idx, 3), [])
            continue

        _set_cell_lines(
            table.cell(row_idx, 0),
            [
                _format_month_day(
                    str(
                        task.get("createdAt", "")
                        or task.get("deadline", "")
                        or task.get("deadlineIso", "")
                    ).strip()
                )
            ],
        )
        item_lines = [
            f"{slot + 1}.",
            str(task.get("name", "") or task.get("title", "")).strip(),
        ]
        length_text = _format_content_seconds(task.get("contentSeconds"))
        if length_text:
            item_lines.append(f"長度:{length_text}")
        work_text = _format_work_minutes(task.get("workMinutes"))
        if work_text:
            item_lines.append(f"實際作業時間:{work_text}")
        _set_cell_lines(table.cell(row_idx, 1), [line for line in item_lines if line])

        comments = task.get("comments", [])
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
    tasks_path: Path,
) -> None:
    data = parse_input(input_path)
    tasks = parse_tasks_payload(tasks_path)
    data[MONTH_KEY] = derive_month_from_tasks(tasks_path, tasks)
    doc = Document(str(resolve_template_path(template_path)))
    replace_placeholders(doc, data)
    apply_header_font_size(doc)
    apply_review_highlights(doc)
    fill_regular_translation_table(doc, tasks)
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
        "--tasks-json",
        default="tasks.json",
        help="Path to tasks JSON that provides exportMonth.",
    )
    args = parser.parse_args()

    generate_review(
        Path(args.template),
        Path(args.source_txt),
        Path(args.output),
        Path(args.tasks_json),
    )


if __name__ == "__main__":
    main()
