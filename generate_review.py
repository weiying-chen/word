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


GOAL_LABEL_TEXT = "本月精進目標:"
MONTH_KEY = "MONTH"
HEADER_FONT_SIZE_PT = REVIEW_TEXT_SIZE_PT
REVIEWER_NAME = "陳威穎"


def resolve_template_path(template_path: Path) -> Path:
    if template_path.is_absolute() or template_path.exists():
        return template_path
    return Path(__file__).resolve().parent / template_path


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
    if not isinstance(raw, list):
        raise ValueError("tasks.json must be a top-level JSON array.")
    return raw


def derive_month_from_tasks(tasks: list[dict]) -> str:
    created_times: list[datetime] = []
    for task in tasks:
        created_at = str(task.get("createdAt", "")).strip()
        if not created_at:
            continue
        try:
            created_times.append(datetime.fromisoformat(created_at.replace("Z", "+00:00")))
        except ValueError:
            continue

    if not created_times:
        return ""

    latest = max(created_times)
    return f"{latest.year}年{latest.month}月"


def _format_month_day(iso_text: str) -> str:
    if not iso_text:
        return ""
    text = iso_text.replace("Z", "+00:00")
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


def _extract_feedback_lines(task: dict) -> list[str]:
    raw_notes = task.get("notes")
    if isinstance(raw_notes, list):
        return [f"• {str(note).strip()}" for note in raw_notes if str(note).strip()]

    raw_comments = task.get("comments")
    if isinstance(raw_comments, list):
        return [f"• {str(note).strip()}" for note in raw_comments if str(note).strip()]
    return []


def _find_regular_translation_rows(table) -> list[int]:
    rows: list[int] = []
    for idx in range(1, len(table.rows)):
        first = table.cell(idx, 0).text.strip()
        second = table.cell(idx, 1).text.strip()
        if first == "日期" and "審稿" in second:
            break
        rows.append(idx)
    return rows


def _find_block(table, start_col0: str, start_col1_contains: str, end_col0_prefix: str) -> tuple[int, int]:
    start_idx = None
    next_section_idx = None
    for idx in range(len(table.rows)):
        first = table.cell(idx, 0).text.strip()
        second = table.cell(idx, 1).text.strip()
        if start_idx is None and first == start_col0 and start_col1_contains in second:
            start_idx = idx + 1
            continue
        if start_idx is not None and first.startswith(end_col0_prefix):
            next_section_idx = idx
            break
    if start_idx is None or next_section_idx is None or start_idx >= next_section_idx:
        raise ValueError("Unable to locate expected row block in template.")
    return start_idx, next_section_idx


def _remove_row(table, row_idx: int) -> None:
    table.rows[row_idx]._tr.getparent().remove(table.rows[row_idx]._tr)


def _insert_cloned_row_before(table, source_row_idx: int, before_row_idx: int) -> None:
    cloned = deepcopy(table.rows[source_row_idx]._tr)
    table.rows[before_row_idx]._tr.addprevious(cloned)


def _find_regular_translation_block(table) -> tuple[int, int]:
    return _find_block(table, "日期", "字幕翻譯", "日期")


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


def _find_temp_work_block(table) -> tuple[int, int]:
    return _find_block(table, "日期", "臨時工作", "本月工作心得")


def _ensure_temp_work_row_count(table, desired_count: int) -> list[int]:
    start_idx, next_section_idx = _find_temp_work_block(table)
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


def _collect_temp_posts(tasks: list[dict]) -> list[dict]:
    posts: list[dict] = []
    for task in tasks:
        children = task.get("children", [])
        if not isinstance(children, list):
            continue
        for child in children:
            if not isinstance(child, dict):
                continue
            if str(child.get("type", "")).strip().lower() != "posts":
                continue
            posts.append(child)
    return posts


def remove_subtitle_review_section(doc: Document) -> None:
    if not doc.tables:
        return
    table = doc.tables[0]
    start_idx = None
    end_idx = None

    for idx in range(len(table.rows)):
        first = table.cell(idx, 0).text.strip()
        second = table.cell(idx, 1).text.strip()
        if start_idx is None and first == "日期" and "字幕審稿" in second:
            start_idx = idx
            continue
        if start_idx is not None and first == "日期" and "臨時工作" in second:
            end_idx = idx
            break

    if start_idx is None or end_idx is None or start_idx >= end_idx:
        return

    for idx in range(end_idx - 1, start_idx - 1, -1):
        _remove_row(table, idx)


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
                    str(task.get("createdAt", "")).strip()
                )
            ],
        )
        item_lines = [
            f"{slot + 1}.",
            str(task.get("name", "")).strip(),
        ]
        length_text = _format_content_seconds(task.get("contentSeconds"))
        if length_text:
            item_lines.append(f"長度:{length_text}")
        work_text = _format_work_minutes(task.get("workMinutes"))
        if work_text:
            item_lines.append(f"實際作業時間:{work_text}")
        _set_cell_lines(table.cell(row_idx, 1), [line for line in item_lines if line])

        _set_cell_lines(table.cell(row_idx, 2), _extract_feedback_lines(task))
        _set_cell_lines(table.cell(row_idx, 3), [])


def fill_temp_work_table(doc: Document, tasks: list[dict]) -> None:
    if not doc.tables:
        return
    table = doc.tables[0]
    posts = _collect_temp_posts(tasks)
    row_indexes = _ensure_temp_work_row_count(table, len(posts))

    for slot, row_idx in enumerate(row_indexes):
        post = posts[slot] if slot < len(posts) else None
        if not post:
            _set_cell_lines(table.cell(row_idx, 0), [])
            _set_cell_lines(table.cell(row_idx, 1), [])
            _set_cell_lines(table.cell(row_idx, 2), [])
            _set_cell_lines(table.cell(row_idx, 3), [])
            continue

        _set_cell_lines(
            table.cell(row_idx, 0),
            [_format_month_day(str(post.get("createdAt", "")).strip())],
        )
        item_lines = [f"{slot + 1}.", str(post.get("name", "")).strip()]
        work_text = _format_work_minutes(post.get("workMinutes"))
        if work_text:
            item_lines.append(f"實際作業時間:{work_text}")
        _set_cell_lines(table.cell(row_idx, 1), [line for line in item_lines if line])
        _set_cell_lines(table.cell(row_idx, 2), _extract_feedback_lines(post))
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
    output_path: Path,
    tasks_path: Path,
) -> None:
    tasks = parse_tasks_payload(tasks_path)
    data = {"NAME": REVIEWER_NAME, MONTH_KEY: derive_month_from_tasks(tasks)}
    doc = Document(str(resolve_template_path(template_path)))
    replace_placeholders(doc, data)
    apply_header_font_size(doc)
    apply_review_highlights(doc)
    fill_regular_translation_table(doc, tasks)
    remove_subtitle_review_section(doc)
    fill_temp_work_table(doc, tasks)
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
        "--output",
        default="output/review_output.docx",
        help="Path to write the rendered review DOCX.",
    )
    parser.add_argument(
        "--tasks-json",
        default="tasks.json",
        help="Path to tasks JSON array.",
    )
    args = parser.parse_args()

    generate_review(
        Path(args.template),
        Path(args.output),
        Path(args.tasks_json),
    )


if __name__ == "__main__":
    main()
