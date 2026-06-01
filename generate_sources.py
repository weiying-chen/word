#!/usr/bin/env python3

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from docx import Document
from docx.enum.text import WD_COLOR_INDEX

from docx_utils import add_hyperlink, apply_font_size_to_runs
from style_tokens import BODY_TEXT_SIZE_PT


HARDCODED_TIMESTAMP_LINE = "07:27-09:20 (1分53秒)"
DESCRIPTION_TIMESTAMP_RE = re.compile(r"^\s*(?P<mm>\d{1,2}):(?P<ss>\d{2})｜")


def _safe_filename(text: str) -> str:
    safe = text
    for ch in '<>:"/\\|?*':
        safe = safe.replace(ch, " ")
    safe = re.sub(r"\s+", " ", safe).strip()
    return safe


def _find_subtitle_file(sources_dir: Path, ep_id: str) -> Path | None:
    if not ep_id:
        return None
    for candidate in sorted(sources_dir.glob("*.txt")):
        if ":Zone.Identifier" in candidate.name:
            continue
        if f"第{ep_id}集_ch_" in candidate.name or candidate.name.startswith(f"{ep_id}_"):
            return candidate
    return None


def _first_summary_line(description: str) -> str:
    for line in description.splitlines():
        stripped = line.strip()
        if stripped:
            return stripped
    return ""


def _format_minutes_seconds(total_seconds: int) -> str:
    minutes, seconds = divmod(max(0, total_seconds), 60)
    return f"{minutes}分{seconds}秒"


def _dynamic_timestamp_line(last_line: str, subtitle_lines: list[str]) -> str:
    match = DESCRIPTION_TIMESTAMP_RE.match(last_line.strip())
    if not match:
        return HARDCODED_TIMESTAMP_LINE
    end_total = int(match.group("mm")) * 60 + int(match.group("ss"))
    start = "00:00"
    start_total = 0
    duration = max(0, end_total - start_total)
    end = f"{int(match.group('mm')):02d}:{int(match.group('ss')):02d}"
    return f"{start}-{end} ({_format_minutes_seconds(duration)})"


def _read_text_with_fallback(path: Path) -> str:
    raw = path.read_bytes()
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw.decode("utf-8-sig")
    if raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff"):
        return raw.decode("utf-16")
    for encoding in ("utf-8", "big5", "cp950", "gb18030", "cp1252"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="ignore")


def _render_docx(
    template_path: Path,
    output_path: Path,
    title: str,
    youtube_url: str,
    summary: str,
    subtitle_lines: list[str],
    timestamp_line: str,
) -> None:
    doc = Document(str(template_path))
    if not doc.paragraphs:
        doc.add_paragraph("")
    first = doc.paragraphs[0]
    first.text = ""
    first.add_run(title)
    apply_font_size_to_runs(first, font_size_pt=BODY_TEXT_SIZE_PT)

    p_url = doc.add_paragraph("")
    add_hyperlink(p_url, youtube_url, youtube_url)
    apply_font_size_to_runs(p_url, font_size_pt=BODY_TEXT_SIZE_PT)

    p_time = doc.add_paragraph(timestamp_line)
    for run in p_time.runs:
        run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    apply_font_size_to_runs(p_time, font_size_pt=BODY_TEXT_SIZE_PT)

    doc.add_paragraph("")
    p_summary = doc.add_paragraph(summary)
    apply_font_size_to_runs(p_summary, font_size_pt=BODY_TEXT_SIZE_PT)

    if subtitle_lines:
        doc.add_paragraph("")
        for line in subtitle_lines:
            p_line = doc.add_paragraph(line)
            apply_font_size_to_runs(p_line, font_size_pt=BODY_TEXT_SIZE_PT)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))


def generate_sources(
    *,
    episodes_json: Path,
    template_path: Path,
    sources_dir: Path,
    output_dir: Path,
) -> dict[str, int]:
    episodes = json.loads(episodes_json.read_text(encoding="utf-8"))
    generated = 0
    skipped = 0
    errors = 0

    for item in episodes:
        ep_id = str(item.get("epId", "")).strip()
        subtitle_file = _find_subtitle_file(sources_dir, ep_id)
        if subtitle_file is None:
            skipped += 1
            continue

        try:
            title = str(item.get("youtubeTitle", "")).strip() or str(
                item.get("titleZh", "")
            ).strip()
            youtube_url = str(item.get("youtubeUrl", "")).strip()
            summary = _first_summary_line(str(item.get("youtubeDescription", "")))
            last_ts_line = str(item.get("descriptionLastTimestampLine", "")).strip()
            subtitle_lines = [
                line.rstrip("\n")
                for line in _read_text_with_fallback(subtitle_file).splitlines()
                if line.strip()
            ]
            timestamp_line = _dynamic_timestamp_line(last_ts_line, subtitle_lines)
            if not title or not youtube_url:
                errors += 1
                continue
            filename = f"{_safe_filename(title)}.docx"
            output_path = output_dir / filename
            _render_docx(
                template_path,
                output_path,
                title,
                youtube_url,
                summary,
                subtitle_lines,
                timestamp_line,
            )
            generated += 1
        except Exception:
            errors += 1

    return {
        "generated": generated,
        "skipped": skipped,
        "errors": errors,
    }


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generate source DOCX files from episodes JSON and subtitle source files."
    )
    parser.add_argument("--episodes-json", required=True, help="Path to episodes.json")
    parser.add_argument(
        "--template",
        default="templates/sources_template.docx",
        help="Path to sources template docx.",
    )
    parser.add_argument(
        "--sources-dir",
        default="sources",
        help="Directory containing subtitle txt files.",
    )
    parser.add_argument(
        "--output-dir",
        default="output",
        help="Directory to write generated source docx files.",
    )
    args = parser.parse_args()

    result = generate_sources(
        episodes_json=Path(args.episodes_json),
        template_path=Path(args.template),
        sources_dir=Path(args.sources_dir),
        output_dir=Path(args.output_dir),
    )
    print(
        f"generated={result['generated']} skipped={result['skipped']} errors={result['errors']}"
    )


if __name__ == "__main__":
    main()
