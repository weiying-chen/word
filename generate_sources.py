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


def _render_docx(
    template_path: Path,
    output_path: Path,
    title: str,
    youtube_url: str,
    summary: str,
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

    p_time = doc.add_paragraph(HARDCODED_TIMESTAMP_LINE)
    for run in p_time.runs:
        run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    apply_font_size_to_runs(p_time, font_size_pt=BODY_TEXT_SIZE_PT)

    doc.add_paragraph("")
    p_summary = doc.add_paragraph(summary)
    apply_font_size_to_runs(p_summary, font_size_pt=BODY_TEXT_SIZE_PT)

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
    skipped_missing_subs = 0
    errors = 0

    for item in episodes:
        ep_id = str(item.get("epId", "")).strip()
        subtitle_file = _find_subtitle_file(sources_dir, ep_id)
        if subtitle_file is None:
            skipped_missing_subs += 1
            continue

        try:
            title = str(item.get("youtubeTitle", "")).strip() or str(
                item.get("titleZh", "")
            ).strip()
            youtube_url = str(item.get("youtubeUrl", "")).strip()
            summary = _first_summary_line(str(item.get("youtubeDescription", "")))
            if not title or not youtube_url:
                errors += 1
                continue
            filename = f"{_safe_filename(title)}.docx"
            output_path = output_dir / filename
            _render_docx(template_path, output_path, title, youtube_url, summary)
            generated += 1
        except Exception:
            errors += 1

    return {
        "generated": generated,
        "skipped_missing_subs": skipped_missing_subs,
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
        f"generated={result['generated']} skipped_missing_subs={result['skipped_missing_subs']} errors={result['errors']}"
    )


if __name__ == "__main__":
    main()
