#!/usr/bin/env python3

from __future__ import annotations

import argparse
import re
from pathlib import Path

from docx import Document

from docx_utils import (
    apply_font_size_to_document_runs,
    ensure_blank_after_labels,
    get_default_tab_stop_inches,
)
from prepare_posts import (
    ensure_blank_after_reference_url,
    fetch_youtube_video_descriptions,
    insert_video_section_spacing,
    normalize_empty_paragraphs,
    remove_paragraph,
    replace_placeholders,
    sync_empty_paragraph_indents,
)
from style_tokens import BODY_TEXT_SIZE_PT


TITLE_RE = re.compile(r"^##\s+(?P<title>.+?)\s*$")
SOURCES_RE = re.compile(r"^###\s+Sources\s*$", re.IGNORECASE)
VIDEO_RE = re.compile(r"^Video(?:\s+URL)?\s*:\s*(?P<url>https?://\S+)\s*$", re.I)
YOUTUBE_RE = re.compile(r"^https?://(?:www\.)?(?:youtube\.com|youtu\.be)/\S+$", re.I)


def _join_content(lines: list[str]) -> str:
    while lines and not lines[0].strip():
        lines.pop(0)
    while lines and not lines[-1].strip():
        lines.pop()
    return "\n".join(lines)


def parse_post_text(text: str) -> dict[str, str]:
    lines = text.lstrip("\ufeff").splitlines()
    title_idx = next(
        (idx for idx, line in enumerate(lines) if TITLE_RE.match(line.strip())),
        None,
    )
    if title_idx is None:
        raise ValueError("[error] Missing post title: expected a '## Title' line.")
    title_match = TITLE_RE.match(lines[title_idx].strip())
    assert title_match is not None
    title = title_match.group("title").strip()

    sources_idx = next(
        (
            idx
            for idx in range(title_idx + 1, len(lines))
            if SOURCES_RE.match(lines[idx].strip())
        ),
        len(lines),
    )
    content = lines[title_idx + 1 : sources_idx]
    source_lines = lines[sources_idx + 1 :] if sources_idx < len(lines) else []

    video_url = ""
    cleaned_content: list[str] = []
    for line in content:
        stripped = line.strip()
        video_match = VIDEO_RE.match(stripped)
        if video_match and not video_url:
            video_url = video_match.group("url")
            continue
        if YOUTUBE_RE.match(stripped) and not video_url:
            video_url = stripped
            continue
        cleaned_content.append(line)

    hashtag_indices = [
        idx for idx, line in enumerate(cleaned_content) if line.strip().startswith("#")
    ]
    if len(hashtag_indices) < 2:
        raise ValueError("[error] Expected English and Chinese hashtag lines.")
    en_hash_idx, zh_hash_idx = hashtag_indices[0], hashtag_indices[1]

    ref_url = ""
    ref_title_lines: list[str] = []
    for line in source_lines:
        stripped = line.strip()
        if not stripped:
            continue
        if stripped.startswith("http") and not ref_url:
            ref_url = stripped
        else:
            ref_title_lines.append(stripped)

    return {
        "title": title,
        "video_url": video_url,
        "post_en": _join_content(cleaned_content[:en_hash_idx]),
        "hashtags_en": cleaned_content[en_hash_idx].strip(),
        "post_zh": _join_content(cleaned_content[en_hash_idx + 1 : zh_hash_idx]),
        "hashtags_zh": cleaned_content[zh_hash_idx].strip(),
        "ref_url": ref_url,
        "ref_title": "\n".join(ref_title_lines),
    }


def parse_post_file(path: Path) -> dict[str, str]:
    return parse_post_text(path.read_text(encoding="utf-8-sig"))


def _remove_draft_scaffolding(doc: Document) -> None:
    remove_text = {
        "{{HEADER_TITLE}}",
        "{{HEADER_URL}}",
        "標題",
        "{{TITLE_LINE_1}}",
        "{{TITLE_LINE_2}}",
    }
    for paragraph in reversed(doc.paragraphs):
        if paragraph.text.strip() in remove_text:
            remove_paragraph(paragraph)
    while doc.paragraphs and not doc.paragraphs[0].text.strip():
        remove_paragraph(doc.paragraphs[0])


def _resolve_template_path(path: Path) -> Path:
    if path.is_absolute() or path.exists():
        return path
    return Path(__file__).resolve().parent / path


def generate_post(
    *,
    input_path: Path,
    template_path: Path,
    output_path: Path,
) -> Path:
    post = parse_post_file(input_path)
    video_url = post["video_url"]
    video_desc_en = ""
    video_desc_zh = ""
    if video_url:
        video_desc_en, video_desc_zh = fetch_youtube_video_descriptions(video_url)

    doc = Document(str(_resolve_template_path(template_path)))
    apply_font_size_to_document_runs(doc, font_size_pt=BODY_TEXT_SIZE_PT)
    _remove_draft_scaffolding(doc)
    default_tab_stop = get_default_tab_stop_inches(doc)
    post_en = post["title"]
    if post["post_en"]:
        post_en = f"{post_en}\n\n{post['post_en']}"
    mapping = {
        "{{POST_EN}}": post_en,
        "{{HASHTAGS_EN}}": post["hashtags_en"],
        "{{POST_ZH}}": post["post_zh"],
        "{{HASHTAGS_ZH}}": post["hashtags_zh"],
        "{{REF_URL}}": post["ref_url"],
        "{{REF_TITLE}}": post["ref_title"],
        "{{REF_SUMMARY_ZH}}": "",
        "{{REF_TITLE_EN}}": "",
        "{{REF_SUMMARY_EN}}": "",
        "{{VIDEO_URL}}": video_url,
        "{{VIDEO_TITLE}}": post["title"],
        "{{VIDEO_DESC_EN}}": video_desc_en,
        "{{VIDEO_DESC_ZH}}": video_desc_zh,
    }
    replace_placeholders(doc, mapping, default_tab_stop)
    insert_video_section_spacing(
        doc,
        video_title=post["title"],
        video_desc_en=video_desc_en,
        video_desc_zh=video_desc_zh,
        indent_inches=default_tab_stop,
    )
    ensure_blank_after_labels(doc, {"參考資料：", "英文翻譯：", "要用的影片："})
    normalize_empty_paragraphs(doc)
    sync_empty_paragraph_indents(doc)
    ensure_blank_after_reference_url(
        doc,
        ref_url=post["ref_url"],
        indent_inches=default_tab_stop,
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return output_path


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generate a finished post DOCX from completed post text."
    )
    parser.add_argument("--input", required=True, help="Completed post TXT file.")
    parser.add_argument(
        "--template",
        default="templates/post_template.docx",
        help="Post DOCX template.",
    )
    parser.add_argument("--output", default="", help="Output DOCX path.")
    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output) if args.output else input_path.with_suffix(".docx")
    generate_post(
        input_path=input_path,
        template_path=Path(args.template),
        output_path=output_path,
    )


if __name__ == "__main__":
    main()
