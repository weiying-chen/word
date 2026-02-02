#!/usr/bin/env python3

from __future__ import annotations

import argparse
import re
from datetime import date
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

from docx_utils import (
    add_highlighted_run,
    add_hyperlink,
    apply_highlight_to_runs,
    clear_paragraph,
    ensure_blank_after_labels,
    get_default_tab_stop_inches,
    set_source_indent,
)


PERSON_LINE_RE = re.compile(r"^\d+\.\s*(\S+)")
PROGRAM_SECTION_RE = re.compile(r"^節目.*則")
STOP_SECTION_RE = re.compile(r"^(?:-+|FB小編文|本周節日)")
TRANSLATOR_TAG_RE = re.compile(r"\s*[A-Za-z]+/[A-Za-z]+\s*$")
BLOCK_INDEX_RE = re.compile(r"^\d+\s*$")
YYMD_DATE_RE = re.compile(r"^(?P<yy>\d{2})/(?P<mm>\d{1,2})/(?P<dd>\d{1,2})$")
MD_DATE_RE = re.compile(r"^(?P<mm>\d{1,2})/(?P<dd>\d{1,2})$")
CJK_RE = re.compile(r"[\u4e00-\u9fff]")
QUOTE_CHARS = "\"'“”‘’"
ALEX_REF_LABELS = {"參考資料:", "參考資料："}
ALEX_VIDEO_LABELS = {"要用的影片:", "要用的影片："}
PAREN_TITLE_RE = re.compile(r"[\(（]([^()（）]+)[\)）]")
DASH_SPLIT_RE = re.compile(r"\s*-\s*")


def _extract_hyperlink_target(paragraph) -> str | None:
    for hyperlink in paragraph._p.findall(".//w:hyperlink", paragraph._p.nsmap):
        r_id = hyperlink.get(qn("r:id"))
        if not r_id:
            continue
        rel = paragraph.part.rels.get(r_id)
        if rel is None:
            continue
        target = getattr(rel, "target_ref", None)
        if target and isinstance(target, str):
            return target
    return None


def iter_non_empty_paragraphs(doc: Document) -> tuple[list[str], list[str | None]]:
    lines: list[str] = []
    url_targets: list[str | None] = []
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        lines.append(text)
        if text.startswith("http"):
            target = _extract_hyperlink_target(paragraph)
            url_targets.append(target.strip() if target and target.startswith("http") else None)
        else:
            url_targets.append(None)
    return lines, url_targets


def normalize_title(title_line: str) -> str:
    title = title_line.replace(" - ", " ").strip()
    title = TRANSLATOR_TAG_RE.sub("", title).strip()
    title = title.replace("/", "")
    return title


def _clean_title_for_display(title_line: str) -> str:
    return TRANSLATOR_TAG_RE.sub("", title_line).strip()


def _strip_quotes(text: str) -> str:
    return text.translate(str.maketrans("", "", QUOTE_CHARS))


def _is_cjk(text: str) -> bool:
    return bool(CJK_RE.search(text))


def _split_program_title(title_line: str) -> tuple[str, str]:
    cleaned = _clean_title_for_display(title_line)
    if " - " in cleaned:
        program, title = cleaned.split(" - ", 1)
        return program.strip(), title.strip()
    if "-" in cleaned:
        parts = DASH_SPLIT_RE.split(cleaned, 1)
        if len(parts) == 2:
            left, right = (part.strip() for part in parts)
            if left and right and (
                " " in left
                or " " in right
                or _is_cjk(left)
                or _is_cjk(right)
            ):
                return left, right
    return cleaned.strip(), cleaned.strip()


def _strip_trailing_filename_punct(text: str) -> str:
    return text.rstrip(" \t\r\n。．.？?！!：:;；,，")


def _clean_filename_component(text: str) -> str:
    cleaned = TRANSLATOR_TAG_RE.sub("", text).strip()
    cleaned = cleaned.replace("/", "")
    cleaned = cleaned.replace("|", " ")
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return _strip_trailing_filename_punct(cleaned)


def build_filename_title_from_title_line(title_line: str) -> str:
    matches = PAREN_TITLE_RE.findall(title_line)
    cjk_parenthetical: str | None = None
    for inner in reversed(matches):
        if _is_cjk(inner):
            cjk_parenthetical = inner.strip()
            break

    source = cjk_parenthetical or _clean_title_for_display(title_line)
    if " - " in source:
        program, episode = (part.strip() for part in source.split(" - ", 1))
    else:
        program, episode = _split_program_title(source)

    program = _clean_filename_component(program)
    episode = _clean_filename_component(episode)
    if program and episode and program != episode:
        return f"{program} {episode}".strip()
    return (program or episode).strip()


def _strip_parenthetical(text: str) -> str:
    return re.sub(r"[\(（][^()（）]+[\)）]", "", text).strip()


def build_hashtags_from_title_line(title_line: str) -> tuple[str, str]:
    display = _clean_title_for_display(title_line)
    matches = PAREN_TITLE_RE.findall(display)
    cjk_parenthetical: str | None = None
    for inner in reversed(matches):
        if _is_cjk(inner):
            cjk_parenthetical = inner.strip()
            break

    en_source = _strip_parenthetical(display)
    en_program, en_title = _split_program_title(en_source)
    hashtags_en = _build_hashtags(en_program, en_title, pascal_case=True)

    if cjk_parenthetical:
        zh_program, zh_title = _split_program_title(cjk_parenthetical)
    else:
        zh_program, zh_title = _split_program_title(display)
    hashtags_zh = _build_hashtags(zh_program, zh_title, pascal_case=False)
    return hashtags_en, hashtags_zh


def _preferred_filename_title(title_line: str) -> str:
    matches = PAREN_TITLE_RE.findall(title_line)
    for inner in reversed(matches):
        if _is_cjk(inner):
            return inner.strip()
    return title_line.strip()


def _normalize_hashtag(text: str, pascal_case: bool) -> str:
    stripped = _strip_quotes(text)
    if _is_cjk(stripped):
        # Keep CJK and alphanumerics, strip other punctuation/spaces.
        return "".join(ch for ch in stripped if ch.isalnum() or _is_cjk(ch))
    cleaned = "".join(ch if ch.isalnum() else " " for ch in stripped)
    words = [word for word in cleaned.split() if word]
    if pascal_case:
        return "".join(word[:1].upper() + word[1:] for word in words)
    return "".join(words)


def _build_hashtags(program: str, title: str, pascal_case: bool) -> str:
    program_tag = _normalize_hashtag(program, pascal_case=pascal_case)
    title_tag = _normalize_hashtag(title, pascal_case=pascal_case)
    tags = []
    for tag in (program_tag, title_tag):
        if not tag:
            continue
        if tags and tags[-1] == tag:
            continue
        tags.append(tag)
    return " ".join(f"#{tag}" for tag in tags)


def _extract_reference(
    lines: list[str], url_targets: list[str | None], start_idx: int
) -> tuple[str, str, str]:
    for idx in range(start_idx, len(lines)):
        line = lines[idx]
        if PERSON_LINE_RE.match(line) or STOP_SECTION_RE.match(line):
            break
        if line.strip() == "搭配":
            ref_url = lines[idx + 1] if idx + 1 < len(lines) else ""
            ref_url_target = (
                url_targets[idx + 1].strip()
                if idx + 1 < len(url_targets) and url_targets[idx + 1]
                else ""
            )
            ref_title = lines[idx + 2] if idx + 2 < len(lines) else ""
            return ref_url, ref_title, ref_url_target
    return "", "", ""


def _parse_date_prefix(text: str, default_year: int | None = None) -> str | None:
    stripped = text.strip()
    match = YYMD_DATE_RE.match(stripped)
    if match:
        yy = int(match.group("yy"))
        mm = int(match.group("mm"))
        dd = int(match.group("dd"))
        if not (1 <= mm <= 12 and 1 <= dd <= 31):
            return None
        return f"{yy:02d}{mm:02d}{dd:02d}"
    match = MD_DATE_RE.match(stripped)
    if not match or default_year is None:
        return None
    mm = int(match.group("mm"))
    dd = int(match.group("dd"))
    if not (1 <= mm <= 12 and 1 <= dd <= 31):
        return None
    yy = default_year % 100
    return f"{yy:02d}{mm:02d}{dd:02d}"


def _detect_schedule_format(lines: list[str]) -> str:
    if any(PROGRAM_SECTION_RE.match(line) for line in lines):
        return "schedule"
    has_ref = any(line.strip() in ALEX_REF_LABELS for line in lines)
    has_video = any(line.strip() in ALEX_VIDEO_LABELS for line in lines)
    if has_ref and has_video:
        return "blocks"
    return "schedule"


def extract_post_entries_from_blocks(schedule_path: Path) -> list[dict[str, str]]:
    doc = Document(str(schedule_path))
    lines, url_targets = iter_non_empty_paragraphs(doc)
    default_year = date.today().year

    block_starts: list[int] = [
        idx for idx, line in enumerate(lines) if BLOCK_INDEX_RE.match(line.strip())
    ]
    entries: list[dict[str, str]] = []

    for pos, start_idx in enumerate(block_starts):
        end_idx = block_starts[pos + 1] if pos + 1 < len(block_starts) else len(lines)
        if end_idx - start_idx <= 1:
            continue
        block = lines[start_idx:end_idx]
        block_targets = url_targets[start_idx:end_idx]

        ref_url = ""
        ref_title = ""
        ref_url_target = ""
        date_prefix = ""
        video_url = ""
        video_url_target = ""
        video_title = ""
        video_desc_en = ""
        video_desc_zh = ""

        for idx, line in enumerate(block):
            label = line.strip()
            if label in ALEX_REF_LABELS:
                if idx + 1 < len(block):
                    ref_url = block[idx + 1].strip()
                    ref_url_target = (
                        block_targets[idx + 1].strip()
                        if block_targets[idx + 1]
                        else ""
                    )
                end_ref = next(
                    (
                        pos
                        for pos in range(idx + 1, len(block))
                        if block[pos].strip() in ALEX_VIDEO_LABELS
                    ),
                    len(block),
                )
                cursor = idx + 2
                if cursor < end_ref:
                    possible_date = _parse_date_prefix(block[cursor], default_year)
                    if possible_date:
                        date_prefix = possible_date
                        cursor += 1
                if cursor < end_ref:
                    ref_lines = [
                        text.strip()
                        for text in block[cursor:end_ref]
                        if text.strip()
                    ]
                    ref_title = "\n".join(ref_lines)
            elif label in ALEX_VIDEO_LABELS:
                if idx + 1 < len(block):
                    video_url = block[idx + 1].strip()
                    video_url_target = (
                        block_targets[idx + 1].strip()
                        if block_targets[idx + 1]
                        else ""
                    )
                if idx + 2 < len(block):
                    video_title = block[idx + 2].strip()
                if idx + 3 < len(block):
                    video_desc_en = block[idx + 3].strip()
                if idx + 4 < len(block):
                    video_desc_zh = block[idx + 4].strip()

        if not video_title:
            continue

        program_name, episode_title = _split_program_title(video_title)
        hashtags_en, hashtags_zh = build_hashtags_from_title_line(video_title)
        filename_title = build_filename_title_from_title_line(video_title)
        entry: dict[str, str] = {
            "filename_title": filename_title,
            "header_title": _clean_title_for_display(video_title),
            "header_url": video_url,
            "header_url_target": video_url_target,
            "video_url": video_url,
            "video_url_target": video_url_target,
            "video_title": _clean_title_for_display(video_title),
            "video_desc_en": video_desc_en,
            "video_desc_zh": video_desc_zh,
            "ref_url": ref_url,
            "ref_url_target": ref_url_target,
            "ref_title": ref_title,
            "hashtags_en": hashtags_en,
            "hashtags_zh": hashtags_zh,
        }
        if date_prefix:
            entry["filename_prefix_override"] = f"{date_prefix}_"
        entries.append(entry)

    return entries


def extract_post_entries(schedule_path: Path) -> list[dict[str, str]]:
    doc = Document(str(schedule_path))
    lines, url_targets = iter_non_empty_paragraphs(doc)
    schedule_format = _detect_schedule_format(lines)
    if schedule_format == "blocks":
        return extract_post_entries_from_blocks(schedule_path)

    entries: list[dict[str, str]] = []
    in_program_section = False

    for idx, line in enumerate(lines):
        if not in_program_section:
            if PROGRAM_SECTION_RE.match(line):
                in_program_section = True
            continue

        if STOP_SECTION_RE.match(line):
            break

        match = PERSON_LINE_RE.match(line)
        if not match:
            continue

        person = match.group(1).strip().lower()
        if idx + 1 >= len(lines):
            continue
        title_line = lines[idx + 1]
        url_line = lines[idx + 2] if idx + 2 < len(lines) else ""
        url_target = url_targets[idx + 2] if idx + 2 < len(url_targets) else None
        if not url_line.startswith("http"):
            url_line = ""
            url_target = None
        if person == "alex":
            ref_url, ref_title, ref_url_target = _extract_reference(
                lines, url_targets, idx + 1
            )
            program_name, episode_title = _split_program_title(title_line)
            hashtags_en, hashtags_zh = build_hashtags_from_title_line(title_line)
            entries.append(
                {
                    "filename_title": build_filename_title_from_title_line(title_line),
                    "header_title": _clean_title_for_display(title_line),
                    "header_url": url_line,
                    "header_url_target": url_target or "",
                    "video_url": url_line,
                    "video_url_target": url_target or "",
                    "video_title": _clean_title_for_display(title_line),
                    "ref_url": ref_url,
                    "ref_url_target": ref_url_target,
                    "ref_title": ref_title,
                    "hashtags_en": hashtags_en,
                    "hashtags_zh": hashtags_zh,
                }
            )

    return entries


def extract_post_titles(schedule_path: Path) -> list[str]:
    return [entry["filename_title"] for entry in extract_post_entries(schedule_path)]


def replace_in_runs(paragraph, placeholder: str, value: str) -> bool:
    changed = False
    for run in paragraph.runs:
        if placeholder in run.text:
            run.text = run.text.replace(placeholder, value)
            changed = True
    return changed


def apply_source_style(paragraph) -> None:
    apply_highlight_to_runs(paragraph)


def sync_empty_paragraph_indents(doc: Document) -> None:
    last_left = None
    last_first = None
    for paragraph in doc.paragraphs:
        if paragraph.text.strip():
            fmt = paragraph.paragraph_format
            last_left = fmt.left_indent
            last_first = fmt.first_line_indent
            continue
        if last_left is None and last_first is None:
            continue
        paragraph.paragraph_format.left_indent = last_left
        paragraph.paragraph_format.first_line_indent = last_first



def replace_placeholders(
    doc: Document,
    mapping: dict[str, str],
    indent_inches: float,
    hyperlink_targets: dict[str, str] | None = None,
) -> None:
    highlight_keys = {"{{REF_TITLE}}", "{{VIDEO_TITLE}}"}
    hyperlink_keys = {"{{HEADER_URL}}", "{{REF_URL}}", "{{VIDEO_URL}}"}
    plain_hyperlink_keys = {"{{HEADER_URL}}"}
    indent_keys = {
        "{{REF_URL}}",
        "{{REF_TITLE}}",
        "{{REF_SUMMARY_ZH}}",
        "{{REF_TITLE_EN}}",
        "{{REF_SUMMARY_EN}}",
        "{{VIDEO_URL}}",
        "{{VIDEO_TITLE}}",
        "{{VIDEO_DESC_EN}}",
        "{{VIDEO_DESC_ZH}}",
    }
    indent_labels = {"參考資料：", "英文翻譯：", "要用的影片："}

    for paragraph in doc.paragraphs:
        paragraph_text = paragraph.text
        if paragraph_text.strip() in indent_labels:
            set_source_indent(paragraph, indent_inches)
        for placeholder, value in mapping.items():
            if placeholder not in paragraph_text:
                continue
            if placeholder in indent_keys:
                set_source_indent(paragraph, indent_inches)
            if placeholder in hyperlink_keys and value:
                target = value
                if hyperlink_targets:
                    target = hyperlink_targets.get(placeholder) or value
                clear_paragraph(paragraph)
                if placeholder in plain_hyperlink_keys:
                    add_hyperlink(paragraph, value, target, highlight=False)
                else:
                    add_hyperlink(paragraph, value, target, highlight=True)
                paragraph_text = paragraph.text
                continue
            if placeholder in highlight_keys and value:
                clear_paragraph(paragraph)
                add_highlighted_run(paragraph, value)
                paragraph_text = paragraph.text
                continue
            if not replace_in_runs(paragraph, placeholder, value):
                paragraph.text = paragraph.text.replace(placeholder, value)
            if placeholder in indent_keys and paragraph_text.strip() not in indent_labels:
                apply_source_style(paragraph)
            paragraph_text = paragraph.text
        for placeholder in indent_keys:
            if placeholder in paragraph.text:
                set_source_indent(paragraph, indent_inches)
                break


def make_unique_path(base: Path) -> Path:
    if not base.exists():
        return base
    stem = base.stem
    suffix = base.suffix
    counter = 2
    while True:
        candidate = base.with_name(f"{stem}_{counter}{suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def generate_docs(
    schedule_path: Path,
    template_path: Path,
    output_dir: Path,
    filename_prefix: str,
    filename_suffix: str,
) -> list[Path]:
    entries = extract_post_entries(schedule_path)
    output_paths: list[Path] = []
    for entry in entries:
        entry_prefix = entry.get("filename_prefix_override", filename_prefix)
        filename = f"{entry_prefix}{entry['filename_title']}小編文{filename_suffix}.docx"
        output_path = make_unique_path(output_dir / filename)
        doc = Document(str(template_path))
        default_tab_stop = get_default_tab_stop_inches(doc)
        mapping = {
            "{{HEADER_TITLE}}": entry["header_title"],
            "{{HEADER_URL}}": entry["header_url"],
            "{{REF_URL}}": entry["ref_url"],
            "{{REF_TITLE}}": entry["ref_title"],
            "{{VIDEO_URL}}": entry["video_url"],
            "{{VIDEO_TITLE}}": entry["video_title"],
            "{{HASHTAGS_EN}}": entry["hashtags_en"],
            "{{HASHTAGS_ZH}}": entry["hashtags_zh"],
        }
        hyperlink_targets = {
            "{{HEADER_URL}}": entry.get("header_url_target", ""),
            "{{REF_URL}}": entry.get("ref_url_target", ""),
            "{{VIDEO_URL}}": entry.get("video_url_target", ""),
        }
        video_desc_en = entry.get("video_desc_en", "").strip()
        if video_desc_en:
            mapping["{{VIDEO_DESC_EN}}"] = video_desc_en
        video_desc_zh = entry.get("video_desc_zh", "").strip()
        if video_desc_zh:
            mapping["{{VIDEO_DESC_ZH}}"] = video_desc_zh
        replace_placeholders(
            doc,
            mapping,
            default_tab_stop,
            hyperlink_targets=hyperlink_targets,
        )
        ensure_blank_after_labels(doc, {"參考資料：", "英文翻譯：", "要用的影片："})
        sync_empty_paragraph_indents(doc)
        doc.save(str(output_path))
        output_paths.append(output_path)
    return output_paths


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generate empty post docs from alex entries in a schedule DOCX."
    )
    parser.add_argument(
        "--schedule",
        default="260302排程_ev_k.docx",
        help="Path to the schedule DOCX.",
    )
    parser.add_argument(
        "--template",
        default="templates/post_template.docx",
        help="Path to the base DOCX template with shared styles.",
    )
    parser.add_argument(
        "--output-dir",
        default="outputs",
        help="Directory to write generated DOCX files.",
    )
    parser.add_argument(
        "--prefix",
        default="日期未定_",
        help="Filename prefix.",
    )
    parser.add_argument(
        "--suffix",
        default="_al",
        help="Filename suffix (without extension).",
    )
    args = parser.parse_args()

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    generate_docs(
        schedule_path=Path(args.schedule),
        template_path=Path(args.template),
        output_dir=output_dir,
        filename_prefix=args.prefix,
        filename_suffix=args.suffix,
    )


if __name__ == "__main__":
    main()
import string
