#!/usr/bin/env python3

from __future__ import annotations

import argparse
import html
import json
import re
from datetime import date
from pathlib import Path
from urllib.request import Request, urlopen

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from docx_utils import (
    add_highlighted_run,
    add_hyperlink,
    apply_font_size_to_document_runs,
    apply_font_size_to_runs,
    apply_highlight_to_runs,
    clear_paragraph,
    ensure_blank_after_labels,
    get_default_tab_stop_inches,
    set_source_indent,
)
from style_tokens import BODY_TEXT_SIZE_PT, REFERENCE_HIGHLIGHT_DEFAULT, REFERENCE_TEXT_SIZE_PT


PERSON_LINE_RE = re.compile(r"^\d+\.\s*(\S+)")
DATE_TASK_LINE_RE = re.compile(
    r"^(?:\d+\.\s*)?\d{1,2}/\d{1,2}(?:\([^()]*\))?\s*發\s*(\S+)"
)
DATE_TASK_PREFIX_RE = re.compile(
    r"^(?:\d+\.\s*)?(?P<date>\d{1,2}/\d{1,2}|\d{2}/\d{1,2}/\d{1,2})(?:\([^()]*\))?\s*發\s*\S+"
)
DATE_TASK_DISPLAY_RE = re.compile(
    r"^(?:\d+\.\s*)?(?P<display>\d{1,2}/\d{1,2}(?:\([^()]*\))?)\s*發\s*\S+"
)
BARE_PERSON_LINE_RE = re.compile(r"^[A-Za-z][A-Za-z0-9._-]*$")
PROGRAM_SECTION_RE = re.compile(r"^節目.*則")
BODHI_SECTION_RE = re.compile(r"^(?:人間)?菩提.*則")
STOP_SECTION_RE = re.compile(r"^(?:-+|FB小編文|本周節日)")
TRANSLATOR_TAG_RE = re.compile(r"\s*[A-Za-z]+/[A-Za-z]+\s*$")
BLOCK_INDEX_RE = re.compile(r"^\d+\s*$")
YYMD_DATE_RE = re.compile(r"^(?P<yy>\d{2})/(?P<mm>\d{1,2})/(?P<dd>\d{1,2})$")
MD_DATE_RE = re.compile(r"^(?P<mm>\d{1,2})/(?P<dd>\d{1,2})$")
BODHI_DATE_PREFIX_RE = re.compile(r"^(?P<mm>\d{1,2})/(?P<dd>\d{1,2})")
CJK_RE = re.compile(r"[\u4e00-\u9fff]")
QUOTE_CHARS = "\"'“”‘’"
ALEX_REF_LABELS = {"參考資料:", "參考資料："}
ALEX_VIDEO_LABELS = {"要用的影片:", "要用的影片："}
PAIRING_LABELS = {"搭配", "搭配:", "搭配："}
PAREN_TITLE_RE = re.compile(r"[\(（]([^()（）]+)[\)）]")
DASH_SPLIT_RE = re.compile(r"\s*-\s*")
TAG_RE = re.compile(r"<[^>]+>")
MULTISPACE_RE = re.compile(r"\s+")
BODHI_TIMELINE_RE = re.compile(r"^\d{2}:\d{2}\s*[│|｜]")
EPISODE_JSON_RE = re.compile(
    r"""\b(?:episodeJson|episdoeJson)\s*=\s*(?:"(?P<double>(?:\\.|[^"\\])*)"|'(?P<single>(?:\\.|[^'\\])*)')""",
    re.DOTALL,
)


def _extract_person_name(line: str) -> str | None:
    for pattern in (DATE_TASK_LINE_RE, PERSON_LINE_RE):
        match = pattern.match(line)
        if match:
            return match.group(1).strip().lower()
    stripped = line.strip()
    if BARE_PERSON_LINE_RE.fullmatch(stripped):
        return stripped.lower()
    return None


def _extract_task_date_prefix(line: str, default_year: int) -> str | None:
    match = DATE_TASK_PREFIX_RE.match(line.strip())
    if not match:
        return None
    return _parse_date_prefix(match.group("date"), default_year=default_year)


def _extract_task_date_display(line: str) -> str | None:
    match = DATE_TASK_DISPLAY_RE.match(line.strip())
    if not match:
        return None
    return match.group("display").strip()


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
        target = _extract_hyperlink_target(paragraph)
        target = target.strip() if target and target.startswith("http") else None
        for raw_line in paragraph.text.splitlines():
            text = raw_line.strip()
            if not text:
                continue
            lines.append(text)
            if text.startswith("http"):
                url_targets.append(target)
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
    if " - " in cleaned and _is_cjk(cleaned):
        parts = [part.strip() for part in cleaned.split(" - ") if part.strip()]
        if len(parts) >= 2:
            # For CJK titles, a trailing 3rd segment is often a speaker name.
            return parts[0], parts[1]
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


def _split_cjk_parenthetical_title(title_line: str) -> tuple[str, str]:
    parts = [part.strip() for part in title_line.split(" - ") if part.strip()]
    if len(parts) >= 2:
        # Some inputs append a trailing speaker name segment.
        return parts[0], parts[1]
    return _split_program_title(title_line)


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
    if cjk_parenthetical:
        program, episode = _split_cjk_parenthetical_title(source)
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
        zh_program, zh_title = _split_cjk_parenthetical_title(cjk_parenthetical)
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


def _normalize_bodhi_english_title_tag(title: str) -> str:
    # Keep possessive "s" while stripping apostrophe punctuation.
    return _normalize_hashtag(title.replace("’", "").replace("'", ""), pascal_case=True)


def _build_bodhi_hashtags(cleaned_title: str, english_title: str) -> tuple[str, str]:
    english_title_tag = _normalize_bodhi_english_title_tag(english_title) if english_title else ""
    hashtags_en_parts = ["#LifeWisdom"]
    if english_title_tag:
        hashtags_en_parts.append(f"#{english_title_tag}")
    hashtags_en_parts.extend(["#VenerableMasterChengYen", "#TzuChi"])

    chinese_title_tag = _normalize_hashtag(cleaned_title, pascal_case=False)
    hashtags_zh_parts = ["#人間菩提"]
    if chinese_title_tag:
        hashtags_zh_parts.append(f"#{chinese_title_tag}")
    hashtags_zh_parts.extend(["#證嚴上人", "#慈濟"])
    return " ".join(hashtags_en_parts), " ".join(hashtags_zh_parts)


def _extract_ref_from_label(
    lines: list[str],
    url_targets: list[str | None],
    label_idx: int,
    *,
    stop_labels: set[str],
    stop_by_person: bool,
    parse_first_line_date: bool,
    default_year: int,
) -> tuple[str, str, str, str]:
    ref_url = lines[label_idx + 1].strip() if label_idx + 1 < len(lines) else ""
    ref_url_target = (
        url_targets[label_idx + 1].strip()
        if label_idx + 1 < len(url_targets) and url_targets[label_idx + 1]
        else ""
    )

    ref_lines: list[str] = []
    date_prefix = ""
    cursor = label_idx + 2
    while cursor < len(lines):
        candidate = lines[cursor]
        stripped = candidate.strip()
        if stop_by_person and _extract_person_name(candidate):
            break
        if STOP_SECTION_RE.match(candidate):
            break
        if stripped in stop_labels:
            break
        if stripped:
            if parse_first_line_date and not date_prefix:
                possible = _parse_date_prefix(stripped, default_year=default_year)
                if possible:
                    date_prefix = possible
                    cursor += 1
                    continue
            ref_lines.append(stripped)
        cursor += 1

    return ref_url, "\n".join(ref_lines), ref_url_target, date_prefix


def _extract_schedule_reference(
    lines: list[str], url_targets: list[str | None], start_idx: int
) -> tuple[str, str, str, str, str, str]:
    for idx in range(start_idx, len(lines)):
        line = lines[idx]
        if _extract_person_name(line) or STOP_SECTION_RE.match(line):
            break
        if line.strip() in PAIRING_LABELS:
            ref_url, ref_title, ref_url_target, _ = _extract_ref_from_label(
                lines,
                url_targets,
                idx,
                stop_labels=PAIRING_LABELS,
                stop_by_person=True,
                parse_first_line_date=False,
                default_year=date.today().year,
            )
            ref_lines = [part.strip() for part in ref_title.splitlines() if part.strip()]
            ref_title = ref_lines[0] if ref_lines else ""
            ref_summary_zh = ""
            ref_title_en = ""
            ref_summary_en = ""
            if len(ref_lines) >= 2:
                ref_summary_en = ref_lines[1]
            if len(ref_lines) >= 3:
                ref_summary_zh = ref_lines[2]
            return (
                ref_url,
                ref_title,
                ref_url_target,
                ref_summary_zh,
                ref_title_en,
                ref_summary_en,
            )
    return "", "", "", "", "", ""


def _build_standard_entry(
    *,
    video_title: str,
    video_url: str,
    video_url_target: str,
    ref_url: str,
    ref_url_target: str,
    ref_title: str,
    video_desc_en: str = "",
    video_desc_zh: str = "",
) -> dict[str, str]:
    hashtags_en, hashtags_zh = build_hashtags_from_title_line(video_title)
    return {
        "filename_title": build_filename_title_from_title_line(video_title),
        "header_title": _clean_title_for_display(video_title),
        "header_url": video_url,
        "header_url_target": video_url_target,
        "video_url": video_url,
        "video_url_target": video_url_target,
        "video_title": _clean_title_for_display(video_title),
        "source_video_title": _clean_title_for_display(video_title),
        "video_desc_en": video_desc_en,
        "video_desc_zh": video_desc_zh,
        "ref_url": ref_url,
        "ref_url_target": ref_url_target,
        "ref_title": ref_title,
        "hashtags_en": hashtags_en,
        "hashtags_zh": hashtags_zh,
    }


def _build_bodhi_entry(
    *,
    title_line: str,
    url_line: str,
    url_target: str,
    default_year: int,
    explicit_english_title: str = "",
) -> dict[str, str]:
    raw_title = title_line.strip()
    cleaned_title, date_prefix = _extract_bodhi_date_prefix(raw_title)
    english_title = explicit_english_title.strip() or fetch_bodhi_english_subtitle(
        url_line, cleaned_title
    )
    ref_excerpt = fetch_bodhi_reference_excerpt(url_line, cleaned_title)
    display_title = cleaned_title
    if english_title:
        display_title = f"{display_title}\n{english_title}"
    if ref_excerpt:
        display_title = f"{display_title}\n\n{ref_excerpt}"

    header_title_lines = ["人間菩提", raw_title]
    if english_title:
        header_title_lines.append(english_title)
    header_title = "\n".join(header_title_lines)
    hashtags_en, hashtags_zh = _build_bodhi_hashtags(cleaned_title, english_title)
    entry = {
        "filename_title": "人間菩提",
        "header_title": _clean_title_for_display(header_title),
        "header_url": url_line,
        "header_url_target": url_target,
        "video_url": url_line,
        "video_url_target": url_target,
        "video_title": _clean_title_for_display(header_title),
        "ref_url": url_line,
        "ref_url_target": url_target,
        "ref_title": _clean_title_for_display(display_title),
        "hashtags_en": hashtags_en,
        "hashtags_zh": hashtags_zh,
        "reference_only": "true",
    }
    if date_prefix:
        parsed_prefix = _parse_date_prefix(date_prefix, default_year=default_year)
        if parsed_prefix:
            entry["filename_prefix_override"] = f"{parsed_prefix}_"
    return entry


def _build_standard_schedule_entry(
    *,
    line: str,
    title_line: str,
    url_line: str,
    url_target: str,
    lines: list[str],
    url_targets: list[str | None],
    start_idx: int,
    default_year: int,
) -> dict[str, str]:
    (
        ref_url,
        ref_title,
        ref_url_target,
        ref_summary_zh,
        ref_title_en,
        ref_summary_en,
    ) = _extract_schedule_reference(
        lines, url_targets, start_idx
    )
    entry = _build_standard_entry(
        video_title=title_line,
        video_url=url_line,
        video_url_target=url_target,
        ref_url=ref_url,
        ref_url_target=ref_url_target,
        ref_title=ref_title,
        video_desc_en=ref_summary_en,
        video_desc_zh=ref_summary_zh,
    )
    if ref_title:
        entry["source_video_title"] = ref_title
    task_prefix = _extract_task_date_prefix(line, default_year=default_year)
    task_display = _extract_task_date_display(line)
    if task_prefix:
        entry["filename_prefix_override"] = f"{task_prefix}_"
    if task_display:
        entry["header_title"] = f"{task_display}\n{entry['header_title']}"
    return entry


def _resolve_bodhi_title_and_url(
    lines: list[str],
    url_targets: list[str | None],
    idx: int,
) -> tuple[str, str, str, str]:
    title_line = lines[idx + 1] if idx + 1 < len(lines) else ""
    explicit_english_title = ""
    url_line = ""
    url_target: str | None = None

    candidate1 = lines[idx + 2] if idx + 2 < len(lines) else ""
    candidate2 = lines[idx + 3] if idx + 3 < len(lines) else ""

    if candidate1.startswith("http"):
        url_line = candidate1
        url_target = url_targets[idx + 2] if idx + 2 < len(url_targets) else None
    elif candidate2.startswith("http"):
        explicit_english_title = candidate1.strip()
        url_line = candidate2
        url_target = url_targets[idx + 3] if idx + 3 < len(url_targets) else None

    if not url_line:
        url_line = ""
        url_target = None

    return title_line, explicit_english_title, url_line, (url_target or "")


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


def _extract_bodhi_date_prefix(title_line: str) -> tuple[str, str | None]:
    stripped = title_line.strip()
    match = BODHI_DATE_PREFIX_RE.match(stripped) if stripped else None
    if not match:
        return stripped, None
    mm = int(match.group("mm"))
    dd = int(match.group("dd"))
    if not (1 <= mm <= 12 and 1 <= dd <= 31):
        return stripped, None
    rest = stripped[len(match.group(0)) :].lstrip()
    if rest.startswith("首播"):
        rest = rest[len("首播") :].lstrip()
    cleaned_title = rest.strip() or stripped
    return cleaned_title, f"{mm:02d}/{dd:02d}"


def _looks_like_english_title(text: str) -> bool:
    candidate = re.sub(r"\s+", " ", text).strip()
    if len(candidate) < 8:
        return False
    if _is_cjk(candidate):
        return False
    if not re.search(r"[A-Za-z]", candidate):
        return False
    if candidate.startswith("http"):
        return False
    return True


def _collect_english_title_lines(lines: list[str], start_idx: int) -> str:
    collected: list[str] = []
    for line in lines[start_idx:]:
        if not _looks_like_english_title(line):
            break
        if line.startswith("#") or _is_bodhi_copyright_line(line):
            break
        collected.append(_normalize_line_for_match(line))
    return " ".join(collected).strip()


def fetch_bodhi_english_subtitle(url: str, chinese_title: str) -> str:
    if not url or not chinese_title:
        return ""
    try:
        req = Request(
            url,
            headers={
                "User-Agent": "Mozilla/5.0",
                "Accept-Language": "en-US,en;q=0.9,zh-TW;q=0.8,zh;q=0.7",
            },
        )
        with urlopen(req, timeout=8) as resp:
            raw = resp.read()
    except Exception:
        return ""

    page = raw.decode("utf-8", errors="ignore")
    if not page:
        return ""

    for match in re.finditer(re.escape(chinese_title), page):
        snippet = page[match.end() : match.end() + 1500]
        text = TAG_RE.sub("\n", snippet)
        text = html.unescape(text)
        lines = [_normalize_line_for_match(segment) for segment in text.splitlines()]
        lines = [line for line in lines if line]
        for idx, line in enumerate(lines):
            if _looks_like_english_title(line):
                return _collect_english_title_lines(lines, idx)
    return ""


def _normalize_line_for_match(text: str) -> str:
    return MULTISPACE_RE.sub(" ", text).strip()


def _is_bodhi_copyright_line(line: str) -> bool:
    return "All rights reserved" in line or "版權註記" in line or "版權所有" in line


def _is_bodhi_excerpt_boundary(line: str) -> bool:
    stripped = line.strip()
    return (
        not stripped
        or stripped.startswith("#")
        or stripped.startswith("http")
        or stripped.startswith("---")
        or BODHI_TIMELINE_RE.match(stripped) is not None
        or _is_bodhi_copyright_line(stripped)
    )


def _parse_episode_json_payload(payload: str) -> dict | None:
    normalized = html.unescape(payload).strip()
    for _ in range(4):
        try:
            parsed = json.loads(normalized)
            if isinstance(parsed, dict):
                return parsed
            if isinstance(parsed, str):
                normalized = parsed
                continue
        except json.JSONDecodeError:
            pass

        unescaped = normalized.replace('\\"', '"').replace("\\/", "/")
        if unescaped == normalized:
            break
        normalized = unescaped
    return None


def _iter_episode_json_objects(page: str):
    for match in EPISODE_JSON_RE.finditer(page):
        parsed = _parse_episode_json_payload(
            match.group("single") or match.group("double") or ""
        )
        if isinstance(parsed, dict):
            yield parsed


def _excerpt_from_bodhi_description(description: str, chinese_title: str = "") -> str:
    raw_lines = description.splitlines()
    lines = [_normalize_line_for_match(html.unescape(line)) for line in raw_lines]

    start_search_idx = 0
    title_norm = _normalize_line_for_match(chinese_title)
    if title_norm:
        for idx, line in enumerate(lines):
            if line == title_norm:
                start_search_idx = idx + 1
                break

    start_idx = -1
    for idx in range(start_search_idx, len(lines)):
        line = lines[idx]
        if not line or line.startswith("#") or _is_bodhi_copyright_line(line):
            continue
        if len(line) >= 12 and _is_cjk(line) and "。" in line:
            start_idx = idx
            break
    if start_idx < 0:
        return ""

    collected: list[str] = []
    for line in lines[start_idx:]:
        if line.startswith("#") or _is_bodhi_copyright_line(line):
            break
        collected.append(line)

    while collected and not collected[-1]:
        collected.pop()
    return "\n".join(collected).strip()


def fetch_bodhi_reference_excerpt(url: str, chinese_title: str) -> str:
    if not url or not chinese_title:
        return ""
    try:
        req = Request(
            url,
            headers={
                "User-Agent": "Mozilla/5.0",
                "Accept-Language": "zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7",
            },
        )
        with urlopen(req, timeout=8) as resp:
            raw = resp.read()
    except Exception:
        return ""

    page = raw.decode("utf-8", errors="ignore")
    if not page:
        return ""

    title_norm = _normalize_line_for_match(chinese_title)
    for episode in _iter_episode_json_objects(page):
        ep_title = _normalize_line_for_match(str(episode.get("EpTitle", "")))
        if ep_title != title_norm:
            continue
        excerpt = _excerpt_from_bodhi_description(
            str(episode.get("Description", "")),
            chinese_title,
        )
        if excerpt:
            return excerpt

    text = html.unescape(TAG_RE.sub("\n", page))
    lines = [_normalize_line_for_match(line) for line in text.splitlines()]
    lines = [line for line in lines if line]

    title_idx = -1
    for idx, line in enumerate(lines):
        if line == title_norm:
            title_idx = idx
            break
    if title_idx < 0:
        return ""

    start_idx = -1
    for idx in range(title_idx + 1, len(lines)):
        line = lines[idx]
        if line.startswith("#"):
            break
        if _is_bodhi_copyright_line(line):
            continue
        if len(line) >= 12 and _is_cjk(line) and "。" in line:
            start_idx = idx
            break
    if start_idx < 0:
        return ""

    collected: list[str] = []
    for idx in range(start_idx, len(lines)):
        line = lines[idx]
        if _is_bodhi_excerpt_boundary(line):
            break
        collected.append(line)
    if not collected:
        return ""
    return "\n".join(collected).strip()


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
                ref_url, ref_title, ref_url_target, date_prefix = _extract_ref_from_label(
                    block,
                    block_targets,
                    idx,
                    stop_labels=ALEX_VIDEO_LABELS,
                    stop_by_person=False,
                    parse_first_line_date=True,
                    default_year=default_year,
                )
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

        entry = _build_standard_entry(
            video_title=video_title,
            video_url=video_url,
            video_url_target=video_url_target,
            ref_url=ref_url,
            ref_url_target=ref_url_target,
            ref_title=ref_title,
            video_desc_en=video_desc_en,
            video_desc_zh=video_desc_zh,
        )
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

    default_year = date.today().year
    entries: list[dict[str, str]] = []
    in_program_section = False
    in_bodhi_section = False

    for idx, line in enumerate(lines):
        if PROGRAM_SECTION_RE.match(line):
            in_program_section = True
            in_bodhi_section = False
            continue
        if BODHI_SECTION_RE.match(line):
            in_program_section = True
            in_bodhi_section = True
            continue
        if not in_program_section:
            continue

        if STOP_SECTION_RE.match(line):
            break

        person = _extract_person_name(line)
        if not person:
            continue
        if idx + 1 >= len(lines):
            continue
        title_line = lines[idx + 1]
        url_line = lines[idx + 2] if idx + 2 < len(lines) else ""
        url_target = url_targets[idx + 2] if idx + 2 < len(url_targets) else None
        explicit_english_title = ""
        if in_bodhi_section:
            (
                title_line,
                explicit_english_title,
                url_line,
                resolved_target,
            ) = _resolve_bodhi_title_and_url(lines, url_targets, idx)
            url_target = resolved_target or None
        elif not url_line.startswith("http"):
            url_line = ""
            url_target = None
        if person == "alex":
            if in_bodhi_section:
                entry = _build_bodhi_entry(
                    title_line=title_line,
                    url_line=url_line,
                    url_target=url_target or "",
                    default_year=default_year,
                    explicit_english_title=explicit_english_title,
                )
                entries.append(entry)
                continue
            entry = _build_standard_schedule_entry(
                line=line,
                title_line=title_line,
                url_line=url_line,
                url_target=url_target or "",
                lines=lines,
                url_targets=url_targets,
                start_idx=idx + 1,
                default_year=default_year,
            )
            entries.append(entry)

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
    apply_font_size_to_runs(paragraph, font_size_pt=REFERENCE_TEXT_SIZE_PT)
    apply_highlight_to_runs(paragraph, highlight_color=REFERENCE_HIGHLIGHT_DEFAULT)


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


def remove_paragraph(paragraph) -> None:
    element = paragraph._element
    element.getparent().remove(element)


def insert_paragraph_after(paragraph, text: str = "") -> Paragraph:
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    if text:
        new_para.add_run(text)
    return new_para


def insert_video_section_spacing(
    doc: Document,
    *,
    video_title: str,
    video_desc_en: str,
    video_desc_zh: str,
    indent_inches: float,
) -> None:
    if not video_title:
        return

    in_video_section = False
    non_empty_in_section: list[Paragraph] = []
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if text == "要用的影片：":
            in_video_section = True
            non_empty_in_section = []
            continue
        if not in_video_section:
            continue
        if text in {"參考資料：", "英文翻譯："}:
            break
        if not text:
            continue
        non_empty_in_section.append(paragraph)

    if len(non_empty_in_section) < 2:
        return

    title_paragraph = non_empty_in_section[1]
    if title_paragraph.text.strip() != video_title.strip():
        return

    current = title_paragraph
    if video_desc_en.strip():
        blank = insert_paragraph_after(current, "")
        set_source_indent(blank, indent_inches)
        current = blank._p.getnext()
        if current is not None and current.tag == qn("w:p"):
            current_para = Paragraph(current, blank._parent)
            if current_para.text.strip() == video_desc_en.strip() and video_desc_zh.strip():
                blank_after_en = insert_paragraph_after(current_para, "")
                set_source_indent(blank_after_en, indent_inches)


def strip_reference_block(doc: Document, ref_url: str, *, keep_ref_title: bool = False) -> None:
    ref_label = "參考資料："
    video_label = "要用的影片："
    ref_url_idx = None
    start_idx = None
    end_idx = None
    for idx, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text.strip()
        if text == ref_label and start_idx is None:
            start_idx = idx
        if ref_url and text == ref_url and ref_url_idx is None and start_idx is not None:
            ref_url_idx = idx
        if text == video_label and end_idx is None:
            end_idx = idx
            break
    if start_idx is None:
        start_idx = ref_url_idx
    if start_idx is None:
        return
    if end_idx is None:
        end_idx = len(doc.paragraphs)

    remove_indices = []
    preserved_ref_title = False
    for idx in range(start_idx, end_idx):
        if doc.paragraphs[idx].text.strip() == ref_label:
            continue
        if idx == ref_url_idx:
            continue
        if keep_ref_title and ref_url_idx is not None and idx > ref_url_idx and not preserved_ref_title:
            if doc.paragraphs[idx].text.strip():
                preserved_ref_title = True
                continue
        remove_indices.append(idx)
    for idx in reversed(remove_indices):
        remove_paragraph(doc.paragraphs[idx])


def strip_bodhi_video_labels(doc: Document) -> None:
    remove_indices = []
    for idx, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text.strip()
        if text == "要用的影片：":
            remove_indices.append(idx)
        if text in {"{{VIDEO_DESC_EN}}", "{{VIDEO_DESC_ZH}}"}:
            remove_indices.append(idx)
    for idx in reversed(remove_indices):
        remove_paragraph(doc.paragraphs[idx])


def strip_bodhi_title_block(doc: Document) -> None:
    remove_indices = []
    for idx, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text.strip()
        if text == "標題":
            remove_indices.append(idx)
        if text in {"{{TITLE_LINE_1}}", "{{TITLE_LINE_2}}"}:
            remove_indices.append(idx)
    for idx in reversed(remove_indices):
        remove_paragraph(doc.paragraphs[idx])


def normalize_empty_paragraphs(doc: Document) -> None:
    prev_empty = False
    for paragraph in list(doc.paragraphs):
        if paragraph.text.strip():
            prev_empty = False
            continue
        fmt = paragraph.paragraph_format
        fmt.left_indent = 0
        fmt.first_line_indent = 0
        fmt.right_indent = 0
        if prev_empty:
            remove_paragraph(paragraph)
        else:
            prev_empty = True


def inject_bodhi_english_under_chinese_title(
    doc: Document, chinese_title: str, english_title: str
) -> None:
    if not chinese_title or not english_title:
        return
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        if text.startswith("#"):
            continue
        if chinese_title not in text:
            continue
        if english_title in text:
            continue
        paragraph.text = f"{paragraph.text}\n{english_title}"


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
            apply_source_style(paragraph)
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
                    apply_font_size_to_runs(
                        paragraph, font_size_pt=BODY_TEXT_SIZE_PT
                    )
                else:
                    add_hyperlink(paragraph, value, target, highlight=True)
                paragraph_text = paragraph.text
                continue
            if placeholder in highlight_keys and value:
                clear_paragraph(paragraph)
                add_highlighted_run(
                    paragraph,
                    value,
                    font_size_pt=REFERENCE_TEXT_SIZE_PT,
                    highlight_color=REFERENCE_HIGHLIGHT_DEFAULT,
                )
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
        apply_font_size_to_document_runs(doc, font_size_pt=BODY_TEXT_SIZE_PT)
        default_tab_stop = get_default_tab_stop_inches(doc)
        mapping = {
            "{{HEADER_TITLE}}": entry["header_title"],
            "{{HEADER_URL}}": entry["header_url"],
            "{{REF_URL}}": entry["ref_url"],
            "{{REF_TITLE}}": entry["ref_title"],
            "{{REF_SUMMARY_ZH}}": "",
            "{{REF_TITLE_EN}}": "",
            "{{REF_SUMMARY_EN}}": "",
            "{{VIDEO_URL}}": entry["video_url"],
            "{{VIDEO_TITLE}}": entry.get("source_video_title", entry["video_title"]),
            "{{HASHTAGS_EN}}": entry["hashtags_en"],
            "{{HASHTAGS_ZH}}": entry["hashtags_zh"],
        }
        if entry.get("reference_only"):
            mapping.update(
                {
                    "{{REF_SUMMARY_ZH}}": "",
                    "{{REF_TITLE_EN}}": "",
                    "{{REF_SUMMARY_EN}}": "",
                    "{{VIDEO_URL}}": "",
                    "{{VIDEO_TITLE}}": "",
                    "{{VIDEO_DESC_EN}}": "",
                    "{{VIDEO_DESC_ZH}}": "",
                }
            )
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
        if entry.get("reference_only"):
            title_lines = entry.get("ref_title", "").splitlines()
            if len(title_lines) >= 2:
                chinese_title = title_lines[0].strip()
                english_title = title_lines[1].strip()
                inject_bodhi_english_under_chinese_title(
                    doc, chinese_title, english_title
                )
        if entry.get("reference_only"):
            strip_reference_block(doc, entry.get("ref_url", ""), keep_ref_title=True)
            strip_bodhi_video_labels(doc)
            strip_bodhi_title_block(doc)
            ensure_blank_after_labels(doc, {"參考資料："})
            normalize_empty_paragraphs(doc)
        else:
            insert_video_section_spacing(
                doc,
                video_title=entry.get("source_video_title", entry["video_title"]),
                video_desc_en=entry.get("video_desc_en", ""),
                video_desc_zh=entry.get("video_desc_zh", ""),
                indent_inches=default_tab_stop,
            )
            ensure_blank_after_labels(doc, {"參考資料：", "英文翻譯：", "要用的影片："})
            normalize_empty_paragraphs(doc)
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
        default="output",
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
