#!/usr/bin/env python3

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from urllib.request import Request, urlopen


ITEM_RE = re.compile(r'<div class="item"\s+id="episode-(?P<idx>\d+)">(?P<body>.*?)<script>', re.S)
TITLE_RE = re.compile(r'<div class="title">(?P<title>.*?)</div>', re.S)
DATE_RE = re.compile(r'<div class="date">(?P<date>\d{4}-\d{2}-\d{2})</div>')

EPID_RE = re.compile(r'\\"EpID\\":\\"(?P<epid>.*?)\\"')
PREMIERE_RE = re.compile(r'\\"EpPremiere\\":\\"(?P<premiere>.*?)\\"')
YTID_RE = re.compile(r'\\"YTID\\":\\"(?P<ytid>.*?)\\"')
EPTITLE_RE = re.compile(r'\\"EpTitle\\":\\"(?P<title>.*?)\\"')
YT_OG_TITLE_RE = re.compile(r'<meta\s+property="og:title"\s+content="([^"]*)"', re.I)
YT_SHORT_DESC_RE = re.compile(r'"shortDescription":"(.*?)"', re.S)
YT_TIMESTAMP_LINE_RE = re.compile(r"^\s*(\d{1,2}:\d{2})｜.+$")


def _clean_html_text(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def _decode_js_escaped_text(text: str) -> str:
    if "\\u" not in text and "\\/" not in text:
        return text
    try:
        return bytes(text, "utf-8").decode("unicode_escape")
    except Exception:
        return text


def _parse_episode_payload(payload: str) -> dict[str, str]:
    out: dict[str, str] = {}
    try:
        data = json.loads(payload.replace('\\"', '"'))
        out["epid"] = str(data.get("EpID", "")).strip()
        out["date"] = str(data.get("EpPremiere", "")).split(" ")[0].strip()
        out["title"] = str(data.get("EpTitle", "")).strip()
        out["ytid"] = str(data.get("YTID", "")).strip()
        return out
    except Exception:
        pass

    epid = EPID_RE.search(payload)
    premiere = PREMIERE_RE.search(payload)
    ytid = YTID_RE.search(payload)
    title = EPTITLE_RE.search(payload)
    if epid:
        out["epid"] = epid.group("epid").strip()
    if premiere:
        out["date"] = premiere.group("premiere").split(" ")[0].strip()
    if ytid:
        out["ytid"] = ytid.group("ytid").strip()
    if title:
        out["title"] = _decode_js_escaped_text(title.group("title").strip())
    return out


def extract_episode_rows_from_html(html: str) -> list[dict[str, str | int]]:
    rows_by_idx: dict[int, dict[str, str | int]] = {}

    for m in ITEM_RE.finditer(html):
        idx = int(m.group("idx"))
        body = m.group("body")
        title_match = TITLE_RE.search(body)
        date_match = DATE_RE.search(body)
        rows_by_idx[idx] = {
            "episode_index": idx,
            "epid": "",
            "date": date_match.group("date").strip() if date_match else "",
            "title": _clean_html_text(title_match.group("title")) if title_match else "",
            "ytid": "",
        }

    script_key = "document.getElementById('episode-"
    pos = 0
    while True:
        start = html.find(script_key, pos)
        if start == -1:
            break
        idx_start = start + len(script_key)
        idx_end = html.find("')", idx_start)
        if idx_end == -1:
            break
        idx_text = html[idx_start:idx_end]
        if not idx_text.isdigit():
            pos = idx_end + 2
            continue
        idx = int(idx_text)
        json_key = "var episodeJson = '"
        json_pos = html.find(json_key, idx_end)
        if json_pos == -1:
            pos = idx_end + 2
            continue
        payload_start = json_pos + len(json_key)
        end_marker = "openEpisodeModal("
        payload_end = html.find(end_marker, payload_start)
        if payload_end == -1:
            end_marker = "});"
            payload_end = html.find(end_marker, payload_start)
            if payload_end == -1:
                pos = payload_start
                continue
        payload = html[payload_start:payload_end]
        payload = payload.strip()
        if payload.endswith(";"):
            payload = payload[:-1].rstrip()
        if payload.endswith("'"):
            payload = payload[:-1].rstrip()
        parsed = _parse_episode_payload(payload)
        existing = rows_by_idx.get(
            idx,
            {"episode_index": idx, "epid": "", "date": "", "title": "", "ytid": ""},
        )
        for key in ("epid", "date", "title", "ytid"):
            value = parsed.get(key, "")
            if value:
                existing[key] = value
        rows_by_idx[idx] = existing
        pos = payload_end + len(end_marker)

    return [rows_by_idx[idx] for idx in sorted(rows_by_idx)]


def fetch_html(url: str) -> str:
    req = Request(url, headers={"User-Agent": "Mozilla/5.0"})
    return urlopen(req, timeout=20).read().decode("utf-8", "ignore")


def _decode_json_escaped_text(text: str) -> str:
    try:
        return json.loads(f'"{text}"')
    except Exception:
        return _decode_js_escaped_text(text)


def _extract_last_timestamp_line(description: str) -> str:
    lines = [line.strip() for line in description.splitlines() if line.strip()]
    last = ""
    for line in lines:
        if YT_TIMESTAMP_LINE_RE.match(line):
            last = line
    return last


def extract_youtube_fields(youtube_html: str) -> dict[str, str]:
    title = ""
    m_title = YT_OG_TITLE_RE.search(youtube_html)
    if m_title:
        title = _decode_json_escaped_text(m_title.group(1).strip())

    description = ""
    m_desc = YT_SHORT_DESC_RE.search(youtube_html)
    if m_desc:
        description = _decode_json_escaped_text(m_desc.group(1))

    return {
        "youtubeTitle": title,
        "youtubeDescription": description,
        "descriptionLastTimestampLine": _extract_last_timestamp_line(description),
    }


def build_episodes(rows: list[dict[str, str | int]]) -> list[dict[str, str | int]]:
    episodes: list[dict[str, str | int]] = []
    for row in rows:
        ytid = str(row.get("ytid", "")).strip()
        youtube_url = f"https://www.youtube.com/watch?v={ytid}" if ytid else ""
        episode = {
            "episodeIndex": int(row.get("episode_index", 0)),
            "epId": str(row.get("epid", "")).strip(),
            "date": str(row.get("date", "")).strip(),
            "titleZh": str(row.get("title", "")).strip(),
            "ytId": ytid,
            "youtubeUrl": youtube_url,
            "youtubeTitle": "",
            "youtubeDescription": "",
            "descriptionLastTimestampLine": "",
        }
        episodes.append(episode)
    return episodes


def enrich_with_youtube(episodes: list[dict[str, str | int]]) -> None:
    for episode in episodes:
        youtube_url = str(episode.get("youtubeUrl", "")).strip()
        if not youtube_url:
            continue
        try:
            yt_html = fetch_html(youtube_url)
        except Exception:
            continue
        fields = extract_youtube_fields(yt_html)
        episode.update(fields)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extract episodes from daai.tv program page and enrich with YouTube fields."
    )
    parser.add_argument("--url", required=True, help="Program page URL")
    parser.add_argument("--out", default="episodes.json", help="Output JSON file path.")
    parser.add_argument(
        "--raw-out",
        default="",
        help="Optional raw extraction JSON output path (e.g. episodes_raw.json).",
    )
    args = parser.parse_args()

    html = fetch_html(args.url)
    rows = extract_episode_rows_from_html(html)
    if args.raw_out:
        Path(args.raw_out).write_text(
            json.dumps(rows, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    episodes = build_episodes(rows)
    enrich_with_youtube(episodes)
    Path(args.out).write_text(
        json.dumps(episodes, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"count\t{len(episodes)}")
    print(f"out\t{args.out}")


if __name__ == "__main__":
    main()
