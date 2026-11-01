#!/usr/bin/env python3
"""Convert bilingual Docs TXT files into SRT files."""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

from generate_subs import is_full_line_comment


DEFAULT_FPS = 30
TIME_TOKEN = r"\d{1,2}:\d{2}:\d{2}:\d{2}"
TIMED_LINE_RE = re.compile(
    rf"^(?P<start>{TIME_TOKEN})[\t \u3000]+(?P<end>{TIME_TOKEN})"
    rf"(?:[\t \u3000]+(?P<text>.*))?$"
)
STAR_MARKER_RE = re.compile(r"\s*\*\s*$")
HASH_MARKER_RE = re.compile(r"\s*#\s*$")


@dataclass
class TextBlock:
    start: str
    end: str
    content: list[str] = field(default_factory=list)


def decode_input(path: Path) -> str:
    raw = path.read_bytes()
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return raw.decode("utf-16")
    for encoding in ("utf-8-sig", "big5", "cp950", "gb18030"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise UnicodeError(f"Unable to decode input file: {path}")


def frame_timecode_to_srt(
    value: str, fps: int, offset_seconds: float = 0
) -> str:
    hours, minutes, seconds, frames = map(int, value.split(":"))
    total_ms = (
        ((hours * 3600 + minutes * 60 + seconds) * 1000)
        + round(frames * 1000 / fps)
        + round(offset_seconds * 1000)
    )
    if total_ms < 0:
        raise ValueError(
            f"timestamp offset moves {value} before 00:00:00,000"
        )
    hours = total_ms // 3_600_000
    total_ms %= 3_600_000
    minutes = total_ms // 60_000
    total_ms %= 60_000
    seconds = total_ms // 1_000
    milliseconds = total_ms % 1_000
    return f"{hours:02}:{minutes:02}:{seconds:02},{milliseconds:03}"


def parse_blocks(text: str) -> list[TextBlock]:
    blocks: list[TextBlock] = []
    current: TextBlock | None = None
    for raw_line in text.lstrip("\ufeff").splitlines():
        line = raw_line.rstrip()
        if is_full_line_comment(line):
            continue
        candidate = STAR_MARKER_RE.sub("", line)
        match = TIMED_LINE_RE.fullmatch(candidate)
        if match is not None:
            current = TextBlock(match.group("start"), match.group("end"))
            blocks.append(current)
            if match.group("text"):
                current.content.append(match.group("text").strip())
            continue
        if current is not None and line.strip():
            current.content.append(line.strip())
    return blocks


def split_bilingual(line: str) -> tuple[str, str | None]:
    if "//" not in line:
        return line.strip(), None
    source, english = line.split("//", 1)
    return source.strip(), english.strip() or None


def english_for_block(source: TextBlock, baseline: TextBlock) -> list[str]:
    baseline_parts = [split_bilingual(line) for line in baseline.content]
    source_parts = [split_bilingual(line) for line in source.content]
    baseline_sources = [part[0] for part in baseline_parts]
    english: list[str] = []
    source_index = 0

    def add_translation(part: tuple[str, str | None]) -> None:
        text = part[1] if not part[0] and part[1] is not None else part[0]
        text = HASH_MARKER_RE.sub("", text).rstrip()
        if text:
            english.append(text)

    for expected_source in baseline_sources:
        match_index = next(
            (
                index
                for index in range(source_index, len(source_parts))
                if source_parts[index][0] == expected_source
            ),
            None,
        )
        if match_index is None:
            if source_index == len(source_parts):
                break
            raise ValueError(
                f"source text does not match baseline at {source.start}"
            )
        for part in source_parts[source_index:match_index]:
            add_translation(part)
        inline_english = source_parts[match_index][1]
        if inline_english is not None:
            english.append(HASH_MARKER_RE.sub("", inline_english).rstrip())
        source_index = match_index + 1

    for part in source_parts[source_index:]:
        add_translation(part)
    return english


def build_srt(
    translated_text: str,
    baseline_text: str,
    fps: int = DEFAULT_FPS,
    offset_seconds: float = 0,
) -> str:
    translated = parse_blocks(translated_text)
    baseline = parse_blocks(baseline_text)
    translated_timeline = [(block.start, block.end) for block in translated]
    baseline_timeline = [(block.start, block.end) for block in baseline]
    if translated_timeline != baseline_timeline:
        raise ValueError("translated timeline does not match baseline timeline")

    output: list[str] = []
    for index, (source, original) in enumerate(zip(translated, baseline), 1):
        output.append(str(index))
        start = frame_timecode_to_srt(source.start, fps, offset_seconds)
        end = frame_timecode_to_srt(source.end, fps, offset_seconds)
        output.append(f"{start} --> {end}")
        output.extend(english_for_block(source, original))
        output.append("")
    return "\n".join(output)


def baseline_path_for(path: Path) -> Path:
    return path.with_name(f"{path.stem}.baseline.txt")


def is_docs_input(path: Path) -> bool:
    name = path.name.casefold()
    return (
        path.is_file()
        and path.suffix.casefold() == ".txt"
        and not name.startswith("~")
        and not name.endswith(".baseline.txt")
        and any(token in name for token in ("chus字幕", "英文字幕", "super"))
    )


def discover_input_paths(directory: Path) -> list[Path]:
    return sorted(path for path in directory.glob("*.txt") if is_docs_input(path))


def convert_txt(
    path: Path,
    fps: int = DEFAULT_FPS,
    offset_seconds: float = 0,
) -> Path:
    baseline_path = baseline_path_for(path)
    if not baseline_path.is_file():
        raise FileNotFoundError(f"matching baseline not found: {baseline_path}")
    output = path.with_suffix(".srt")
    srt = build_srt(
        decode_input(path),
        decode_input(baseline_path),
        fps,
        offset_seconds,
    )
    output.write_text(srt, encoding="utf-8")
    return output


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="convert-docs",
        description=(
            "Convert bilingual Docs TXT files into SRT files. By default, "
            "convert recognized non-baseline TXT files in the current directory."
        ),
    )
    parser.add_argument("paths", nargs="*", type=Path)
    parser.add_argument("--fps", type=int, default=DEFAULT_FPS)
    parser.add_argument("--offset-seconds", type=float, default=0)
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)
    targets = args.paths or discover_input_paths(Path.cwd())
    if not targets:
        print("[warn] no recognized non-baseline Docs TXT files found", file=sys.stderr)
        return 1

    failed = False
    for path in targets:
        if not is_docs_input(path):
            print(f"[skip] not a Docs TXT input: {path}", file=sys.stderr)
            continue
        try:
            output = convert_txt(path, args.fps, args.offset_seconds)
        except (FileNotFoundError, UnicodeError, ValueError) as error:
            print(f"[error] {error}", file=sys.stderr)
            failed = True
            continue
        print(f"[written] {output}")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
