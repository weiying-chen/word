from pathlib import Path

import pytest

from convert_docs import (
    build_srt,
    convert_txt,
    discover_input_paths,
)


def test_convert_docs_wrapper_uses_project_virtual_environment() -> None:
    wrapper = Path(__file__).resolve().parents[1] / "convert-docs"

    assert wrapper.stat().st_mode & 0o111
    assert ".venv/bin/python" in wrapper.read_text(encoding="utf-8")


def test_discovery_uses_recognized_non_baseline_docs_txt_only(
    tmp_path: Path,
) -> None:
    subtitle = tmp_path / "episode_chus字幕.txt"
    super_file = tmp_path / "episode_super.txt"
    subtitle.write_text("content", encoding="utf-8")
    super_file.write_text("content", encoding="utf-8")
    (tmp_path / "episode_chus字幕.baseline.txt").write_text(
        "baseline", encoding="utf-8"
    )
    (tmp_path / "~episode_super.txt").write_text("backup", encoding="utf-8")
    (tmp_path / "interview.txt").write_text("unrelated", encoding="utf-8")

    assert discover_input_paths(tmp_path) == [subtitle, super_file]


def test_build_srt_supports_separate_and_inline_subtitle_english() -> None:
    baseline = (
        "00:00:01:15\t00:00:03:00\t中文\n"
        "00:00:03:00\t00:00:05:15\t中文//Existing English\n"
    )
    translated = (
        "00:00:01:15\t00:00:03:00\t中文 *\n"
        "New English #\n"
        "00:00:03:00\t00:00:05:15\t中文//Existing English\n"
    )

    assert build_srt(translated, baseline, fps=30) == (
        "1\n"
        "00:00:01,500 --> 00:00:03,000\n"
        "New English\n\n"
        "2\n"
        "00:00:03,000 --> 00:00:05,500\n"
        "Existing English\n"
    )


def test_build_srt_ignores_full_line_comments_between_cues() -> None:
    baseline = (
        "00:00:01:00\t00:00:02:00\t第一句\n"
        "00:00:02:00\t00:00:03:00\t第二句\n"
    )
    translated = (
        "00:00:01:00\t00:00:02:00\t第一句\n"
        "First line.\n\n"
        "// host\n\n"
        "00:00:02:00\t00:00:03:00\t第二句\n"
        "Second line.\n\n"
        "   // head\n"
    )

    output = build_srt(translated, baseline, fps=30)

    assert "host" not in output
    assert "head" not in output
    assert "First line." in output
    assert "Second line." in output


def test_build_srt_uses_baseline_to_isolate_multiline_super_english() -> None:
    baseline = (
        "00:00:03:00\t00:00:06:00\n"
        "主持人\n"
        "陳竹琪\n\n"
        "製作團隊\n\n"
        "00:00:06:00\t00:00:08:00\n"
        "風月同天\n"
    )
    translated = (
        "00:00:03:00\t00:00:06:00 *\n"
        "主持人\n"
        "陳竹琪\n"
        "Chen Zhu-qi\n"
        "Presenter\n\n"
        "製作團隊\n"
        "Production Team\n\n"
        "00:00:06:00\t00:00:08:00 *\n"
        "風月同天\n"
    )

    assert build_srt(translated, baseline, fps=30) == (
        "1\n"
        "00:00:03,000 --> 00:00:06,000\n"
        "Chen Zhu-qi\n"
        "Presenter\n"
        "Production Team\n\n"
        "2\n"
        "00:00:06,000 --> 00:00:08,000\n"
    )


def test_convert_txt_reads_utf16_and_writes_utf8_srt(tmp_path: Path) -> None:
    source = tmp_path / "episode_chus字幕.txt"
    baseline = tmp_path / "episode_chus字幕.baseline.txt"
    baseline.write_text(
        "00:00:00:00\t00:00:01:00\t中文\n", encoding="utf-16"
    )
    source.write_text(
        "00:00:00:00\t00:00:01:00\t中文\nEnglish\n",
        encoding="utf-16",
    )

    output = convert_txt(source, fps=30)

    assert output == source.with_suffix(".srt")
    assert output.read_text(encoding="utf-8") == (
        "1\n00:00:00,000 --> 00:00:01,000\nEnglish\n"
    )


def test_build_srt_rejects_timeline_mismatch() -> None:
    with pytest.raises(ValueError, match="timeline does not match"):
        build_srt(
            "00:00:00:00\t00:00:01:00\t中文\nEnglish\n",
            "00:00:00:00\t00:00:02:00\t中文\n",
            fps=30,
        )
