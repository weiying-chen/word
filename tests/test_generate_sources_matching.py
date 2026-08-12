from pathlib import Path

from generate_sources import _match_subtitle_files


def test_matches_filename_phrases_across_title_and_description(tmp_path: Path) -> None:
    subtitle = tmp_path / "大愛真健康第1136集_ch_體態雕塑 改善臀型.txt"
    subtitle.write_text("subtitle", encoding="utf-8")
    episodes = [
        {
            "epId": "video-a",
            "youtubeTitle": "告別駝背術｜體態雕塑｜大愛真健康",
            "youtubeDescription": "改善圓肩駝背。",
        },
        {
            "epId": "video-b",
            "youtubeTitle": "3招臀部增肌術｜體態雕塑｜大愛真健康",
            "youtubeDescription": "提升臀肌力量，改善臀型，同時保護腰椎。",
        },
    ]

    matches = _match_subtitle_files(episodes, tmp_path)

    assert matches == {1: subtitle}


def test_does_not_guess_when_phrase_match_is_ambiguous(tmp_path: Path) -> None:
    (tmp_path / "大愛真健康第1000集_ch_體態雕塑.txt").write_text(
        "subtitle", encoding="utf-8"
    )
    episodes = [
        {"epId": "a", "youtubeTitle": "體態雕塑 A", "youtubeDescription": ""},
        {"epId": "b", "youtubeTitle": "體態雕塑 B", "youtubeDescription": ""},
    ]

    assert _match_subtitle_files(episodes, tmp_path) == {}


def test_matches_internal_episode_number_using_alias(tmp_path: Path) -> None:
    subtitle = tmp_path / "大愛真健康第1102集_ch_功能性訓練 彎腰取物.txt"
    subtitle.write_text("subtitle", encoding="utf-8")
    episodes = [
        {
            "epId": "youtube-video-id",
            "youtubeTitle": "3招教你輕鬆取物",
            "youtubeDescription": "",
        }
    ]

    matches = _match_subtitle_files(
        episodes, tmp_path, {"1102": "youtube-video-id"}
    )

    assert matches == {0: subtitle}


def test_matches_date_only_filename(tmp_path: Path) -> None:
    subtitle = tmp_path / "20260416.txt"
    subtitle.write_text("subtitle", encoding="utf-8")
    episodes = [
        {"epId": "a", "date": "2026-04-15", "youtubeTitle": "First"},
        {"epId": "b", "date": "2026-04-16", "youtubeTitle": "Second"},
    ]

    assert _match_subtitle_files(episodes, tmp_path) == {1: subtitle}


def test_date_only_filename_warns_for_missing_date(
    tmp_path: Path, capsys
) -> None:
    (tmp_path / "20260417.txt").write_text("subtitle", encoding="utf-8")
    episodes = [{"epId": "a", "date": "2026-04-16", "youtubeTitle": "20260417"}]

    assert _match_subtitle_files(episodes, tmp_path) == {}
    assert capsys.readouterr().err == (
        "[warn] 20260417.txt: no episode found for 2026-04-17\n"
    )


def test_date_only_filename_warns_for_duplicate_date(
    tmp_path: Path, capsys
) -> None:
    (tmp_path / "20260416.txt").write_text("subtitle", encoding="utf-8")
    episodes = [
        {"epId": "a", "date": "2026-04-16", "youtubeTitle": "First 20260416"},
        {"epId": "b", "date": "2026-04-16", "youtubeTitle": "Second 20260416"},
    ]

    assert _match_subtitle_files(episodes, tmp_path) == {}
    assert capsys.readouterr().err == (
        "[warn] 20260416.txt: 2 episodes found for 2026-04-16; "
        "skipped because the date is ambiguous\n"
    )
