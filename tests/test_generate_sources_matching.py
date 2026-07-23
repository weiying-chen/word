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
