from pathlib import Path
from unittest.mock import patch

from docx import Document

from generate_posts import generate_post, parse_post_text


POST_TEXT = """## Da Ai Journal - Staying Young at 105 (大愛全紀實 - 人生歌未央 [1])

Video: https://www.youtube.com/watch?v=example

Lin You-mao is 105 years old—and he's still playing badminton.

#DaAiJournal #StayingYoungAt105

林友茂今年105歲了，至今仍在打羽球。

#大愛全紀實 #人生歌未央

### Sources

9/20 World Cleanup Day 世界環境清潔日

https://www.worldcleanupday.org/
"""


def test_parse_completed_post_text() -> None:
    post = parse_post_text(POST_TEXT)

    assert post["title"] == (
        "Da Ai Journal - Staying Young at 105 (大愛全紀實 - 人生歌未央 [1])"
    )
    assert post["video_url"] == "https://www.youtube.com/watch?v=example"
    assert post["post_en"] == (
        "Lin You-mao is 105 years old—and he's still playing badminton."
    )
    assert post["hashtags_en"] == "#DaAiJournal #StayingYoungAt105"
    assert post["post_zh"] == "林友茂今年105歲了，至今仍在打羽球。"
    assert post["hashtags_zh"] == "#大愛全紀實 #人生歌未央"
    assert post["ref_title"] == "9/20 World Cleanup Day 世界環境清潔日"
    assert post["ref_url"] == "https://www.worldcleanupday.org/"


def test_generate_completed_post_uses_post_template_without_draft_header(
    tmp_path: Path,
) -> None:
    input_path = tmp_path / "post.txt"
    output_path = tmp_path / "post.docx"
    input_path.write_text(POST_TEXT, encoding="utf-8")

    with patch(
        "generate_posts.fetch_youtube_video_descriptions",
        return_value=("English video summary.", "中文影片摘要。"),
    ):
        generate_post(
            input_path=input_path,
            template_path=Path("templates/post_template.docx"),
            output_path=output_path,
        )

    texts = [paragraph.text for paragraph in Document(output_path).paragraphs]
    assert "標題" not in texts
    assert not any(text.startswith("9/20(日)") for text in texts)
    assert texts[0] == (
        "Da Ai Journal - Staying Young at 105 "
        "(大愛全紀實 - 人生歌未央 [1])\n\n"
        "Lin You-mao is 105 years old—and he's still playing badminton."
    )
    assert "https://www.worldcleanupday.org/" in texts
    assert "9/20 World Cleanup Day 世界環境清潔日" in texts
    assert "https://www.youtube.com/watch?v=example" in texts
    assert "English video summary." in texts
    assert "中文影片摘要。" in texts
