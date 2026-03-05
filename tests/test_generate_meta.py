import tempfile
import unittest
from pathlib import Path

from docx import Document

from generate_meta import default_output_path, generate_meta, parse_input


class RenderMetaTests(unittest.TestCase):
    def _build_template(self, path: Path) -> None:
        doc = Document()
        doc.add_paragraph("重點標")
        doc.add_paragraph("{{TITLE_EN}}")
        doc.add_paragraph("名字職銜")
        doc.add_paragraph("")
        doc.add_paragraph("{{PEOPLE}}")
        doc.add_paragraph("")
        doc.add_paragraph("YT簡介")
        doc.add_paragraph("{{OVERVIEW_EN}}")
        doc.save(str(path))

    def test_renders_title_people_overview(self) -> None:
        source_text = "\n".join(
            [
                "TITLE: 中文標題",
                "SUMMARY:",
                "中文摘要",
                "",
                "META_TITLE_EN: English Title",
                "META_OVERVIEW_EN:",
                "English overview.",
                "",
                "BODY:",
                "(  13   Alice )",
                "/*SUPER:",
                "病患│甲//",
                "引言一//",
                "*/",
                "Quote one.",
                "",
                "(  14   Bob )",
                "/*SUPER:",
                "醫師│乙//",
                "引言二//",
                "*/",
                "Quote two.",
                "",
            ]
        )

        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir_path = Path(tmpdir)
            template_path = tmpdir_path / "meta_template.docx"
            payload_path = tmpdir_path / "news_input.txt"
            output_path = tmpdir_path / "meta.docx"

            self._build_template(template_path)
            payload_path.write_text(source_text, encoding="utf-8")

            generate_meta(template_path, payload_path, output_path)

            doc = Document(str(output_path))
            texts = [p.text for p in doc.paragraphs]

        self.assertEqual(
            texts,
            [
                "重點標",
                "English Title",
                "名字職銜",
                "",
                "病患｜甲",
                "Alice",
                "{{病患}}",
                "",
                "醫師｜乙",
                "Bob",
                "{{醫師}}",
                "",
                "YT簡介",
                "English overview.",
            ],
        )

    def test_parse_input_extracts_meta_fields_and_people(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            input_path = Path(tmpdir) / "news_input.txt"
            input_path.write_text(
                "\n".join(
                    [
                        "TITLE: 中文標題",
                        "SUMMARY:",
                        "中文摘要",
                        "",
                        "META_TITLE_EN: English Title",
                        "META_OVERVIEW_EN:",
                        "English overview.",
                        "",
                        "BODY:",
                        "(  13   Alice )",
                        "/*SUPER:",
                        "病患│甲//",
                        "引言一//",
                        "*/",
                        "Quote one.",
                        "",
                    ]
                ),
                encoding="utf-8",
            )

            data = parse_input(input_path)

        self.assertEqual(data["title_zh"], "中文標題")
        self.assertEqual(data["summary_zh"], "中文摘要")
        self.assertEqual(data["title_en"], "English Title")
        self.assertEqual(data["overview_en"], "English overview.")
        self.assertEqual(
            data["people"],
            [
                {
                    "name_zh": "甲",
                    "name_en": "Alice",
                    "role_zh": "病患",
                    "role_en": "",
                    "org_en": "",
                }
            ],
        )

    def test_missing_en_fields_render_empty(self) -> None:
        source_text = "\n".join(
            [
                "TITLE: 中文標題",
                "SUMMARY:",
                "中文摘要",
                "",
                "BODY:",
                "",
            ]
        )

        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir_path = Path(tmpdir)
            template_path = tmpdir_path / "meta_template.docx"
            payload_path = tmpdir_path / "news_input.txt"
            output_path = tmpdir_path / "meta.docx"

            doc = Document()
            doc.add_paragraph("{{TITLE_EN}}")
            doc.add_paragraph("{{OVERVIEW_EN}}")
            doc.save(str(template_path))
            payload_path.write_text(source_text, encoding="utf-8")

            generate_meta(template_path, payload_path, output_path)

            doc = Document(str(output_path))
            texts = [p.text for p in doc.paragraphs]

        self.assertEqual(texts, ["", ""])

    def test_renders_from_news_txt_input(self) -> None:
        source_text = "\n".join(
            [
                "TITLE: 沿海義診守護居民健康",
                "TITLE_URL: https://example.com/news/story",
                "SUMMARY:",
                "A volunteer team hosted a two-day community clinic in a coastal town.",
                "(  11/16~17 )",
                "",
                "META_TITLE_EN: Coastal Clinic Restores Access to Care",
                "META_OVERVIEW_EN:",
                "A volunteer medical team brought screenings and referrals to local residents.",
                "",
                "BODY:",
                "1_0014",
                "居民一早就到現場排隊。",
                "Residents lined up early for registration and screening.",
                "",
                "(  13   Mr. Chen )",
                "/*SUPER:",
                "居民│陳先生//",
                "真的很感謝大家的幫忙//",
                "*/",
                "I am truly grateful for everyone's help.",
                "",
                "(  14   )",
                "/*SUPER:",
                "醫師│林醫師//",
                "我們會持續協助需要後續治療的居民//",
                "*/",
                "We will continue helping residents who need follow-up care.",
                "",
            ]
        )

        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir_path = Path(tmpdir)
            template_path = tmpdir_path / "meta_template.docx"
            payload_path = tmpdir_path / "news_input.txt"
            output_path = tmpdir_path / "meta.docx"

            self._build_template(template_path)
            payload_path.write_text(source_text, encoding="utf-8")

            generate_meta(template_path, payload_path, output_path)

            doc = Document(str(output_path))
            texts = [p.text for p in doc.paragraphs]

        self.assertEqual(
            texts,
            [
                "重點標",
                "Coastal Clinic Restores Access to Care",
                "名字職銜",
                "",
                "居民｜陳先生",
                "Mr. Chen",
                "{{居民}}",
                "",
                "醫師｜林醫師",
                "{{林醫師}}",
                "{{醫師}}",
                "",
                "YT簡介",
                "A volunteer medical team brought screenings and referrals to local residents.",
            ],
        )

    def test_meta_people_overrides_matching_blocks(self) -> None:
        source_text = "\n".join(
            [
                "TITLE: 測試標題",
                "SUMMARY:",
                "測試摘要",
                "",
                "META_TITLE_EN: Test English Title",
                "META_OVERVIEW_EN:",
                "Test English overview.",
                "",
                "META_PEOPLE:",
                "居民｜受訪者",
                "Guest (Edited)",
                "Resident",
                "",
                "醫師｜林醫師",
                "Dr. Lin",
                "Doctor",
                "Clinic Team",
                "",
                "BODY:",
                "(  13   Guest )",
                "/*SUPER:",
                "居民│受訪者//",
                "內容//",
                "*/",
                "English line.",
                "",
                "(  14   )",
                "/*SUPER:",
                "醫師│林醫師//",
                "內容//",
                "*/",
                "English line.",
                "",
            ]
        )

        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir_path = Path(tmpdir)
            template_path = tmpdir_path / "meta_template.docx"
            input_path = tmpdir_path / "news_input.txt"
            output_path = tmpdir_path / "meta.docx"

            self._build_template(template_path)
            input_path.write_text(source_text, encoding="utf-8")

            generate_meta(template_path, input_path, output_path)
            doc = Document(str(output_path))
            texts = [p.text for p in doc.paragraphs]

        self.assertEqual(
            texts,
            [
                "重點標",
                "Test English Title",
                "名字職銜",
                "",
                "居民｜受訪者",
                "Guest (Edited)",
                "Resident",
                "",
                "醫師｜林醫師",
                "Dr. Lin",
                "Doctor",
                "Clinic Team",
                "",
                "YT簡介",
                "Test English overview.",
            ],
        )


def test_default_output_path_uses_source_stem(tmp_path: Path) -> None:
    source = tmp_path / "sample_story_final.docx"
    output_dir = tmp_path / "output"
    output = default_output_path(source, output_dir)
    assert output == output_dir / "sample_story_標題職銜_final.docx"


def test_default_output_path_adds_final_suffix_when_missing(tmp_path: Path) -> None:
    source = tmp_path / "sample_story.docx"
    output_dir = tmp_path / "output"
    output = default_output_path(source, output_dir)
    assert output == output_dir / "sample_story_標題職銜_final.docx"


if __name__ == "__main__":
    unittest.main()
