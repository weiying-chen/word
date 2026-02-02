import json
import tempfile
import unittest
from pathlib import Path

from docx import Document

from generate_meta import generate_meta


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
        payload = {
            "title_zh": "標題",
            "summary_zh": "摘要",
            "narration_zh": [],
            "supers_zh": [],
            "report_zh": [],
            "people": [
                {
                    "name_zh": "甲",
                    "name_en": "Alice",
                    "role_zh": "病患",
                    "role_en": "Patient",
                },
                {
                    "name_zh": "乙",
                    "name_en": "Bob",
                    "role_zh": "醫師",
                    "role_en": "Doctor",
                },
            ],
            "title_en": "English Title",
            "overview_en": "English overview.",
        }

        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir_path = Path(tmpdir)
            template_path = tmpdir_path / "meta_template.docx"
            payload_path = tmpdir_path / "meta_filled.json"
            output_path = tmpdir_path / "meta.docx"

            self._build_template(template_path)
            payload_path.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )

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
                "Patient",
                "",
                "醫師｜乙",
                "Bob",
                "Doctor",
                "",
                "YT簡介",
                "English overview.",
            ],
        )

    def test_missing_en_fields_render_empty(self) -> None:
        payload = {
            "title_zh": "中文標題",
            "summary_zh": "中文摘要",
            "people": [],
        }

        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir_path = Path(tmpdir)
            template_path = tmpdir_path / "meta_template.docx"
            payload_path = tmpdir_path / "meta_filled.json"
            output_path = tmpdir_path / "meta.docx"

            doc = Document()
            doc.add_paragraph("{{TITLE_EN}}")
            doc.add_paragraph("{{OVERVIEW_EN}}")
            doc.save(str(template_path))
            payload_path.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )

            generate_meta(template_path, payload_path, output_path)

            doc = Document(str(output_path))
            texts = [p.text for p in doc.paragraphs]

        self.assertEqual(texts, ["", ""])


if __name__ == "__main__":
    unittest.main()
