from pathlib import Path

from docx import Document
from docx.enum.text import WD_COLOR_INDEX

import generate_review


def _write_review_template(path: Path) -> None:
    doc = Document()
    doc.add_paragraph("外文編譯中心QCD")
    doc.add_paragraph("姓名: {{NAME}}")
    doc.add_paragraph("{{MONTH}}")
    doc.add_paragraph("本月精進目標:")
    doc.save(path)


def test_generate_review_renders_header_fields_from_txt(tmp_path: Path) -> None:
    template_path = tmp_path / "review_template.docx"
    source_txt = tmp_path / "review.txt"
    output_path = tmp_path / "review_output.docx"

    _write_review_template(template_path)
    source_txt.write_text(
        "\n".join(
            [
                "NAME: 王小明",
                "MONTH: 2022年11月",
            ]
        ),
        encoding="utf-8",
    )

    generate_review.generate_review(template_path, source_txt, output_path)

    out_doc = Document(output_path)
    assert [p.text for p in out_doc.paragraphs] == [
        "外文編譯中心QCD",
        "姓名: 王小明",
        "2022年11月",
        "本月精進目標:",
    ]
    goal_label_runs = out_doc.paragraphs[3].runs
    assert goal_label_runs
    assert all(
        run.font.highlight_color == WD_COLOR_INDEX.YELLOW for run in goal_label_runs
    )
    assert all(run.font.size is None for run in goal_label_runs)


def test_parse_input_supports_key_value_fields(tmp_path: Path) -> None:
    source_txt = tmp_path / "review.txt"
    source_txt.write_text(
        "\n".join(
            [
                "NAME: Alice",
                "MONTH: 2022年12月",
            ]
        ),
        encoding="utf-8",
    )

    data = generate_review.parse_input(source_txt)

    assert data == {"NAME": "Alice", "MONTH": "2022年12月"}


def test_resolve_template_path_accepts_relative_repo_template() -> None:
    resolved = generate_review.resolve_template_path(Path("templates/review_template.docx"))
    assert resolved.exists()
