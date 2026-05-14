from pathlib import Path
import json

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.shared import Pt

import generate_review
from style_tokens import REVIEW_TEXT_SIZE_PT


def _write_review_template(path: Path) -> None:
    doc = Document()
    doc.add_paragraph("外文編譯中心QCD")
    doc.add_paragraph("姓名: {{NAME}}")
    doc.add_paragraph("{{MONTH}}")
    doc.add_paragraph("本月精進目標:")
    table = doc.add_table(rows=3, cols=4)
    table.cell(0, 0).text = "日期"
    table.cell(0, 1).text = "(例行)字幕翻譯"
    table.cell(0, 2).text = "編輯回饋"
    table.cell(0, 3).text = "主管回饋"
    table.cell(1, 0).text = ""
    table.cell(1, 1).text = ""
    table.cell(1, 2).text = ""
    table.cell(1, 3).text = ""
    table.cell(2, 0).text = "日期"
    table.cell(2, 1).text = "(例行)字幕審稿"
    doc.save(path)


def test_generate_review_renders_header_fields_from_sources(tmp_path: Path) -> None:
    template_path = tmp_path / "review_template.docx"
    source_txt = tmp_path / "review.txt"
    tasks_json = tmp_path / "tasks.json"
    output_path = tmp_path / "review_output.docx"

    _write_review_template(template_path)
    source_txt.write_text(
        "\n".join(
            [
                "NAME: 王小明",
            ]
        ),
        encoding="utf-8",
    )
    tasks_json.write_text(
        json.dumps({"exportMonth": "2022-11"}, ensure_ascii=False),
        encoding="utf-8",
    )

    generate_review.generate_review(
        template_path,
        source_txt,
        output_path,
        tasks_json,
    )

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
    assert all(run.font.size == Pt(REVIEW_TEXT_SIZE_PT) for run in goal_label_runs)


def test_parse_input_supports_key_value_fields(tmp_path: Path) -> None:
    source_txt = tmp_path / "review.txt"
    source_txt.write_text(
        "\n".join(
            [
                "NAME: Alice",
            ]
        ),
        encoding="utf-8",
    )

    data = generate_review.parse_input(source_txt)

    assert data == {"NAME": "Alice"}


def test_resolve_template_path_accepts_relative_repo_template() -> None:
    resolved = generate_review.resolve_template_path(Path("templates/review_template.docx"))
    assert resolved.exists()


def test_generate_review_populates_regular_translation_rows(tmp_path: Path) -> None:
    template_path = tmp_path / "review_template.docx"
    source_txt = tmp_path / "review.txt"
    tasks_json = tmp_path / "tasks.json"
    output_path = tmp_path / "review_output.docx"

    _write_review_template(template_path)
    source_txt.write_text("NAME: 王小明\n", encoding="utf-8")
    tasks_json.write_text(
        json.dumps(
            {
                "exportMonth": "2026-05",
                "tasks": [
                    {
                        "title": "回眸(中翻英)",
                        "deadlineIso": "2026-05-08T00:00:00.000Z",
                        "workMinutes": 240,
                        "comments": ["This is a comment"],
                    }
                ],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    generate_review.generate_review(
        template_path,
        source_txt,
        output_path,
        tasks_json,
    )

    out_doc = Document(output_path)
    table = out_doc.tables[0]
    assert table.cell(1, 0).text.strip() == "5/8"
    assert table.cell(1, 1).text.strip() == "1.\n回眸(中翻英)\n實際作業時間:4時"
    assert table.cell(1, 2).text.strip() == "• This is a comment"


def test_generate_review_inserts_rows_for_multiple_tasks(tmp_path: Path) -> None:
    template_path = tmp_path / "review_template.docx"
    source_txt = tmp_path / "review.txt"
    tasks_json = tmp_path / "tasks.json"
    output_path = tmp_path / "review_output.docx"

    _write_review_template(template_path)
    source_txt.write_text("NAME: 王小明\n", encoding="utf-8")
    tasks_json.write_text(
        json.dumps(
            {
                "exportMonth": "2026-05",
                "tasks": [
                    {
                        "title": "A",
                        "deadlineIso": "2026-05-08T00:00:00.000Z",
                        "workMinutes": 60,
                        "comments": ["c1"],
                    },
                    {
                        "title": "B",
                        "deadlineIso": "2026-05-09T00:00:00.000Z",
                        "workMinutes": 120,
                        "comments": ["c2"],
                    },
                ],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    generate_review.generate_review(
        template_path,
        source_txt,
        output_path,
        tasks_json,
    )

    out_doc = Document(output_path)
    table = out_doc.tables[0]
    assert table.cell(1, 1).text.strip() == "1.\nA\n實際作業時間:1時"
    assert table.cell(2, 1).text.strip() == "2.\nB\n實際作業時間:2時"
    assert table.cell(3, 0).text.strip() == "日期"
    assert table.cell(3, 1).text.strip() == "(例行)字幕審稿"


def test_generate_review_uses_template_font_for_generated_table_content(tmp_path: Path) -> None:
    template_path = tmp_path / "review_template.docx"
    source_txt = tmp_path / "review.txt"
    tasks_json = tmp_path / "tasks.json"
    output_path = tmp_path / "review_output.docx"

    _write_review_template(template_path)
    source_txt.write_text("NAME: 王小明\n", encoding="utf-8")
    tasks_json.write_text(
        json.dumps(
            {
                "exportMonth": "2026-05",
                "tasks": [
                    {
                        "title": "回眸(中翻英)",
                        "deadlineIso": "2026-05-08T00:00:00.000Z",
                        "workMinutes": 240,
                        "comments": ["This is a comment"],
                    }
                ],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    generate_review.generate_review(
        template_path,
        source_txt,
        output_path,
        tasks_json,
    )

    out_doc = Document(output_path)
    row_cell_runs = out_doc.tables[0].cell(1, 1).paragraphs[0].runs
    comment_runs = out_doc.tables[0].cell(1, 2).paragraphs[0].runs
    assert row_cell_runs
    assert comment_runs
    assert all(run.font.size is None for run in row_cell_runs)
    assert all(run.font.size is None for run in comment_runs)
