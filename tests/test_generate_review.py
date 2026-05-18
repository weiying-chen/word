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
    table = doc.add_table(rows=7, cols=4)
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
    table.cell(3, 0).text = ""
    table.cell(3, 1).text = ""
    table.cell(3, 2).text = ""
    table.cell(3, 3).text = ""
    table.cell(4, 0).text = "日期"
    table.cell(4, 1).text = "臨時工作"
    table.cell(5, 0).text = ""
    table.cell(5, 1).text = ""
    table.cell(5, 2).text = ""
    table.cell(5, 3).text = ""
    table.cell(6, 0).text = "本月工作心得:"
    doc.save(path)


def test_generate_review_renders_header_fields_from_sources(tmp_path: Path) -> None:
    template_path = tmp_path / "review_template.docx"
    tasks_json = tmp_path / "tasks.json"
    output_path = tmp_path / "review_output.docx"

    _write_review_template(template_path)
    tasks_json.write_text(
        json.dumps(
            [
                {
                    "name": "A",
                    "createdAt": "2022-11-08T00:00:00.000Z",
                    "workMinutes": 60,
                    "contentSeconds": 120,
                    "comments": [],
                    "children": [],
                }
            ],
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    generate_review.generate_review(
        template_path,
        output_path,
        tasks_json,
    )

    out_doc = Document(output_path)
    assert [p.text for p in out_doc.paragraphs] == [
        "外文編譯中心QCD",
        "姓名: 陳威穎",
        "2022年11月",
        "本月精進目標:",
    ]
    goal_label_runs = out_doc.paragraphs[3].runs
    assert goal_label_runs
    assert all(
        run.font.highlight_color == WD_COLOR_INDEX.YELLOW for run in goal_label_runs
    )
    assert all(run.font.size == Pt(REVIEW_TEXT_SIZE_PT) for run in goal_label_runs)


def test_resolve_template_path_accepts_relative_repo_template() -> None:
    resolved = generate_review.resolve_template_path(Path("templates/review_template.docx"))
    assert resolved.exists()


def test_generate_review_populates_regular_translation_rows(tmp_path: Path) -> None:
    template_path = tmp_path / "review_template.docx"
    tasks_json = tmp_path / "tasks.json"
    output_path = tmp_path / "review_output.docx"

    _write_review_template(template_path)
    tasks_json.write_text(
        json.dumps(
            [
                {
                    "name": "回眸(中翻英)",
                    "createdAt": "2026-05-08T00:00:00.000Z",
                    "workMinutes": 240,
                    "contentSeconds": 210,
                    "comments": ["This is a comment"],
                    "children": [],
                }
            ],
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    generate_review.generate_review(
        template_path,
        output_path,
        tasks_json,
    )

    out_doc = Document(output_path)
    table = out_doc.tables[0]
    assert table.cell(1, 0).text.strip() == "5/8"
    assert table.cell(1, 1).text.strip() == "1.\n回眸(中翻英)\n長度:3分30秒\n實際作業時間:4時"
    assert table.cell(1, 2).text.strip() == "• This is a comment"


def test_generate_review_inserts_rows_for_multiple_tasks(tmp_path: Path) -> None:
    template_path = tmp_path / "review_template.docx"
    tasks_json = tmp_path / "tasks.json"
    output_path = tmp_path / "review_output.docx"

    _write_review_template(template_path)
    tasks_json.write_text(
        json.dumps(
            [
                {
                    "name": "A",
                    "createdAt": "2026-05-08T00:00:00.000Z",
                    "workMinutes": 60,
                    "contentSeconds": 120,
                    "comments": ["c1"],
                    "children": [],
                },
                {
                    "name": "B",
                    "createdAt": "2026-05-09T00:00:00.000Z",
                    "workMinutes": 120,
                    "contentSeconds": 180,
                    "comments": ["c2"],
                    "children": [],
                },
            ],
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    generate_review.generate_review(
        template_path,
        output_path,
        tasks_json,
    )

    out_doc = Document(output_path)
    table = out_doc.tables[0]
    assert table.cell(1, 1).text.strip() == "1.\nA\n長度:2分\n實際作業時間:1時"
    assert table.cell(2, 1).text.strip() == "2.\nB\n長度:3分\n實際作業時間:2時"
    assert table.cell(3, 0).text.strip() == "日期"
    assert table.cell(3, 1).text.strip() == "(例行)字幕審稿"


def test_generate_review_uses_template_font_for_generated_table_content(tmp_path: Path) -> None:
    template_path = tmp_path / "review_template.docx"
    tasks_json = tmp_path / "tasks.json"
    output_path = tmp_path / "review_output.docx"

    _write_review_template(template_path)
    tasks_json.write_text(
        json.dumps(
            [
                {
                    "name": "回眸(中翻英)",
                    "createdAt": "2026-05-08T00:00:00.000Z",
                    "workMinutes": 240,
                    "contentSeconds": 210,
                    "comments": ["This is a comment"],
                    "children": [],
                }
            ],
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    generate_review.generate_review(
        template_path,
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


def test_generate_review_supports_top_level_tasks_list_with_new_field_names(
    tmp_path: Path,
) -> None:
    template_path = tmp_path / "review_template.docx"
    tasks_json = tmp_path / "tasks.json"
    output_path = tmp_path / "review_output.docx"

    _write_review_template(template_path)
    tasks_json.write_text(
        json.dumps(
            [
                {
                    "name": "A",
                    "createdAt": "2026-05-01T01:02:03Z",
                    "workMinutes": 60,
                    "contentSeconds": 210,
                    "comments": ["c1"],
                    "children": [],
                },
                {
                    "name": "B",
                    "createdAt": "2026-05-02T04:05:06Z",
                    "workMinutes": 120,
                    "comments": ["c2"],
                    "children": [],
                },
            ],
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    generate_review.generate_review(
        template_path,
        output_path,
        tasks_json,
    )

    out_doc = Document(output_path)
    assert out_doc.paragraphs[2].text == "2026年5月"
    table = out_doc.tables[0]
    assert table.cell(1, 0).text.strip() == "5/1"
    assert table.cell(1, 1).text.strip() == "1.\nA\n長度:3分30秒\n實際作業時間:1時"
    assert table.cell(2, 0).text.strip() == "5/2"
    assert table.cell(2, 1).text.strip() == "2.\nB\n實際作業時間:2時"


def test_generate_review_uses_notes_for_editor_feedback(tmp_path: Path) -> None:
    template_path = tmp_path / "review_template.docx"
    tasks_json = tmp_path / "tasks.json"
    output_path = tmp_path / "review_output.docx"

    _write_review_template(template_path)
    tasks_json.write_text(
        json.dumps(
            [
                {
                    "name": "A",
                    "createdAt": "2026-05-01T01:02:03Z",
                    "workMinutes": 60,
                    "contentSeconds": 120,
                    "notes": ["note one", "note two"],
                    "children": [],
                }
            ],
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    generate_review.generate_review(
        template_path,
        output_path,
        tasks_json,
    )

    out_doc = Document(output_path)
    table = out_doc.tables[0]
    assert table.cell(1, 2).text.strip() == "• note one\n• note two"


def test_generate_review_populates_temp_work_from_posts_children_only(
    tmp_path: Path,
) -> None:
    template_path = tmp_path / "review_template.docx"
    tasks_json = tmp_path / "tasks.json"
    output_path = tmp_path / "review_output.docx"

    _write_review_template(template_path)
    tasks_json.write_text(
        json.dumps(
            [
                {
                    "name": "主任務",
                    "createdAt": "2026-05-01T01:02:03Z",
                    "workMinutes": 60,
                    "contentSeconds": 120,
                    "children": [
                        {
                            "name": "POST A",
                            "type": "posts",
                            "createdAt": "2026-05-18T11:37:41.273370Z",
                            "workMinutes": 50,
                        },
                        {
                            "name": "NEWS B",
                            "type": "news",
                            "createdAt": "2026-05-18T11:37:41.273370Z",
                            "workMinutes": 40,
                        },
                    ],
                }
            ],
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    generate_review.generate_review(
        template_path,
        output_path,
        tasks_json,
    )

    out_doc = Document(output_path)
    table = out_doc.tables[0]
    # row 4 is 臨時工作 header, row 5 is first temp work row
    assert table.cell(5, 0).text.strip() == "5/18"
    assert table.cell(5, 1).text.strip() == "1.\nPOST A\n實際作業時間:50分"
    assert "長度:" not in table.cell(5, 1).text
