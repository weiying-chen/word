from generate_sources import _description_for_docx


def test_keeps_easy_fitness_summary_and_exercise_block() -> None:
    description = (
        "臀部肌肉是身體的重要支撐。改善臀型，同時保護腰椎。\n\n"
        "➯5分鐘動起來！\n"
        "00:45 內拍抬腿\n"
        "02:00 提臀踢腿\n"
        "03:25 瘦臀側抬"
    )

    assert _description_for_docx(description) == description


def test_keeps_only_first_summary_line_for_other_descriptions() -> None:
    assert _description_for_docx("第一行摘要\n第二行") == "第一行摘要"
