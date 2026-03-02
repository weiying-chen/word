from pathlib import Path
import warnings
import zipfile

from docx import Document
from lxml import etree

import generate_news


def test_parse_input_multiline_summary_and_body(tmp_path: Path) -> None:
    input_path = tmp_path / "news_input.txt"
    input_path.write_text(
        "\n".join(
            [
                "TITLE: Sample News Title",
                "TITLE_URL: https://example.com/news",
                "SUMMARY:",
                "Summary line one.",
                "(  11/16~17 )",
                "",
                "SUPER_PEOPLE:",
                "病患 | 羅伯托",
                "Roberto",
                "Patient",
                "",
                "BODY:",
                "1_0014",
                "中文內文。",
                "English line.",
                "",
            ]
        ),
        encoding="utf-8",
    )

    data = generate_news.parse_input(input_path)
    assert data["TITLE"] == "Sample News Title"
    assert data["TITLE_URL"] == "https://example.com/news"
    assert data["SUMMARY"] == "Summary line one.\n(  11/16~17 )"
    assert data["SUPER_PEOPLE"] == "病患 | 羅伯托\nRoberto\nPatient"
    assert data["BODY"] == "1_0014\n中文內文。\nEnglish line."


def test_parse_input_fallback_encoding_warns_and_rewrites_utf8(
    tmp_path: Path,
) -> None:
    input_path = tmp_path / "news_input.txt"
    input_path.write_bytes("TITLE: Café News\n".encode("cp1252"))

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        data = generate_news.parse_input(input_path)

    assert data["TITLE"] == "Café News"
    assert caught
    assert "fallback encoding" in str(caught[0].message).lower()
    assert "cp1252" in str(caught[0].message).lower()
    assert input_path.read_bytes() == "TITLE: Café News\n".encode("utf-8")


def test_generate_news_renders_title_summary_marker_and_body(tmp_path: Path) -> None:
    input_path = tmp_path / "news_input.txt"
    output_path = tmp_path / "news_output.docx"
    input_path.write_text(
        "\n".join(
            [
                "TITLE: Community Clinic Brings Care to Coastal Town",
                "TITLE_URL: https://example.com/news/story",
                "SUMMARY:",
                "Volunteers organized a two-day clinic to support families in a coastal town.",
                "(  11/16~17 )",
                "",
                "SUPER_PEOPLE:",
                "病患 | 羅伯托",
                "Roberto",
                "Patient",
                "",
                "BODY:",
                "1_0014",
                "Two days of free screenings brought steady lines of local residents.",
                "Residents arrived early to receive free screenings and follow-up advice.",
                "",
                "2_0025",
                "(NS)",
                "A longtime fisherman said his eyesight had been fading for years.",
                "One resident said his vision had been gradually worsening for years.",
                "",
                "(  13   Guest )",
                "/*SUPER:",
                "Guest|Participant//",
                "I can barely see from my left eye now.//",
                "*/",
                "I can barely see from my left eye now,",
                "",
            ]
        ),
        encoding="utf-8",
    )

    generate_news.generate_news(input_path, output_path)

    doc = Document(output_path)
    texts = [p.text for p in doc.paragraphs]
    assert texts[:14] == [
        "Community Clinic Brings Care to Coastal Town",
        "",
        "Volunteers organized a two-day clinic to support families in a coastal town.",
        "(  11/16~17 )",
        "",
        "病患 | 羅伯托",
        "Roberto",
        "Patient",
        "<",
        "",
        "1_0014",
        "Two days of free screenings brought steady lines of local residents.",
        "Residents arrived early to receive free screenings and follow-up advice.",
        "",
    ]

    assert doc.paragraphs[2].runs[0].font.highlight_color.name == "WHITE"
    assert doc.paragraphs[8].runs[0].font.highlight_color.name == "WHITE"
    assert doc.paragraphs[10].runs[0].font.highlight_color.name == "BRIGHT_GREEN"
    assert doc.paragraphs[14].runs[0].font.highlight_color.name == "BRIGHT_GREEN"

    with zipfile.ZipFile(output_path) as zf:
        document_xml = etree.fromstring(zf.read("word/document.xml"))

    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    first_paragraph = document_xml.find(".//w:body/w:p", ns)
    assert first_paragraph is not None
    hyperlinks = first_paragraph.findall("w:hyperlink", ns)
    assert len(hyperlinks) == 1


def test_default_output_path_uses_source_stem_with_final_suffix(tmp_path: Path) -> None:
    source = tmp_path / "coastal_story.docx"
    output_dir = tmp_path / "output"
    output = generate_news.default_output_path(source, output_dir)
    assert output == output_dir / "coastal_story_final.docx"


def test_default_output_path_preserves_existing_final_suffix(tmp_path: Path) -> None:
    source = tmp_path / "coastal_story_final.docx"
    output_dir = tmp_path / "output"
    output = generate_news.default_output_path(source, output_dir)
    assert output == output_dir / "coastal_story_final.docx"
