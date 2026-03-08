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
                "TITLE_TEXT: Sample News Title",
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
    assert data["TITLE_TEXT"] == "Sample News Title"
    assert data["TITLE_URL"] == "https://example.com/news"
    assert data["SUMMARY"] == "Summary line one.\n(  11/16~17 )"
    assert data["SUPER_PEOPLE"] == "病患 | 羅伯托\nRoberto\nPatient"
    assert data["BODY"] == "1_0014\n中文內文。\nEnglish line."


def test_parse_input_fallback_encoding_warns_and_rewrites_utf8(
    tmp_path: Path,
) -> None:
    input_path = tmp_path / "news_input.txt"
    input_path.write_bytes("TITLE_TEXT: Café News\n".encode("cp1252"))

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        data = generate_news.parse_input(input_path)

    assert data["TITLE_TEXT"] == "Café News"
    assert caught
    assert "fallback encoding" in str(caught[0].message).lower()
    assert "cp1252" in str(caught[0].message).lower()
    assert input_path.read_bytes() == "TITLE_TEXT: Café News\n".encode("utf-8")


def test_generate_news_renders_title_summary_marker_and_body(tmp_path: Path) -> None:
    input_path = tmp_path / "news_input.txt"
    output_path = tmp_path / "news_output.docx"
    input_path.write_text(
        "\n".join(
            [
                "TITLE_TEXT: Community Clinic Brings Care to Coastal Town",
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


def test_generate_news_falls_back_to_url_when_title_text_missing(tmp_path: Path) -> None:
    input_path = tmp_path / "news_input.txt"
    output_path = tmp_path / "news_output.docx"
    input_path.write_text(
        "\n".join(
            [
                "TITLE_URL: https://example.com/news/story",
                "SUMMARY:",
                "Summary line one.",
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

    generate_news.generate_news(input_path, output_path)

    doc = Document(output_path)
    assert doc.paragraphs[0].text == "https://example.com/news/story"


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


def test_parse_sources_merges_docx_and_txt_sections(tmp_path: Path) -> None:
    source_docx = tmp_path / "source.docx"
    source_txt = tmp_path / "source.txt"

    doc = Document()
    doc.add_paragraph("Sample title from docx")
    doc.add_paragraph("https://example.com/story")
    doc.add_paragraph("Summary from docx")
    doc.add_paragraph("( 11/16~17 )")
    doc.save(source_docx)

    source_txt.write_text(
        "\n".join(
            [
                "SUPER_PEOPLE：",
                "Patient | Alex Wang",
                "Alex Wang",
                "Patient",
                "",
                "字幕：",
                "1_0001",
                "中文內文。",
                "English body line.",
            ]
        ),
        encoding="utf-8",
    )

    data = generate_news.parse_sources(source_docx, source_txt)

    assert data["TITLE_TEXT"] == "Sample title from docx"
    assert data["TITLE_URL"] == "https://example.com/story"
    assert data["SUMMARY"] == "Summary from docx"
    assert data["SUPER_PEOPLE"] == "Patient | Alex Wang\nAlex Wang\nPatient"
    assert data["BODY"] == "1_0001\n中文內文。\nEnglish body line."


def test_generate_news_from_sources_renders_without_intermediate_input(
    tmp_path: Path,
) -> None:
    source_docx = tmp_path / "source.docx"
    source_txt = tmp_path / "source.txt"
    output_path = tmp_path / "news_output.docx"

    doc = Document()
    doc.add_paragraph("Direct generation title")
    doc.add_paragraph("https://example.com/direct")
    doc.add_paragraph("Direct summary line")
    doc.add_paragraph("( 11/16~17 )")
    doc.save(source_docx)

    source_txt.write_text(
        "\n".join(
            [
                "SUPER_PEOPLE：",
                "Patient | Alex Wang",
                "Alex Wang",
                "Patient",
                "",
                "字幕：",
                "1_0001",
                "中文內文。",
                "English body line.",
                "",
            ]
        ),
        encoding="utf-8",
    )

    generate_news.generate_news_from_sources(source_docx, source_txt, output_path)

    texts = [p.text for p in Document(output_path).paragraphs]
    assert texts[:10] == [
        "Direct generation title",
        "",
        "Direct summary line",
        "",
        "Patient | Alex Wang",
        "Alex Wang",
        "Patient",
        "<",
        "",
        "1_0001",
    ]


def test_parse_sources_prefers_english_fields_in_source_txt(tmp_path: Path) -> None:
    source_docx = tmp_path / 'source.docx'
    source_txt = tmp_path / 'source.txt'

    doc = Document()
    doc.add_paragraph('Docx title')
    doc.add_paragraph('https://example.com/from-docx')
    doc.add_paragraph('Docx summary')
    doc.add_paragraph('( 11/16~17 )')
    doc.save(source_docx)

    source_txt.write_text(
        '\n'.join(
            [
                'TITLE_TEXT: Source title',
                'TITLE_URL: https://example.com/from-source',
                'SUMMARY:',
                'Source summary',
                '',
                'SUPER_PEOPLE:',
                'Patient | Alex Wang',
                'Alex Wang',
                'Patient',
                '',
                'BODY:',
                '1_0001',
                '中文內文。',
                'English body line.',
            ]
        ),
        encoding='utf-8',
    )

    data = generate_news.parse_sources(source_docx, source_txt)

    assert data['TITLE_TEXT'] == 'Source title'
    assert data['TITLE_URL'] == 'https://example.com/from-source'
    assert data['SUMMARY'] == 'Source summary'
