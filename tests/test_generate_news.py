from pathlib import Path
import warnings
import zipfile

from docx import Document
from lxml import etree

from docx_utils import add_hyperlink
import generate_news


def _write_source_docx(path: Path) -> None:
    doc = Document()
    doc.add_paragraph("Community Clinic Brings Care to Coastal Town")
    link_paragraph = doc.add_paragraph("")
    add_hyperlink(link_paragraph, "https://example.com/news/story", "https://example.com/news/story")
    doc.add_paragraph("Volunteers organized a two-day clinic to support families in a coastal town.")
    doc.add_paragraph("(  11/16~17 )")
    doc.add_paragraph("")
    doc.add_paragraph("<")
    doc.add_paragraph("old body line")
    doc.save(path)


def test_parse_input_extracts_body_section(tmp_path: Path) -> None:
    input_path = tmp_path / "news_input.txt"
    input_path.write_text(
        "\n".join(
            [
                "TITLE_TEXT: Ignored title",
                "SUMMARY:",
                "Ignored summary",
                "",
                "BODY:",
                "1_0014",
                "中文內文。",
                "English line.",
            ]
        ),
        encoding="utf-8",
    )

    data = generate_news.parse_input(input_path)

    assert data["BODY"] == "1_0014\n中文內文。\nEnglish line."


def test_parse_input_uses_whole_file_when_no_body_label(tmp_path: Path) -> None:
    input_path = tmp_path / "news_input.txt"
    input_path.write_text("1_0014\n中文內文。\nEnglish line.\n", encoding="utf-8")

    data = generate_news.parse_input(input_path)

    assert data["BODY"] == "1_0014\n中文內文。\nEnglish line."


def test_parse_input_fallback_encoding_warns_and_rewrites_utf8(tmp_path: Path) -> None:
    input_path = tmp_path / "news_input.txt"
    input_path.write_bytes("BODY: Café News\n".encode("cp1252"))

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        data = generate_news.parse_input(input_path)

    assert data["BODY"] == "Café News"
    assert caught
    assert "fallback encoding" in str(caught[0].message).lower()
    assert "cp1252" in str(caught[0].message).lower()
    assert input_path.read_bytes() == "BODY: Café News\n".encode("utf-8")


def test_generate_news_preserves_header_and_replaces_body(tmp_path: Path) -> None:
    source_docx = tmp_path / "source.docx"
    input_path = tmp_path / "news_input.txt"
    output_path = tmp_path / "news_output.docx"

    _write_source_docx(source_docx)
    input_path.write_text(
        "\n".join(
            [
                "BODY:",
                "1_0014",
                "Two days of free screenings brought steady lines of local residents.",
                "",
                "2_0025",
                "Residents arrived early to receive free screenings and follow-up advice.",
            ]
        ),
        encoding="utf-8",
    )

    generate_news.generate_news(source_docx, input_path, output_path)

    doc = Document(output_path)
    texts = [p.text for p in doc.paragraphs]
    assert texts == [
        "Community Clinic Brings Care to Coastal Town",
        "https://example.com/news/story",
        "Volunteers organized a two-day clinic to support families in a coastal town.",
        "(  11/16~17 )",
        "",
        "<",
        "",
        "1_0014",
        "Two days of free screenings brought steady lines of local residents.",
        "",
        "2_0025",
        "Residents arrived early to receive free screenings and follow-up advice.",
    ]
    assert "old body line" not in texts
    assert doc.paragraphs[7].runs[0].font.highlight_color.name == "BRIGHT_GREEN"
    assert doc.paragraphs[10].runs[0].font.highlight_color.name == "BRIGHT_GREEN"

    with zipfile.ZipFile(output_path) as zf:
        document_xml = etree.fromstring(zf.read("word/document.xml"))

    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    paragraphs = document_xml.findall(".//w:body/w:p", ns)
    assert paragraphs
    hyperlinks = paragraphs[1].findall("w:hyperlink", ns)
    assert len(hyperlinks) == 1


def test_generate_news_from_sources_uses_body_text_file(tmp_path: Path) -> None:
    source_docx = tmp_path / "source.docx"
    source_txt = tmp_path / "source.txt"
    output_path = tmp_path / "news_output.docx"

    _write_source_docx(source_docx)
    source_txt.write_text("1_0001\n中文內文。\nEnglish body line.\n", encoding="utf-8")

    generate_news.generate_news_from_sources(source_docx, source_txt, output_path)

    texts = [p.text for p in Document(output_path).paragraphs]
    assert texts[-5:] == ["<", "", "1_0001", "中文內文。", "English body line."]


def test_generate_news_requires_marker_in_source_docx(tmp_path: Path) -> None:
    source_docx = tmp_path / "source.docx"
    input_path = tmp_path / "news_input.txt"
    output_path = tmp_path / "news_output.docx"

    doc = Document()
    doc.add_paragraph("Header without marker")
    doc.save(source_docx)
    input_path.write_text("BODY:\n1_0001\nBody line.\n", encoding="utf-8")

    try:
        generate_news.generate_news(source_docx, input_path, output_path)
    except ValueError as exc:
        assert "<" in str(exc)
    else:
        raise AssertionError("Expected generate_news to require a marker paragraph")


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
