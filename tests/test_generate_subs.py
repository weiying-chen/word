from pathlib import Path
import zipfile

from docx import Document
from lxml import etree

import generate_subs


def _write_docx(path: Path, paragraphs: list[str]) -> None:
    doc = Document()
    for text in paragraphs:
        doc.add_paragraph(text)
    doc.save(path)


def test_parse_input_multiline_summary(tmp_path: Path) -> None:
    input_path = tmp_path / "input.txt"
    input_path.write_text(
        "\n".join(
            [
                "TITLE: Sample Title",
                "SUMMARY:",
                "First summary line.",
                "Second summary line.",
                "TITLE_SUGGESTED: Another Title",
                "",
            ]
        ),
        encoding="utf-8",
    )

    data = generate_subs.parse_input(input_path)
    assert data["SUMMARY"] == "First summary line.\nSecond summary line."


def test_parse_input_multiline_intro_body(tmp_path: Path) -> None:
    input_path = tmp_path / "input.txt"
    input_path.write_text(
        "\n".join(
            [
                "INTRO:",
                "Intro line one.",
                "Intro line two.",
                "BODY:",
                "Body line one.",
                "Body line two.",
                "",
            ]
        ),
        encoding="utf-8",
    )

    data = generate_subs.parse_input(input_path)
    assert data["INTRO"] == "Intro line one.\nIntro line two."
    assert data["BODY"] == "Body line one.\nBody line two."


def test_generate_subs_removes_empty_summary_paragraph(tmp_path: Path) -> None:
    template_path = tmp_path / "template.docx"
    input_path = tmp_path / "input.txt"
    output_path = tmp_path / "output.docx"

    _write_docx(template_path, ["{{SUMMARY}}", "After summary."])
    input_path.write_text("TITLE: Sample Title\n", encoding="utf-8")

    generate_subs.generate_subs(template_path, input_path, output_path)
    doc = Document(output_path)
    texts = [p.text for p in doc.paragraphs if p.text]
    assert texts == ["After summary."]


def test_generate_subs_removes_empty_time_range_paragraph(tmp_path: Path) -> None:
    template_path = tmp_path / "template.docx"
    input_path = tmp_path / "input.txt"
    output_path = tmp_path / "output.docx"

    _write_docx(template_path, ["{{TIME_RANGE}}", "After time range."])
    input_path.write_text("TITLE: Sample Title\n", encoding="utf-8")

    generate_subs.generate_subs(template_path, input_path, output_path)
    doc = Document(output_path)
    texts = [p.text for p in doc.paragraphs if p.text]
    assert texts == ["After time range."]


def test_generate_subs_inserts_blank_after_labels(tmp_path: Path) -> None:
    template_path = tmp_path / "template.docx"
    input_path = tmp_path / "input.txt"
    output_path = tmp_path / "output.docx"

    _write_docx(template_path, ["簡介：", "Line after label."])
    input_path.write_text("TITLE: Sample Title\n", encoding="utf-8")

    generate_subs.generate_subs(template_path, input_path, output_path)
    doc = Document(output_path)
    assert doc.paragraphs[0].text.strip() == "簡介："
    assert not doc.paragraphs[1].text.strip()
    assert doc.paragraphs[2].text.strip() == "Line after label."


def test_generate_subs_inserts_blank_after_last_label(tmp_path: Path) -> None:
    template_path = tmp_path / "template.docx"
    input_path = tmp_path / "input.txt"
    output_path = tmp_path / "output.docx"

    _write_docx(template_path, ["字幕："])
    input_path.write_text("TITLE: Sample Title\n", encoding="utf-8")

    generate_subs.generate_subs(template_path, input_path, output_path)
    doc = Document(output_path)
    assert doc.paragraphs[0].text.strip() == "字幕："
    assert not doc.paragraphs[1].text.strip()


def test_source_block_highlight_and_link(tmp_path: Path) -> None:
    template_path = tmp_path / "template.docx"
    input_path = tmp_path / "input.txt"
    output_path = tmp_path / "output.docx"

    _write_docx(template_path, ["{{BODY}}"])
    input_path.write_text(
        "\n".join(
            [
                "BODY:",
                "00:00:00:00\t00:00:02:00\tFirst line.",
                "First translation.",
                "",
                "https://example.com/source",
                "Source line one.",
                "",
                "00:00:02:00\t00:00:04:00\tSecond line.",
                "Second translation.",
                "",
            ]
        ),
        encoding="utf-8",
    )

    generate_subs.generate_subs(template_path, input_path, output_path)

    with zipfile.ZipFile(output_path) as zf:
        doc = etree.fromstring(zf.read("word/document.xml"))

    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    url_paragraph = None
    source_paragraph = None
    for p in doc.findall(".//w:p", ns):
        text = "".join(t.text or "" for t in p.findall(".//w:t", ns))
        if "https://example.com/source" in text:
            url_paragraph = p
        if "Source line one." in text:
            source_paragraph = p

    assert url_paragraph is not None
    assert source_paragraph is not None

    hyperlinks = url_paragraph.findall("w:hyperlink", ns)
    assert len(hyperlinks) == 1

    h_runs = hyperlinks[0].findall("w:r", ns)
    assert h_runs
    h_rpr = h_runs[0].find("w:rPr", ns)
    h_highlight = h_rpr.find("w:highlight", ns)
    h_size = h_rpr.find("w:sz", ns)
    assert h_highlight.get("{%s}val" % ns["w"]) == "cyan"
    assert h_size.get("{%s}val" % ns["w"]) == "20"

    s_run = None
    for r in source_paragraph.findall("w:r", ns):
        text = "".join(t.text or "" for t in r.findall(".//w:t", ns))
        if "Source line one." in text:
            s_run = r
            break

    assert s_run is not None
    s_rpr = s_run.find("w:rPr", ns)
    s_highlight = s_rpr.find("w:highlight", ns)
    s_size = s_rpr.find("w:sz", ns)
    assert s_highlight.get("{%s}val" % ns["w"]) == "cyan"
    assert s_size.get("{%s}val" % ns["w"]) == "20"
