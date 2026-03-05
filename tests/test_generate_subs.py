from pathlib import Path
import zipfile
import warnings

import pytest
from docx import Document
from lxml import etree

import generate_subs


def _write_docx(path: Path, paragraphs: list[str]) -> None:
    doc = Document()
    for text in paragraphs:
        doc.add_paragraph(text)
    doc.save(path)


def _assert_title_replaced_for_encoded_input(
    tmp_path: Path,
    encoded_text: bytes,
    expected_title: str,
) -> None:
    template_path = tmp_path / "template.docx"
    input_path = tmp_path / "input.txt"
    output_path = tmp_path / "output.docx"

    _write_docx(template_path, ["{{TITLE}}"])
    input_path.write_bytes(encoded_text)

    generate_subs.generate_subs(template_path, input_path, output_path)

    doc = Document(output_path)
    assert doc.paragraphs[0].text == expected_title


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


def test_title_replacement_utf8_no_bom(tmp_path: Path) -> None:
    _assert_title_replaced_for_encoded_input(
        tmp_path,
        "TITLE: UTF8 Title\n".encode("utf-8"),
        "UTF8 Title",
    )


def test_title_replacement_utf8_bom(tmp_path: Path) -> None:
    _assert_title_replaced_for_encoded_input(
        tmp_path,
        b"\xef\xbb\xbfTITLE: UTF8 BOM Title\n",
        "UTF8 BOM Title",
    )


def test_title_replacement_utf16_le_bom(tmp_path: Path) -> None:
    _assert_title_replaced_for_encoded_input(
        tmp_path,
        "TITLE: UTF16 BOM Title\n".encode("utf-16"),
        "UTF16 BOM Title",
    )


def test_parse_input_fallback_encoding_warns_and_rewrites_utf8(
    tmp_path: Path,
) -> None:
    input_path = tmp_path / "input.txt"
    input_path.write_bytes("TITLE: Café News\n".encode("cp1252"))

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        data = generate_subs.parse_input(input_path)

    assert data["TITLE"] == "Café News"
    assert caught
    assert "fallback encoding" in str(caught[0].message).lower()
    assert "cp1252" in str(caught[0].message).lower()
    assert input_path.read_bytes() == "TITLE: Café News\n".encode("utf-8")


def test_title_replacement_fallback_encoding(tmp_path: Path) -> None:
    with pytest.warns(UserWarning, match="fallback encoding 'cp1252'"):
        _assert_title_replaced_for_encoded_input(
            tmp_path,
            "TITLE: Café from fallback\n".encode("cp1252"),
            "Café from fallback",
        )


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


def test_generate_subs_replaces_box_drawing_horizontal(tmp_path: Path) -> None:
    template_path = tmp_path / "template.docx"
    input_path = tmp_path / "input.txt"
    output_path = tmp_path / "output.docx"

    _write_docx(template_path, ["{{TITLE}}"])
    input_path.write_text(
        "TITLE: 為什麼要蓋醫院─貧中帶病拖垮家庭\n",
        encoding="utf-8",
    )

    generate_subs.generate_subs(template_path, input_path, output_path)
    doc = Document(output_path)
    assert doc.paragraphs[0].text == "為什麼要蓋醫院 - 貧中帶病拖垮家庭"


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


def test_parse_input_multiline_time_range(tmp_path: Path) -> None:
    input_path = tmp_path / "input.txt"
    input_path.write_text(
        "\n".join(
            [
                "TIME_RANGE:",
                "(1) 00:42-05:41 (4m59s)",
                "(2) 05:44-13:21 (7m37s)",
                "19'33",
                "BODY:",
                "dummy",
                "",
            ]
        ),
        encoding="utf-8",
    )

    data = generate_subs.parse_input(input_path)
    assert (
        data["TIME_RANGE"]
        == "(1) 00:42-05:41 (4m59s)\n(2) 05:44-13:21 (7m37s)\n19'33"
    )


def test_generate_subs_applies_time_range_style_in_time_range(tmp_path: Path) -> None:
    template_path = tmp_path / "template.docx"
    input_path = tmp_path / "input.txt"
    output_path = tmp_path / "output.docx"

    _write_docx(template_path, ["{{TIME_RANGE}}"])
    input_path.write_text(
        "\n".join(
            [
                "TIME_RANGE:",
                "(1) 00:42-05:41 (4m59s)",
                "(2) 05:44-13:21 (7m37s)",
                "19'33",
                "BODY:",
                "dummy",
                "",
            ]
        ),
        encoding="utf-8",
    )

    generate_subs.generate_subs(template_path, input_path, output_path)

    with zipfile.ZipFile(output_path) as zf:
        doc = etree.fromstring(zf.read("word/document.xml"))

    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    wanted = {
        "(1) 00:42-05:41 (4m59s)",
        "(2) 05:44-13:21 (7m37s)",
        "19'33",
    }

    found = {}
    for p in doc.findall(".//w:p", ns):
        text = "".join(t.text or "" for t in p.findall(".//w:t", ns)).strip()
        if text in wanted:
            styles = set()
            for run in p.findall("w:r", ns):
                r_style = run.find("w:rPr/w:rStyle", ns)
                if r_style is not None:
                    styles.add(r_style.get("{%s}val" % ns["w"]))
            found[text] = styles

    assert "TimeRange" in found["(1) 00:42-05:41 (4m59s)"]
    assert "TimeRange" in found["(2) 05:44-13:21 (7m37s)"]
    assert "TimeRange" in found["19'33"]


def test_generate_subs_layout_matches_current_output_structure(tmp_path: Path) -> None:
    template_path = Path("templates/subs_template.docx")
    input_path = tmp_path / "input.txt"
    output_path = tmp_path / "output.docx"

    input_path.write_text(
        "\n".join(
            [
                "TITLE: Sample Title",
                "URL: https://example.com/watch?v=1",
                "SUMMARY:",
                "Summary line A.",
                "",
                "Summary line B.",
                "YT_TITLE_SUGGESTED: YT title",
                "TITLE_SUGGESTED: Title only",
                "INTRO:",
                "Intro EN.",
                "",
                "Intro ZH.",
                "THUMBNAIL: missing.png",
                "TIME_RANGE:",
                "(1) 00:42-05:41 (4m59s)",
                "(2) 05:44-13:21 (7m37s)",
                "",
                "19'33",
                "BODY:",
                "00:00:01:00\t00:00:02:00\t測試",
                "Test",
                "",
            ]
        ),
        encoding="utf-8",
    )

    generate_subs.generate_subs(template_path, input_path, output_path)
    doc = Document(output_path)
    texts = [p.text for p in doc.paragraphs]

    # Key structure order should match current output behavior.
    idx_summary_a = texts.index("Summary line A.")
    idx_summary_b = texts.index("Summary line B.")
    idx_yt_label = texts.index("建議YT標題：")
    idx_title_label = texts.index("建議標題：")
    idx_intro_label = texts.index("簡介：")
    idx_thumb_label = texts.index("選圖：")
    idx_subtitle_label = texts.index("字幕：")
    idx_t1 = texts.index("(1) 00:42-05:41 (4m59s)")
    idx_t2 = texts.index("(2) 05:44-13:21 (7m37s)")
    idx_total = texts.index("19'33")
    idx_body_src = texts.index("00:00:01:00\t00:00:02:00\t測試")

    assert idx_summary_a < idx_summary_b < idx_yt_label
    assert idx_yt_label < idx_title_label < idx_intro_label < idx_thumb_label < idx_subtitle_label
    assert idx_subtitle_label < idx_t1 < idx_t2 < idx_total < idx_body_src
