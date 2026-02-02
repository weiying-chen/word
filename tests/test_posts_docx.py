from pathlib import Path

from docx import Document
from docx.shared import Inches

import zipfile
import xml.etree.ElementTree as ET

from generate_posts import generate_docs


def _write_docx(path: Path, paragraphs: list[str]) -> None:
    doc = Document()
    for text in paragraphs:
        doc.add_paragraph(text)
    doc.save(path)


def test_generated_docx_is_well_formed(tmp_path: Path) -> None:
    schedule_path = tmp_path / "schedule.docx"
    template_path = tmp_path / "template.docx"
    output_dir = tmp_path / "outputs"

    _write_docx(
        schedule_path,
        [
            "節目1則",
            "1. alex",
            "Program - Test Title st/rc",
            "https://example.com/video",
            "搭配",
            "https://example.com/news",
            "News title",
            "--------------------------------",
        ],
    )
    _write_docx(
        template_path,
        [
            "{{HEADER_TITLE}}",
            "{{HEADER_URL}}",
            "{{REF_URL}}",
            "{{REF_TITLE}}",
            "{{VIDEO_URL}}",
            "{{VIDEO_TITLE}}",
        ],
    )
    output_dir.mkdir()

    output_paths = generate_docs(
        schedule_path=schedule_path,
        template_path=template_path,
        output_dir=output_dir,
        filename_prefix="",
        filename_suffix="",
    )
    assert len(output_paths) == 1

    output_path = output_paths[0]
    doc = Document(str(output_path))
    texts = [p.text for p in doc.paragraphs if p.text.strip()]
    assert texts[0] == "Program - Test Title"
    assert texts[1] == "https://example.com/video"
    assert texts[2] == "https://example.com/news"
    assert texts[3] == "News title"
    assert texts[4] == "https://example.com/video"
    assert texts[5] == "Program - Test Title"
    with zipfile.ZipFile(output_path) as zf:
        xml_text = zf.read("word/document.xml")
    ET.fromstring(xml_text)


def test_generated_docx_from_alex_blocks_uses_date_prefix(tmp_path: Path) -> None:
    schedule_path = tmp_path / "alex_blocks.docx"
    template_path = tmp_path / "template.docx"
    output_dir = tmp_path / "outputs"

    _write_docx(
        schedule_path,
        [
            "1",
            "參考資料:",
            "https://example.com/news",
            "26/1/23",
            "News title",
            "要用的影片:",
            "https://example.com/video",
            "Program - Test Title (大愛醫生館 - 測試標題)",
            "English prompt",
            "中文提示",
        ],
    )
    _write_docx(
        template_path,
        [
            "{{HEADER_TITLE}}",
            "{{HEADER_URL}}",
            "{{REF_URL}}",
            "{{REF_TITLE}}",
            "{{VIDEO_URL}}",
            "{{VIDEO_TITLE}}",
            "{{VIDEO_DESC_EN}}",
            "{{VIDEO_DESC_ZH}}",
        ],
    )
    output_dir.mkdir()

    output_paths = generate_docs(
        schedule_path=schedule_path,
        template_path=template_path,
        output_dir=output_dir,
        filename_prefix="日期未定_",
        filename_suffix="",
    )
    assert len(output_paths) == 1

    output_path = output_paths[0]
    assert output_path.name.startswith("260123_")
    doc = Document(str(output_path))
    texts = [p.text for p in doc.paragraphs if p.text.strip()]
    assert texts[0] == "Program - Test Title (大愛醫生館 - 測試標題)"
    assert texts[1] == "https://example.com/video"
    assert texts[2] == "https://example.com/news"
    assert texts[3] == "News title"
    assert texts[4] == "https://example.com/video"
    assert texts[5] == "Program - Test Title (大愛醫生館 - 測試標題)"
    assert texts[6] == "English prompt"
    assert texts[7] == "中文提示"


def test_generated_docx_has_highlighted_ref_hyperlink(tmp_path: Path) -> None:
    schedule_path = tmp_path / "schedule.docx"
    template_path = tmp_path / "template.docx"
    output_dir = tmp_path / "outputs"

    _write_docx(
        schedule_path,
        [
            "節目1則",
            "1. alex",
            "Program - Test Title st/rc",
            "https://example.com/video",
            "搭配",
            "https://example.com/news",
            "News title",
            "--------------------------------",
        ],
    )
    _write_docx(
        template_path,
        [
            "{{REF_URL}}",
        ],
    )
    output_dir.mkdir()

    output_paths = generate_docs(
        schedule_path=schedule_path,
        template_path=template_path,
        output_dir=output_dir,
        filename_prefix="",
        filename_suffix="",
    )
    output_path = output_paths[0]

    with zipfile.ZipFile(output_path) as zf:
        doc = ET.fromstring(zf.read("word/document.xml"))

    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    url_paragraph = None
    for p in doc.findall(".//w:p", ns):
        text = "".join(t.text or "" for t in p.findall(".//w:t", ns))
        if "https://example.com/news" in text:
            url_paragraph = p
            break

    assert url_paragraph is not None
    hyperlinks = url_paragraph.findall("w:hyperlink", ns)
    assert len(hyperlinks) == 1
    h_runs = hyperlinks[0].findall("w:r", ns)
    assert h_runs
    h_rpr = h_runs[0].find("w:rPr", ns)
    h_highlight = h_rpr.find("w:highlight", ns)
    h_size = h_rpr.find("w:sz", ns)
    assert h_highlight.get("{%s}val" % ns["w"]) == "cyan"
    assert h_size.get("{%s}val" % ns["w"]) == "20"


def test_generated_docx_syncs_empty_paragraph_indent(tmp_path: Path) -> None:
    schedule_path = tmp_path / "schedule.docx"
    template_path = tmp_path / "template.docx"
    output_dir = tmp_path / "outputs"

    _write_docx(
        schedule_path,
        [
            "節目1則",
            "1. alex",
            "Program - Test Title st/rc",
            "https://example.com/video",
            "搭配",
            "https://example.com/news",
            "News title",
            "--------------------------------",
        ],
    )
    _write_docx(
        template_path,
        [
            "{{VIDEO_TITLE}}",
            "",
            "{{VIDEO_DESC_EN}}",
        ],
    )
    output_dir.mkdir()

    output_paths = generate_docs(
        schedule_path=schedule_path,
        template_path=template_path,
        output_dir=output_dir,
        filename_prefix="",
        filename_suffix="",
    )
    output_path = output_paths[0]
    doc = Document(str(output_path))

    video_title = doc.paragraphs[0]
    empty_para = doc.paragraphs[1]
    assert video_title.text.strip()
    assert empty_para.text.strip() == ""
    assert empty_para.paragraph_format.left_indent == video_title.paragraph_format.left_indent
    assert empty_para.paragraph_format.first_line_indent == video_title.paragraph_format.first_line_indent
    assert empty_para.paragraph_format.left_indent == Inches(0.5)


def test_generated_docx_sets_source_indent_for_labels(tmp_path: Path) -> None:
    schedule_path = tmp_path / "schedule.docx"
    template_path = tmp_path / "template.docx"
    output_dir = tmp_path / "outputs"

    _write_docx(
        schedule_path,
        [
            "節目1則",
            "1. alex",
            "Program - Test Title st/rc",
            "https://example.com/video",
            "搭配",
            "https://example.com/news",
            "News title",
            "--------------------------------",
        ],
    )
    _write_docx(
        template_path,
        [
            "要用的影片：",
            "{{VIDEO_URL}}",
        ],
    )
    output_dir.mkdir()

    output_paths = generate_docs(
        schedule_path=schedule_path,
        template_path=template_path,
        output_dir=output_dir,
        filename_prefix="",
        filename_suffix="",
    )
    output_path = output_paths[0]
    doc = Document(str(output_path))

    label_para = doc.paragraphs[0]
    blank_para = doc.paragraphs[1]
    url_para = doc.paragraphs[2]
    assert label_para.text.strip() == "要用的影片："
    assert not blank_para.text.strip()
    assert url_para.text.strip() == "https://example.com/video"
    assert label_para.paragraph_format.left_indent == Inches(0.5)
    assert label_para.paragraph_format.first_line_indent == 0
    assert url_para.paragraph_format.left_indent == Inches(0.5)
    assert url_para.paragraph_format.first_line_indent == 0
