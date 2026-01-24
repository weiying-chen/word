from pathlib import Path

from docx import Document

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
            "--------------------------------",
        ],
    )
    _write_docx(template_path, ["{{HEADER_TITLE}}", "{{HEADER_URL}}", "Body."])
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
    with zipfile.ZipFile(output_path) as zf:
        xml_text = zf.read("word/document.xml")
    ET.fromstring(xml_text)
