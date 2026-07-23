from docx import Document
from docx.oxml.ns import qn

from generate_sources import _add_text_with_symbol_runs


def test_uses_symbol_font_only_for_symbol_character() -> None:
    document = Document()
    paragraph = document.add_paragraph("")

    _add_text_with_symbol_runs(paragraph, "摘要\n\n➯5分鐘動起來！")

    arrow_run = next(run for run in paragraph.runs if "➯" in run.text)
    text_run = next(run for run in paragraph.runs if "摘要" in run.text)
    arrow_fonts = arrow_run._element.find("w:rPr/w:rFonts", arrow_run._element.nsmap)
    text_fonts = text_run._element.find("w:rPr/w:rFonts", text_run._element.nsmap)

    assert arrow_fonts.get(qn("w:ascii")) == "Segoe UI Symbol"
    assert arrow_fonts.get(qn("w:hAnsi")) == "Segoe UI Symbol"
    assert text_fonts.get(qn("w:eastAsia")) == "新細明體"
