from __future__ import annotations

from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.shared import Inches, Pt
from docx.text.paragraph import Paragraph


def get_default_tab_stop_inches(doc) -> float:
    settings = doc.part.settings.element
    node = settings.find(qn("w:defaultTabStop"))
    if node is None:
        return 0.5
    value = node.get(qn("w:val"))
    if not value:
        return 0.5
    return int(value) / 1440


def clear_paragraph(paragraph) -> None:
    for run in paragraph.runs:
        run._element.getparent().remove(run._element)


def set_source_indent(paragraph, indent_inches: float) -> None:
    paragraph.paragraph_format.left_indent = Inches(indent_inches)
    paragraph.paragraph_format.first_line_indent = 0


def ensure_blank_after_labels(doc, labels: set[str]) -> None:
    for para in list(doc.paragraphs):
        if para.text.strip() not in labels:
            continue
        next_elm = para._p.getnext()
        while next_elm is not None and next_elm.tag != qn("w:p"):
            next_elm = next_elm.getnext()
        if next_elm is None:
            new_p = OxmlElement("w:p")
            para._p.addnext(new_p)
            Paragraph(new_p, para._parent)
            continue
        next_para = Paragraph(next_elm, para._parent)
        if next_para.text.strip():
            new_p = OxmlElement("w:p")
            para._p.addnext(new_p)
            Paragraph(new_p, para._parent)


def add_highlighted_run(
    paragraph,
    text: str,
    *,
    font_size_pt: int = 10,
    highlight_color=WD_COLOR_INDEX.TURQUOISE,
):
    run = paragraph.add_run(text)
    run.font.size = Pt(font_size_pt)
    run.font.highlight_color = highlight_color
    return run


def apply_highlight_to_runs(
    paragraph,
    *,
    font_size_pt: int = 10,
    highlight_color=WD_COLOR_INDEX.TURQUOISE,
) -> None:
    for run in paragraph.runs:
        run.font.size = Pt(font_size_pt)
        run.font.highlight_color = highlight_color


def add_hyperlink(
    paragraph,
    text: str,
    url: str,
    *,
    highlight: bool = False,
    highlight_color: str = "cyan",
    size: int = 20,
) -> None:
    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")
    r_style = OxmlElement("w:rStyle")
    r_style.set(qn("w:val"), "Hyperlink")
    r_pr.append(r_style)
    r_color = OxmlElement("w:color")
    r_color.set(qn("w:val"), "0563C1")
    r_pr.append(r_color)
    if highlight:
        r_highlight = OxmlElement("w:highlight")
        r_highlight.set(qn("w:val"), highlight_color)
        r_pr.append(r_highlight)
        r_sz = OxmlElement("w:sz")
        r_sz.set(qn("w:val"), str(size))
        r_pr.append(r_sz)
    r_underline = OxmlElement("w:u")
    r_underline.set(qn("w:val"), "single")
    r_pr.append(r_underline)
    run.append(r_pr)

    t = OxmlElement("w:t")
    t.text = text
    run.append(t)

    hyperlink.append(run)
    paragraph._p.append(hyperlink)
