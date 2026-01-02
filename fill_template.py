#!/usr/bin/env python3

from __future__ import annotations

import argparse
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.shared import RGBColor, Inches


PLACEHOLDER_KEYS = [
    "TITLE",
    "URL",
    "SUMMARY",
    "YT_TITLE_SUGGESTED",
    "TITLE_SUGGESTED",
    "INTRO",
    "THUMBNAIL",
    "TIME_RANGE",
    "BODY",
]
PLACEHOLDER_KEY_SET = set(PLACEHOLDER_KEYS)


def parse_input(path: Path) -> dict[str, str]:
    data: dict[str, str] = {}
    lines = path.read_text(encoding="utf-8").splitlines()
    idx = 0
    while idx < len(lines):
        raw_line = lines[idx]
        if ":" not in raw_line:
            idx += 1
            continue

        key, value = raw_line.split(":", 1)
        key = key.strip().upper()
        value = value.lstrip()

        if key not in PLACEHOLDER_KEY_SET:
            idx += 1
            continue

        if key in {"BODY", "INTRO"}:
            collected: list[str] = []
            if value:
                collected.append(value)
            idx += 1
            while idx < len(lines):
                next_line = lines[idx]
                if ":" in next_line:
                    next_key = next_line.split(":", 1)[0].strip().upper()
                    if next_key in PLACEHOLDER_KEY_SET:
                        break
                collected.append(next_line)
                idx += 1
            data[key] = "\n".join(collected).rstrip()
            continue

        data[key] = value
        idx += 1

    data.setdefault("BODY", "")
    data.setdefault("INTRO", "")

    return data


def replace_placeholder(paragraph, placeholder: str, value: str) -> bool:
    if placeholder not in paragraph.text:
        return False

    paragraph.text = paragraph.text.replace(placeholder, value)
    return True


def insert_paragraph_after(paragraph, text: str):
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    new_para.add_run(text)
    return new_para


def replace_body_paragraph(paragraph, body_text: str) -> None:
    lines = body_text.splitlines() if body_text else []
    paragraph.text = ""
    if not lines:
        return

    paragraph.add_run(lines[0])
    current = paragraph
    for line in lines[1:]:
        current = insert_paragraph_after(current, line)


def remove_paragraph(paragraph) -> None:
    element = paragraph._element
    element.getparent().remove(element)


def clear_paragraph(paragraph) -> None:
    for run in paragraph.runs:
        run._element.getparent().remove(run._element)


def add_hyperlink(paragraph, text: str, url: str) -> None:
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
    r_underline = OxmlElement("w:u")
    r_underline.set(qn("w:val"), "single")
    r_pr.append(r_underline)
    run.append(r_pr)

    t = OxmlElement("w:t")
    t.text = text
    run.append(t)

    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def ensure_time_range_style(doc: Document):
    style_name = "TimeRange"
    styles = doc.styles
    if style_name in [s.name for s in styles]:
        style = styles[style_name]
    else:
        style = styles.add_style(style_name, WD_STYLE_TYPE.CHARACTER)
    style.font.color.rgb = RGBColor(0x05, 0x63, 0xC1)
    return style_name


def ensure_hyperlink_style(doc: Document):
    style_name = "Hyperlink"
    styles = doc.styles
    if style_name in [s.name for s in styles]:
        style = styles[style_name]
    else:
        style = styles.add_style(style_name, WD_STYLE_TYPE.CHARACTER)
    style.font.color.rgb = RGBColor(0x05, 0x63, 0xC1)
    style.font.underline = True
    return style_name


def convert_inline_to_anchor(inline, align: str = "left") -> None:
    anchor = OxmlElement("wp:anchor")
    anchor.set("behindDoc", "0")
    anchor.set("distT", "0")
    anchor.set("distB", "0")
    anchor.set("distL", "0")
    anchor.set("distR", "0")
    anchor.set("simplePos", "0")
    anchor.set("locked", "0")
    anchor.set("layoutInCell", "0")
    anchor.set("allowOverlap", "1")
    anchor.set("relativeHeight", "2")

    simple_pos = OxmlElement("wp:simplePos")
    simple_pos.set("x", "0")
    simple_pos.set("y", "0")
    anchor.append(simple_pos)

    pos_h = OxmlElement("wp:positionH")
    pos_h.set("relativeFrom", "column")
    pos_h_align = OxmlElement("wp:align")
    pos_h_align.text = align
    pos_h.append(pos_h_align)
    anchor.append(pos_h)

    pos_v = OxmlElement("wp:positionV")
    pos_v.set("relativeFrom", "paragraph")
    pos_v_offset = OxmlElement("wp:posOffset")
    pos_v_offset.text = "0"
    pos_v.append(pos_v_offset)
    anchor.append(pos_v)

    extent = inline.find(qn("wp:extent"))
    if extent is not None:
        anchor.append(extent)

    effect_extent = inline.find(qn("wp:effectExtent"))
    if effect_extent is None:
        effect_extent = OxmlElement("wp:effectExtent")
        effect_extent.set("l", "0")
        effect_extent.set("t", "0")
        effect_extent.set("r", "0")
        effect_extent.set("b", "0")
    anchor.append(effect_extent)

    wrap = OxmlElement("wp:wrapSquare")
    wrap.set("wrapText", "largest")
    anchor.append(wrap)

    doc_pr = inline.find(qn("wp:docPr"))
    if doc_pr is not None:
        anchor.append(doc_pr)

    c_nv = inline.find(qn("wp:cNvGraphicFramePr"))
    if c_nv is not None:
        anchor.append(c_nv)

    graphic = inline.find(qn("a:graphic"))
    if graphic is not None:
        anchor.append(graphic)

    drawing = inline.getparent()
    drawing.remove(inline)
    drawing.append(anchor)


def _get_section_metrics(doc: Document):
    section = doc.sections[0]
    return {
        "page_width": section.page_width,
        "left_margin": section.left_margin,
        "right_margin": section.right_margin,
        "usable_width": section.page_width - section.left_margin - section.right_margin,
    }


def fill_template(template_path: Path, input_path: Path, output_path: Path) -> None:
    data = parse_input(input_path)
    input_base = input_path.parent
    doc = Document(str(template_path))
    time_range_style = ensure_time_range_style(doc)
    ensure_hyperlink_style(doc)
    metrics = _get_section_metrics(doc)

    for paragraph in list(doc.paragraphs):
        if "{{INTRO}}" in paragraph.text:
            intro = data.get("INTRO", "")
            if not intro and paragraph.text.strip() == "{{INTRO}}":
                remove_paragraph(paragraph)
                continue
            replace_body_paragraph(paragraph, intro)
            continue
        if "{{BODY}}" in paragraph.text:
            replace_body_paragraph(paragraph, data.get("BODY", ""))
            continue

        for key in PLACEHOLDER_KEYS:
            placeholder = f"{{{{{key}}}}}"
            if placeholder in paragraph.text:
                value = data.get(key, "")
                if not value and paragraph.text.strip() == placeholder:
                    remove_paragraph(paragraph)
                    break
                if key == "URL" and value:
                    clear_paragraph(paragraph)
                    add_hyperlink(paragraph, value, value)
                elif key == "THUMBNAIL" and value:
                    thumbnail_path = Path(value)
                    if not thumbnail_path.is_absolute():
                        thumbnail_path = input_base / thumbnail_path
                    if thumbnail_path.exists():
                        clear_paragraph(paragraph)
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        paragraph.paragraph_format.left_indent = 0
                        paragraph.paragraph_format.right_indent = 0
                        paragraph.paragraph_format.first_line_indent = 0
                        run = paragraph.add_run()
                        run.add_picture(str(thumbnail_path), width=metrics["usable_width"])
                        inline = paragraph._p.find(
                            ".//wp:inline",
                            namespaces={
                                "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                            },
                        )
                        if inline is not None:
                            convert_inline_to_anchor(inline, align="left")
                    else:
                        replace_placeholder(paragraph, placeholder, value)
                elif key == "TIME_RANGE" and value:
                    clear_paragraph(paragraph)
                    run = paragraph.add_run(value)
                    run.style = time_range_style
                else:
                    replace_placeholder(paragraph, placeholder, value)
                break

    doc.save(str(output_path))


def main() -> None:
    parser = argparse.ArgumentParser(description="Fill a DOCX template from a text file.")
    parser.add_argument(
        "--template",
        default="template.docx",
        help="Path to the DOCX template.",
    )
    parser.add_argument(
        "--input",
        default="input.txt",
        help="Path to the input text file.",
    )
    parser.add_argument(
        "--output",
        default="output.docx",
        help="Path to write the filled DOCX.",
    )
    args = parser.parse_args()

    fill_template(Path(args.template), Path(args.input), Path(args.output))


if __name__ == "__main__":
    main()
