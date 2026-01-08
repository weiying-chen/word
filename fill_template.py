#!/usr/bin/env python3

from __future__ import annotations

import argparse
import re
import zipfile
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


def _get_section_metrics(doc: Document):
    section = doc.sections[0]
    return {
        "page_width": section.page_width,
        "left_margin": section.left_margin,
        "right_margin": section.right_margin,
        "usable_width": section.page_width - section.left_margin - section.right_margin,
    }


def _normalize_document_namespace(xml_text: str) -> str:
    match = re.search(r"<w:document[^>]*>", xml_text)
    if not match:
        return xml_text
    tag = match.group(0)
    if "ns1:Ignorable" not in tag:
        return xml_text
    new_tag = (
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
        'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
        'xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" '
        'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
        'mc:Ignorable="w14 w15 w16se wp14">'
    )
    return xml_text.replace(tag, new_tag, 1)


def fix_docx_namespaces(path: Path) -> None:
    with zipfile.ZipFile(path, "r") as zin:
        xml_text = zin.read("word/document.xml").decode("utf-8")
        fixed_text = _normalize_document_namespace(xml_text)
        if fixed_text == xml_text:
            return

        tmp_path = path.with_suffix(path.suffix + ".tmp")
        with zipfile.ZipFile(tmp_path, "w") as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename == "word/document.xml":
                    data = fixed_text.encode("utf-8")
                zout.writestr(info, data)

    tmp_path.replace(path)


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
    fix_docx_namespaces(output_path)


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
