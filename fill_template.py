#!/usr/bin/env python3

from __future__ import annotations

import argparse
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.shared import RGBColor


PLACEHOLDER_KEYS = [
    "TITLE",
    "URL",
    "SUMMARY",
    "YT_TITLE_SUGGESTED",
    "TITLE_SUGGESTED",
    "THUMBNAIL",
    "TIME_RANGE",
    "BODY",
]
PLACEHOLDER_KEY_SET = set(PLACEHOLDER_KEYS)


def parse_input(path: Path) -> dict[str, str]:
    data: dict[str, str] = {}
    body_lines: list[str] = []
    in_body = False

    for raw_line in path.read_text(encoding="utf-8").splitlines():
        if in_body:
            body_lines.append(raw_line)
            continue

        if ":" not in raw_line:
            continue

        key, value = raw_line.split(":", 1)
        key = key.strip().upper()
        value = value.lstrip()

        if key == "BODY":
            in_body = True
            if value:
                body_lines.append(value)
            continue

        if key in PLACEHOLDER_KEY_SET:
            data[key] = value

    if in_body:
        data["BODY"] = "\n".join(body_lines).rstrip()
    else:
        data.setdefault("BODY", "")

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


def fill_template(template_path: Path, input_path: Path, output_path: Path) -> None:
    data = parse_input(input_path)
    doc = Document(str(template_path))
    time_range_style = ensure_time_range_style(doc)
    ensure_hyperlink_style(doc)

    for paragraph in list(doc.paragraphs):
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
