#!/usr/bin/env python3

from __future__ import annotations

import argparse
from copy import deepcopy
import re
import unicodedata
import warnings
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.text import WD_COLOR_INDEX
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import RGBColor, Inches, Pt
from docx.text.paragraph import Paragraph

from docx_utils import (
    add_hyperlink,
    clear_paragraph,
    ensure_blank_after_labels,
    get_default_tab_stop_inches,
    set_source_indent,
)


PLACEHOLDER_KEYS = [
    "YT_TITLE_SUGGESTED",
    "TITLE_SUGGESTED",
    "INTRO",
    "THUMBNAIL",
    "THUMBNAIL_CREDIT",
]
PLACEHOLDER_KEY_SET = set(PLACEHOLDER_KEYS)
INPUT_KEY_SET = PLACEHOLDER_KEY_SET | {"BODY"}
GENERIC_PLACEHOLDER_RE = re.compile(r"^\{\{([A-Z_]+)\}\}$")
SOURCE_LINK_RE = re.compile(r"^https?://\S+")
SUBTITLE_LINE_RE = re.compile(
    r"^(?:[^\t]+\t)?\d{2}:\d{2}:\d{2}:\d{2}\t\d{2}:\d{2}:\d{2}:\d{2}\t"
)
SYMBOL_FONT_NAME = "Segoe UI Symbol"
CJK_FONT_NAME = "新細明體"
CJK_MIDDLE_DOT = "\u2027"
HIGHLIGHT_MARKER_RE = re.compile(r"\*([^*]+)\*")
SOURCE_HIGHLIGHT_DEFAULT = WD_COLOR_INDEX.TURQUOISE
SOURCE_HIGHLIGHT_MARKED = WD_COLOR_INDEX.BRIGHT_GREEN
SOURCE_HYPERLINK_HIGHLIGHT_MARKED = "brightGreen"
BOX_DRAWING_HORIZONTAL = "\u2500"
SPACED_HYPHEN_MINUS = " - "
SUBS_OUTPUT_SUFFIX = "_al"
SUBTITLE_LABELS = {"字幕：", "字幕:"}
SECTION_LABELS = {
    "建議YT標題：",
    "建議YT標題:",
    "建議標題：",
    "建議標題:",
    "簡介：",
    "簡介:",
    "選圖：",
    "選圖:",
    *SUBTITLE_LABELS,
}


def normalize_input_text(text: str) -> str:
    if not text:
        return text
    return text.replace(BOX_DRAWING_HORIZONTAL, SPACED_HYPHEN_MINUS)


def _normalized_paragraph_text(text: str) -> str:
    return text.lstrip("\ufeff").lstrip()


def normalize_title_text(text: str) -> str:
    if not text:
        return text
    normalized = normalize_input_text(text)
    return re.sub(r"[ \t]+", " ", normalized).strip()


def _decode_input_text(path: Path) -> tuple[str, str, bool]:
    raw = path.read_bytes()
    tried_encodings: list[str] = []

    if raw.startswith(b"\xef\xbb\xbf"):
        return raw.decode("utf-8-sig"), "utf-8-sig", False

    if raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff"):
        return raw.decode("utf-16"), "utf-16", False

    for encoding in ("utf-8", "big5", "cp950", "gb18030", "cp1252"):
        tried_encodings.append(encoding)
        try:
            return raw.decode(encoding), encoding, encoding != "utf-8"
        except UnicodeDecodeError:
            continue

    tried = ", ".join(tried_encodings)
    raise UnicodeError(
        f"Unable to decode input file '{path}' with supported encodings: {tried}"
    )


def parse_input(path: Path) -> dict[str, str]:
    data: dict[str, str] = {}
    text, encoding_used, used_fallback = _decode_input_text(path)
    if used_fallback:
        warnings.warn(
            f"Using fallback encoding '{encoding_used}' for {path}; rewriting as UTF-8.",
            stacklevel=2,
        )
        path.write_text(text, encoding="utf-8")

    lines = text.splitlines()
    idx = 0
    while idx < len(lines):
        raw_line = lines[idx]
        if ":" not in raw_line:
            idx += 1
            continue

        key, value = raw_line.split(":", 1)
        key = key.lstrip("\ufeff").strip().upper()
        value = value.lstrip()

        if key not in INPUT_KEY_SET:
            idx += 1
            continue

        if key in {"INTRO", "BODY"}:
            collected: list[str] = []
            if value:
                collected.append(value)
            idx += 1
            while idx < len(lines):
                next_line = lines[idx]
                if ":" in next_line:
                    next_key = next_line.split(":", 1)[0].strip().upper()
                    if next_key in INPUT_KEY_SET:
                        break
                collected.append(next_line)
                idx += 1
            data[key] = normalize_input_text("\n".join(collected).rstrip())
            continue

        if key in {"TITLE_SUGGESTED", "YT_TITLE_SUGGESTED"}:
            data[key] = normalize_title_text(value)
        else:
            data[key] = normalize_input_text(value)
        idx += 1

    data.setdefault("INTRO", "")
    data.setdefault("BODY", "")

    return data


def _run_contains_symbol(text: str) -> bool:
    return any(unicodedata.category(char) == "So" for char in text)


def _set_run_font(run, font_name: str) -> None:
    run.font.name = font_name
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.get_or_add_rFonts()
    r_fonts.set(qn("w:ascii"), font_name)
    r_fonts.set(qn("w:hAnsi"), font_name)
    r_fonts.set(qn("w:cs"), font_name)


def apply_symbol_font(run) -> None:
    if not SYMBOL_FONT_NAME:
        return
    if run.text and _run_contains_symbol(run.text):
        _set_run_font(run, SYMBOL_FONT_NAME)


def apply_cjk_middle_dot_font(run) -> None:
    if not CJK_FONT_NAME:
        return
    if run.text and CJK_MIDDLE_DOT in run.text:
        _set_run_font(run, CJK_FONT_NAME)


def apply_symbol_fonts_in_paragraph(paragraph) -> None:
    for run in paragraph.runs:
        apply_symbol_font(run)
        apply_cjk_middle_dot_font(run)


def replace_placeholder(paragraph, placeholder: str, value: str) -> bool:
    if placeholder not in paragraph.text:
        return False

    paragraph.text = paragraph.text.replace(placeholder, value)
    apply_symbol_fonts_in_paragraph(paragraph)
    return True


def insert_paragraph_after(paragraph, text: str):
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    if text:
        run = new_para.add_run(text)
        apply_symbol_font(run)
        apply_cjk_middle_dot_font(run)
    return new_para


def _split_marked_parts(text: str) -> list[tuple[str, bool]]:
    parts: list[tuple[str, bool]] = []
    last_idx = 0

    for match in HIGHLIGHT_MARKER_RE.finditer(text):
        if match.start() > last_idx:
            parts.append((text[last_idx : match.start()], False))
        parts.append((match.group(1), True))
        last_idx = match.end()

    if last_idx < len(text):
        parts.append((text[last_idx:], False))

    if not parts:
        parts.append((text, False))

    return parts


def _split_symbol_chunks(text: str) -> list[tuple[str, bool]]:
    if not text:
        return []

    parts: list[tuple[str, bool]] = []
    start = 0
    current_is_symbol = unicodedata.category(text[0]) == "So"

    for idx in range(1, len(text)):
        is_symbol = unicodedata.category(text[idx]) == "So"
        if is_symbol != current_is_symbol:
            parts.append((text[start:idx], current_is_symbol))
            start = idx
            current_is_symbol = is_symbol
    parts.append((text[start:], current_is_symbol))
    return parts


def _add_text_runs(
    paragraph,
    text: str,
    *,
    run_style: str | None = None,
    highlight_color=None,
    font_size_pt: int | None = None,
) -> None:
    for chunk, is_symbol in _split_symbol_chunks(text):
        if not chunk:
            continue
        run = paragraph.add_run(chunk)
        if run_style:
            run.style = run_style
        if font_size_pt is not None:
            run.font.size = Pt(font_size_pt)
        if highlight_color is not None:
            run.font.highlight_color = highlight_color
        if is_symbol:
            apply_symbol_font(run)
        else:
            apply_cjk_middle_dot_font(run)


def _add_marked_runs(
    paragraph,
    text: str,
    *,
    default_highlight=None,
    marked_highlight=None,
    run_style: str | None = None,
    apply_default_size: bool = True,
) -> None:
    for part_text, marked in _split_marked_parts(text):
        if not part_text:
            continue

        highlight_color = marked_highlight if marked else default_highlight
        _add_text_runs(
            paragraph,
            part_text,
            run_style=run_style,
            highlight_color=highlight_color,
            font_size_pt=10 if apply_default_size else None,
        )


def replace_body_paragraph(
    paragraph,
    body_text: str,
    source_indent_inches: float,
    timing_style: str | None = None,
    marked_highlight=None,
) -> None:
    lines = body_text.splitlines() if body_text else []
    clear_paragraph(paragraph)
    if not lines:
        return
    while lines and not lines[0].strip():
        lines.pop(0)
    if not lines:
        return

    current = paragraph
    in_source_block = False

    def _add_source_runs(target, text: str) -> None:
        _add_marked_runs(
            target,
            text,
            default_highlight=SOURCE_HIGHLIGHT_DEFAULT,
            marked_highlight=SOURCE_HIGHLIGHT_MARKED,
        )

    def write_line(target, text: str, source_line: bool, is_url: bool) -> None:
        if not text:
            return
        if source_line:
            set_source_indent(target, source_indent_inches)
            if is_url:
                add_hyperlink(
                    target,
                    text,
                    text,
                    highlight=True,
                    highlight_color="cyan",
                )
                return
            _add_source_runs(target, text)
        else:
            if timing_style or marked_highlight is not None:
                _add_marked_runs(
                    target,
                    text,
                    marked_highlight=marked_highlight,
                    run_style=timing_style,
                    apply_default_size=False,
                )
            else:
                _add_text_runs(target, text)

    for idx, line in enumerate(lines):
        if idx > 0:
            current = insert_paragraph_after(current, "")

        normalized_line = _normalized_paragraph_text(line)
        if SUBTITLE_LINE_RE.match(normalized_line):
            in_source_block = False

        cleaned_line = HIGHLIGHT_MARKER_RE.sub(r"\1", line)
        is_link = SOURCE_LINK_RE.match(cleaned_line)
        if is_link:
            in_source_block = True

        write_line(
            current,
            cleaned_line if is_link else line,
            in_source_block and not SUBTITLE_LINE_RE.match(normalized_line),
            bool(is_link),
        )


def remove_paragraph(paragraph) -> None:
    element = paragraph._element
    element.getparent().remove(element)


def ensure_annotation_style(doc: Document):
    style_name = "Annotation"
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


def apply_default_margins(doc: Document) -> None:
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.25)
        section.right_margin = Inches(1.25)


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


def with_subs_output_suffix(path: Path) -> Path:
    if path.stem.endswith(SUBS_OUTPUT_SUFFIX):
        return path
    return path.with_name(f"{path.stem}{SUBS_OUTPUT_SUFFIX}{path.suffix}")


def _extract_source_paragraphs(
    source_docx_path: Path,
) -> tuple[Document, list[Paragraph], list[Paragraph]]:
    source_doc = Document(str(source_docx_path))
    paragraphs = list(source_doc.paragraphs)

    for idx, paragraph in enumerate(paragraphs):
        if SUBTITLE_LINE_RE.match(_normalized_paragraph_text(paragraph.text)):
            return source_doc, paragraphs[:idx], paragraphs[idx:]

    raise ValueError("source.docx must contain at least one subtitle timestamp paragraph.")


def _remap_relationship_ids(source_part, target_part, paragraph_element) -> None:
    for element in paragraph_element.iter():
        relation_id = element.get(qn("r:id"))
        if not relation_id:
            continue
        relation = source_part.rels[relation_id]
        if relation.reltype != RT.HYPERLINK:
            continue
        new_relation_id = target_part.relate_to(
            relation.target_ref,
            RT.HYPERLINK,
            is_external=True,
        )
        element.set(qn("r:id"), new_relation_id)


def _remap_hyperlink_run_styles(target_paragraph: Paragraph, paragraph_element) -> None:
    try:
        hyperlink_style_id = target_paragraph.part.document.styles["Hyperlink"].style_id
    except KeyError:
        return

    def set_run_style(run_element) -> None:
        run_properties = run_element.find(qn("w:rPr"))
        if run_properties is None:
            run_properties = OxmlElement("w:rPr")
            run_element.insert(0, run_properties)
        run_style = run_properties.find(qn("w:rStyle"))
        if run_style is None:
            run_style = OxmlElement("w:rStyle")
            run_properties.insert(0, run_style)
        run_style.set(qn("w:val"), hyperlink_style_id)

    for hyperlink in paragraph_element.iter(qn("w:hyperlink")):
        for run in hyperlink.iter(qn("w:r")):
            set_run_style(run)

    # Also handle field-code hyperlinks in the form:
    # fldChar(begin) + instrText(HYPERLINK ...) + fldChar(separate) + display runs + fldChar(end)
    in_field = False
    in_result = False
    hyperlink_field = False
    instruction_text = ""
    for child in paragraph_element:
        if child.tag != qn("w:r"):
            continue

        fld_char = child.find(qn("w:fldChar"))
        if fld_char is not None:
            fld_char_type = fld_char.get(qn("w:fldCharType"))
            if fld_char_type == "begin":
                in_field = True
                in_result = False
                hyperlink_field = False
                instruction_text = ""
            elif fld_char_type == "separate" and in_field:
                in_result = True
                hyperlink_field = "HYPERLINK" in instruction_text.upper()
            elif fld_char_type == "end":
                in_field = False
                in_result = False
                hyperlink_field = False
                instruction_text = ""
            continue

        if not in_field:
            continue

        if in_result:
            if hyperlink_field:
                set_run_style(child)
            continue

        for instr_text in child.findall(qn("w:instrText")):
            instruction_text += instr_text.text or ""


def _clone_paragraph_before(source_paragraph: Paragraph, target_paragraph: Paragraph) -> None:
    cloned = deepcopy(source_paragraph._p)
    _remap_relationship_ids(source_paragraph.part, target_paragraph.part, cloned)
    _remap_hyperlink_run_styles(target_paragraph, cloned)
    target_paragraph._p.addprevious(cloned)


def _clone_paragraph_after(source_paragraph: Paragraph, target_paragraph: Paragraph) -> Paragraph:
    cloned = deepcopy(source_paragraph._p)
    _remap_relationship_ids(source_paragraph.part, target_paragraph.part, cloned)
    _remap_hyperlink_run_styles(target_paragraph, cloned)
    target_paragraph._p.addnext(cloned)
    return Paragraph(cloned, target_paragraph._parent)


def _find_subtitle_target_paragraph(doc: Document):
    for paragraph in doc.paragraphs:
        if paragraph.text.strip() not in SUBTITLE_LABELS:
            continue
        next_elm = paragraph._p.getnext()
        while next_elm is not None and next_elm.tag != qn("w:p"):
            next_elm = next_elm.getnext()
        if next_elm is None:
            return insert_paragraph_after(paragraph, "")
        return Paragraph(next_elm, paragraph._parent)
    return None


def generate_subs(
    template_path: Path,
    source_docx_path: Path,
    input_path: Path,
    output_path: Path,
) -> None:
    data = parse_input(input_path)
    input_base = input_path.parent
    _, source_header, source_body = _extract_source_paragraphs(source_docx_path)
    doc = Document(str(template_path))
    apply_default_margins(doc)
    annotation_style = ensure_annotation_style(doc)
    ensure_hyperlink_style(doc)
    source_indent_inches = get_default_tab_stop_inches(doc)
    metrics = _get_section_metrics(doc)

    if doc.paragraphs and source_header:
        anchor = doc.paragraphs[0]
        for paragraph in source_header:
            _clone_paragraph_before(paragraph, anchor)

    for paragraph in list(doc.paragraphs):
        if "{{INTRO}}" in paragraph.text:
            intro = data.get("INTRO", "")
            if not intro and paragraph.text.strip() == "{{INTRO}}":
                remove_paragraph(paragraph)
                continue
            replace_body_paragraph(paragraph, intro, source_indent_inches)
            continue
        if "{{BODY}}" in paragraph.text:
            body = data.get("BODY", "")
            if not body and paragraph.text.strip() == "{{BODY}}":
                remove_paragraph(paragraph)
                continue
            replace_body_paragraph(paragraph, body, source_indent_inches)
            continue
        stripped = paragraph.text.strip()
        match = GENERIC_PLACEHOLDER_RE.fullmatch(stripped)
        if match and match.group(1) not in (INPUT_KEY_SET | {"INTRO", "BODY"}):
            remove_paragraph(paragraph)
            continue

        for key in PLACEHOLDER_KEYS:
            placeholder = f"{{{{{key}}}}}"
            if placeholder in paragraph.text:
                value = data.get(key, "")
                if not value and paragraph.text.strip() == placeholder:
                    remove_paragraph(paragraph)
                    break
                if key == "THUMBNAIL" and value:
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
                        thumbnail_credit = data.get("THUMBNAIL_CREDIT", "").strip()
                        if thumbnail_credit:
                            credit_paragraph = insert_paragraph_after(paragraph, "")
                            credit_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                            credit_paragraph.paragraph_format.left_indent = 0
                            credit_paragraph.paragraph_format.right_indent = 0
                            credit_paragraph.paragraph_format.first_line_indent = 0
                            _add_text_runs(
                                credit_paragraph,
                                thumbnail_credit,
                                run_style=annotation_style,
                            )
                    else:
                        replace_placeholder(paragraph, placeholder, value)
                else:
                    replace_placeholder(paragraph, placeholder, value)
                break
    ensure_blank_after_labels(doc, SECTION_LABELS)
    subtitle_target = _find_subtitle_target_paragraph(doc)
    body_text = data.get("BODY", "").strip()
    if body_text and subtitle_target is not None:
        # Keep one blank line between "字幕：" and the first timestamp line.
        clear_paragraph(subtitle_target)
        body_paragraph = insert_paragraph_after(subtitle_target, "")
        replace_body_paragraph(body_paragraph, body_text, source_indent_inches)
    elif subtitle_target is not None:
        current = subtitle_target
        for paragraph in source_body:
            current = _clone_paragraph_after(paragraph, current)

    doc.save(str(output_path))
    fix_docx_namespaces(output_path)


def main() -> None:
    parser = argparse.ArgumentParser(description="Fill a DOCX template from a text file.")
    parser.add_argument(
        "--template",
        default="templates/subs_template.docx",
        help="Path to the DOCX template.",
    )
    parser.add_argument(
        "--source-docx",
        required=True,
        help="Original source DOCX whose header and subtitles are preserved.",
    )
    parser.add_argument(
        "--input",
        default="input.txt",
        help="Path to the input text file.",
    )
    parser.add_argument(
        "--output",
        default="output/output.docx",
        help="Path to write the filled DOCX.",
    )
    args = parser.parse_args()

    generate_subs(
        Path(args.template),
        Path(args.source_docx),
        Path(args.input),
        with_subs_output_suffix(Path(args.output)),
    )


if __name__ == "__main__":
    main()
