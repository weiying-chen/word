#!/usr/bin/env python3

from __future__ import annotations

import argparse
import re
import unicodedata
import warnings
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
from docx.shared import RGBColor, Inches, Pt

from docx_utils import (
    add_hyperlink,
    clear_paragraph,
    ensure_blank_after_labels,
    get_default_tab_stop_inches,
    set_source_indent,
)


PLACEHOLDER_KEYS = [
    "TITLE",
    "URL",
    "SUMMARY",
    "YT_TITLE_SUGGESTED",
    "TITLE_SUGGESTED",
    "INTRO",
    "THUMBNAIL",
    "TIMING",
    "BODY",
]
PLACEHOLDER_KEY_SET = set(PLACEHOLDER_KEYS)
SOURCE_URL_RE = re.compile(r"^https?://\S+")
TIMING_LINE_RE = re.compile(
    r"^\d{2}:\d{2}:\d{2}:\d{2}\t\d{2}:\d{2}:\d{2}:\d{2}\t"
)
SYMBOL_FONT_NAME = "Segoe UI Symbol"
CJK_FONT_NAME = "新細明體"
CJK_MIDDLE_DOT = "\u2027"
HIGHLIGHT_MARKER_RE = re.compile(r"\*([^*]+)\*")
SOURCE_HIGHLIGHT_DEFAULT = WD_COLOR_INDEX.TURQUOISE
SOURCE_HIGHLIGHT_MARKED = WD_COLOR_INDEX.BRIGHT_GREEN
TIMING_HIGHLIGHT_MARKED = WD_COLOR_INDEX.YELLOW
SOURCE_HYPERLINK_HIGHLIGHT_MARKED = "brightGreen"
BOX_DRAWING_HORIZONTAL = "\u2500"
SPACED_HYPHEN_MINUS = " - "
SUBS_OUTPUT_SUFFIX = "_al"


def normalize_input_text(text: str) -> str:
    if not text:
        return text
    return text.replace(BOX_DRAWING_HORIZONTAL, SPACED_HYPHEN_MINUS)


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

        if key not in PLACEHOLDER_KEY_SET:
            idx += 1
            continue

        if key in {"BODY", "INTRO", "SUMMARY", "TIMING"}:
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
            data[key] = normalize_input_text("\n".join(collected).rstrip())
            continue

        data[key] = normalize_input_text(value)
        idx += 1

    data.setdefault("BODY", "")
    data.setdefault("INTRO", "")
    data.setdefault("SUMMARY", "")

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
            if timing_style:
                _add_marked_runs(
                    target,
                    text,
                    marked_highlight=TIMING_HIGHLIGHT_MARKED,
                    run_style=timing_style,
                    apply_default_size=False,
                )
            else:
                _add_text_runs(target, text)

    for idx, line in enumerate(lines):
        if idx > 0:
            current = insert_paragraph_after(current, "")

        if TIMING_LINE_RE.match(line):
            in_source_block = False

        cleaned_line = HIGHLIGHT_MARKER_RE.sub(r"\1", line)
        is_url = SOURCE_URL_RE.match(cleaned_line)
        if is_url:
            in_source_block = True

        write_line(
            current,
            cleaned_line if is_url else line,
            in_source_block and not TIMING_LINE_RE.match(line),
            bool(is_url),
        )


def remove_paragraph(paragraph) -> None:
    element = paragraph._element
    element.getparent().remove(element)


def ensure_timing_style(doc: Document):
    style_name = "Timing"
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


def generate_subs(template_path: Path, input_path: Path, output_path: Path) -> None:
    data = parse_input(input_path)
    input_base = input_path.parent
    doc = Document(str(template_path))
    apply_default_margins(doc)
    timing_style = ensure_timing_style(doc)
    ensure_hyperlink_style(doc)
    source_indent_inches = get_default_tab_stop_inches(doc)
    metrics = _get_section_metrics(doc)

    for paragraph in list(doc.paragraphs):
        if "{{INTRO}}" in paragraph.text:
            intro = data.get("INTRO", "")
            if not intro and paragraph.text.strip() == "{{INTRO}}":
                remove_paragraph(paragraph)
                continue
            replace_body_paragraph(paragraph, intro, source_indent_inches)
            continue
        if "{{SUMMARY}}" in paragraph.text:
            summary = data.get("SUMMARY", "")
            if not summary and paragraph.text.strip() == "{{SUMMARY}}":
                remove_paragraph(paragraph)
                continue
            replace_body_paragraph(paragraph, summary, source_indent_inches)
            continue
        if "{{BODY}}" in paragraph.text:
            body = data.get("BODY", "")
            if not body and paragraph.text.strip() == "{{BODY}}":
                remove_paragraph(paragraph)
                continue
            replace_body_paragraph(paragraph, body, source_indent_inches)
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
                elif key == "TIMING" and value:
                    replace_body_paragraph(
                        paragraph, value, source_indent_inches, timing_style=timing_style
                    )
                else:
                    replace_placeholder(paragraph, placeholder, value)
                break
    ensure_blank_after_labels(doc, {"簡介：", "簡介:", "字幕：", "字幕:"})

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
        Path(args.input),
        with_subs_output_suffix(Path(args.output)),
    )


if __name__ == "__main__":
    main()
