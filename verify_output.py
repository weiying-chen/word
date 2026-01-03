#!/usr/bin/env python3

from __future__ import annotations

import argparse
import zipfile
from pathlib import Path
from typing import List

import xml.etree.ElementTree as ET



NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
}

SECTION_LABELS = [
    "建議YT標題：",
    "建議標題：",
    "簡介：",
    "選圖：",
    "字幕：",
]

SECTION_LABEL_SET = set(SECTION_LABELS)


def load_docx_xml(path: Path) -> ET.Element:
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    return ET.fromstring(xml)


def iter_paragraphs(root: ET.Element):
    for p in root.findall(".//w:p", NS):
        parts: List[str] = []
        for node in p.iter():
            if node.tag == f"{{{NS['w']}}}t":
                parts.append(node.text or "")
            elif node.tag == f"{{{NS['w']}}}tab":
                parts.append("\t")
        text = "".join(parts)
        has_drawing = p.find(".//w:drawing", NS) is not None
        yield p, text, has_drawing


def get_section_metrics(root: ET.Element) -> tuple[int, int, int]:
    sect = root.find(".//w:sectPr", NS)
    if sect is None:
        raise ValueError("No section properties found.")
    pg_sz = sect.find("w:pgSz", NS)
    pg_mar = sect.find("w:pgMar", NS)
    if pg_sz is None or pg_mar is None:
        raise ValueError("Missing page size or margin definitions.")
    page_width = int(pg_sz.attrib[f"{{{NS['w']}}}w"])
    left = int(pg_mar.attrib[f"{{{NS['w']}}}left"])
    right = int(pg_mar.attrib[f"{{{NS['w']}}}right"])
    return page_width, left, right


def find_thumbnail_anchor(root: ET.Element, label_idx: int, next_label_idx: int, paragraphs):
    for idx in range(label_idx + 1, next_label_idx):
        p, _, has_drawing = paragraphs[idx]
        if not has_drawing:
            continue
        anchor = p.find(".//wp:anchor", NS)
        if anchor is not None:
            return anchor
    return None


def main() -> int:
    parser = argparse.ArgumentParser(description="Verify generated output.docx.")
    parser.add_argument("--output", default="output.docx", help="Generated docx.")
    args = parser.parse_args()

    output_path = Path(args.output)
    errors: List[str] = []

    root = load_docx_xml(output_path)
    paragraphs = list(iter_paragraphs(root))
    paragraph_texts = [text.strip() for _, text, _ in paragraphs if text.strip()]
    paragraph_data = [(p, text.strip(), has_drawing) for p, text, has_drawing in paragraphs]

    # Ensure no raw placeholders remain.
    for _, text, _ in paragraphs:
        if "{{" in text or "}}" in text:
            errors.append(f"Placeholder text remains in docx: {text.strip()}")
            break

    # Section presence checks.
    section_labels_present = {label: label in paragraph_texts for label in SECTION_LABELS}
    for label, present in section_labels_present.items():
        if not present:
            errors.append(f"Missing section label: {label}")

    # Title/URL/Summary (non-empty paragraphs before first section label).
    first_label_idx = next(
        (i for i, (_, text, _) in enumerate(paragraph_data) if text in SECTION_LABEL_SET),
        None,
    )
    pre_section = []
    end_idx = first_label_idx if first_label_idx is not None else len(paragraph_data)
    for i in range(end_idx):
        _, text, _ = paragraph_data[i]
        if text:
            pre_section.append(text)

    title_ok = url_ok = summary_ok = False
    url_idx = next((i for i, text in enumerate(pre_section) if text.startswith("http")), None)
    if url_idx is not None:
        url_ok = True
        if url_idx > 0:
            title_ok = True
        if url_idx + 1 < len(pre_section):
            summary_ok = True
    else:
        if pre_section:
            # No URL line means URL missing; treat first line as title if present.
            title_ok = True
        else:
            errors.append("Missing title/url/summary block before section labels.")

    def find_label_idx(label: str):
        return next(
            (i for i, (_, text, _) in enumerate(paragraph_data) if text == label),
            None,
        )

    def find_next_label_idx(start_idx: int):
        if start_idx is None:
            return None
        for i in range(start_idx + 1, len(paragraph_data)):
            if paragraph_data[i][1] in SECTION_LABEL_SET:
                return i
        return None

    def has_text_between(start_idx: int, end_idx: int) -> bool:
        if start_idx is None or end_idx is None:
            return False
        for i in range(start_idx + 1, end_idx):
            _, text, _ = paragraph_data[i]
            if text:
                return True
        return False

    def has_drawing_between(start_idx: int, end_idx: int) -> bool:
        if start_idx is None or end_idx is None:
            return False
        for i in range(start_idx + 1, end_idx):
            _, _, has_drawing = paragraph_data[i]
            if has_drawing:
                return True
        return False

    yt_idx = find_label_idx("建議YT標題：")
    title_suggested_idx = find_label_idx("建議標題：")
    intro_idx = find_label_idx("簡介：")
    thumb_idx = find_label_idx("選圖：")
    body_label_idx = find_label_idx("字幕：")

    yt_ok = has_text_between(yt_idx, find_next_label_idx(yt_idx))
    title_suggested_ok = has_text_between(title_suggested_idx, find_next_label_idx(title_suggested_idx))
    intro_ok = has_text_between(intro_idx, find_next_label_idx(intro_idx))
    thumbnail_ok = has_drawing_between(thumb_idx, find_next_label_idx(thumb_idx))

    time_range_ok = False
    body_ok = False
    if body_label_idx is not None:
        next_label = find_next_label_idx(body_label_idx)
        # Find first non-empty paragraph after "字幕：".
        time_idx = None
        for i in range(body_label_idx + 1, len(paragraph_data) if next_label is None else next_label):
            _, text, _ = paragraph_data[i]
            if text:
                time_idx = i
                time_range_ok = True
                break
        if time_idx is not None:
            # Any non-empty paragraph after time range counts as body content.
            for i in range(time_idx + 1, len(paragraph_data) if next_label is None else next_label):
                _, text, _ = paragraph_data[i]
                if text:
                    body_ok = True
                    break

    if errors:
        for error in errors:
            print(f"ERROR: {error}")
        return 1

    def mark(ok: bool) -> str:
        return "[OK]" if ok else "[X]"

    print("Output verification:")
    print(f"- TITLE: {mark(title_ok)}")
    print(f"- URL: {mark(url_ok)}")
    print(f"- SUMMARY: {mark(summary_ok)}")
    print(f"- YT_TITLE_SUGGESTED: {mark(yt_ok)}")
    print(f"- TITLE_SUGGESTED: {mark(title_suggested_ok)}")
    print(f"- INTRO: {mark(intro_ok)}")
    print(f"- THUMBNAIL: {mark(thumbnail_ok)}")
    print(f"- TIME_RANGE: {mark(time_range_ok)}")
    print(f"- BODY: {mark(body_ok)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
