#!/usr/bin/env python3

from __future__ import annotations

import argparse
import re
import shutil
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from docx import Document


def _non_empty_paragraphs(path: Path) -> list[str]:
    doc = Document(str(path))
    return [p.text.strip() for p in doc.paragraphs if p.text and p.text.strip()]


def _extract_embedded_url(path: Path) -> str:
    rel_ns = "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
    with zipfile.ZipFile(path) as zf:
        for name in zf.namelist():
            if not name.endswith('.rels'):
                continue
            try:
                root = ET.fromstring(zf.read(name))
            except Exception:
                continue
            for rel in root.findall(rel_ns):
                target = rel.attrib.get('Target', '')
                if target.startswith('http://') or target.startswith('https://'):
                    return target
    return ''


def _clean_title(text: str) -> str:
    t = text.strip()
    if not t:
        return ''
    t = re.split(r"\s*[｜|]\s*大愛新聞", t, maxsplit=1)[0].strip()
    t = re.split(r"\s+#", t, maxsplit=1)[0].strip()
    return t


def _extract_super_people(body_lines: list[str]) -> list[tuple[str, str]]:
    entries: list[tuple[str, str]] = []
    seen: set[str] = set()

    for i, line in enumerate(body_lines):
        if not line.startswith('/*SUPER'):
            continue

        role = ''
        zh_name = ''

        j = i + 1
        while j < len(body_lines):
            s = body_lines[j].strip()
            if s.startswith('*/'):
                break
            if '｜' in s and '//' in s:
                head = s.split('//', 1)[0].strip()
                parts = [x.strip() for x in head.split('｜', 1)]
                if len(parts) == 2:
                    role, zh_name = parts
                break
            j += 1

        if not role or not zh_name:
            continue

        en_name = ''
        # Look backward for a timing/name hint like: (6． Uyanda烏漾達)
        for k in range(max(0, i - 4), i):
            prev = body_lines[k].strip()
            m = re.search(r"\((?:[^)]*?)[\.．]\s*([A-Za-z][A-Za-z\s\-']*)", prev)
            if m:
                en_name = m.group(1).strip()
                break

        key = f"{role} | {zh_name}"
        if key in seen:
            continue
        seen.add(key)
        entries.append((key, en_name))

    return entries


def build_source_sections(source_docx: Path) -> dict[str, str]:
    paras = _non_empty_paragraphs(source_docx)
    if not paras:
        return {
            'TITLE_TEXT': '',
            'TITLE_URL': _extract_embedded_url(source_docx),
            'SUMMARY': '',
            'SUPER_PEOPLE': '',
            'BODY': '',
        }

    title_text = _clean_title(paras[0])
    title_url = _extract_embedded_url(source_docx)

    marker_idx = next((i for i, t in enumerate(paras) if t == '<'), None)
    summary = ''
    body_start = 1

    if marker_idx is not None:
        body_start = marker_idx
        if marker_idx > 1:
            summary = paras[marker_idx - 1]
    elif len(paras) > 1:
        summary = paras[1]
        body_start = 2

    body_lines = paras[body_start:] if body_start < len(paras) else []

    people_entries = _extract_super_people(body_lines)
    super_lines: list[str] = []
    for zh, en in people_entries:
        super_lines.append(zh)
        super_lines.append(en)
        super_lines.append('')
    if super_lines and super_lines[-1] == '':
        super_lines.pop()

    return {
        'TITLE_TEXT': title_text,
        'TITLE_URL': title_url,
        'SUMMARY': summary,
        'SUPER_PEOPLE': '\n'.join(super_lines),
        'BODY': '\n'.join(body_lines),
    }


def write_source_txt(path: Path, sections: dict[str, str]) -> None:
    out = [
        f"TITLE_TEXT: {sections.get('TITLE_TEXT', '')}",
        f"TITLE_URL: {sections.get('TITLE_URL', '')}",
        'SUMMARY:',
    ]
    summary = sections.get('SUMMARY', '')
    out.extend(summary.splitlines() if summary else [''])

    out.extend(['', 'SUPER_PEOPLE:'])
    super_people = sections.get('SUPER_PEOPLE', '')
    out.extend(super_people.splitlines() if super_people else [''])

    out.extend(['', 'BODY:'])
    body = sections.get('BODY', '')
    out.extend(body.splitlines() if body else [''])
    out.append('')

    path.write_text('\n'.join(out), encoding='utf-8')


def extract_news(input_docx: Path, source_docx: Path, source_txt: Path) -> None:
    shutil.copyfile(input_docx, source_docx)
    sections = build_source_sections(source_docx)
    write_source_txt(source_txt, sections)


def main() -> None:
    parser = argparse.ArgumentParser(
        description='Extract source.docx into a news-native source.txt scaffold.'
    )
    parser.add_argument('--input-docx', required=True)
    parser.add_argument('--source-docx', default='source.docx')
    parser.add_argument('--source-txt', default='source.txt')
    args = parser.parse_args()

    extract_news(Path(args.input_docx), Path(args.source_docx), Path(args.source_txt))


if __name__ == '__main__':
    main()
