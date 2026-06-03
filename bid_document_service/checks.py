from __future__ import annotations

import re
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

from .schemas import MATRIX_HEADERS


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
POLLUTION_PHRASES = ["<think>", "</think>", "需人工确认", "TODO", "TBD", "待补充", "XXX", "内部备注"]


def docx_paragraphs(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    paras = []
    for paragraph in root.findall(".//w:p", NS):
        text = "".join(t.text or "" for t in paragraph.findall(".//w:t", NS)).strip()
        if text:
            paras.append(text)
    return paras


def heading_depth(path: Path) -> int:
    max_depth = 0
    with zipfile.ZipFile(path) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    for paragraph in root.findall(".//w:p", NS):
        text = "".join(t.text or "" for t in paragraph.findall(".//w:t", NS)).strip()
        if not text:
            continue
        style_el = paragraph.find("./w:pPr/w:pStyle", NS)
        style = style_el.attrib.get(f"{{{NS['w']}}}val", "") if style_el is not None else ""
        level = 0
        match = re.search(r"Heading(\d+)|标题(\d+)", style)
        if match:
            level = int(next(g for g in match.groups() if g))
        else:
            m = re.match(r"^\s*(\d+(?:\.\d+){0,4})[\s、.．]", text)
            if m:
                level = m.group(1).count(".") + 1
        max_depth = max(max_depth, level)
    return max_depth


def scan_docx_pollution(path: Path) -> list[str]:
    findings = []
    for idx, text in enumerate(docx_paragraphs(path), 1):
        for phrase in POLLUTION_PHRASES:
            if phrase in text:
                findings.append(f"段落{idx}包含 `{phrase}`：{text[:120]}")
    return findings


def matrix_missing_fields(rows: list[dict]) -> list[str]:
    if not rows:
        return MATRIX_HEADERS
    existing = set()
    for row in rows:
        existing.update(str(key) for key in row.keys())
    return [header for header in MATRIX_HEADERS if header not in existing]
