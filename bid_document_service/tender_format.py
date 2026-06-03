from __future__ import annotations

import json
import re
import shutil
from pathlib import Path
from typing import Any

from docx import Document
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

from .generator import (
    OUTPUT_ROOT,
    file_response,
    generate_report,
    generate_xlsx,
    new_job_dir,
    safe_name,
    sanitize_generated_text,
    value,
)
from .models import GeneratedFile, GenerateResponse, TenderFormatFillRequest, BidPackageRequest


START_KEYWORDS = [
    "响应文件格式",
    "投标文件格式",
    "应答文件格式",
    "报价文件格式",
    "响应文件组成",
    "投标文件组成",
    "附件格式",
]

END_KEYWORDS = [
    "采购需求",
    "采购项目内容及要求",
    "采购内容及要求",
    "项目内容及要求",
    "技术要求",
    "合同条款",
    "合同格式",
    "评审办法",
    "评分办法",
    "用户需求书",
]

FIELD_LABELS = [
    "项目名称",
    "项目编号",
    "供应商名称",
    "投标人名称",
    "投标人",
    "报价总价",
    "报价",
    "税率",
    "法定代表人",
    "授权代表",
    "委托代理人",
    "联系人",
    "联系电话",
    "地址",
    "日期",
    "服务期",
    "交付期",
    "质保期",
]


def local_name(element: Any) -> str:
    return element.tag.rsplit("}", 1)[-1]


def element_text(element: Any) -> str:
    texts = []
    for node in element.iter():
        if local_name(node) in {"t", "delText"} and node.text:
            texts.append(node.text)
    return "".join(texts).strip()


def is_start_text(text: str) -> bool:
    return any(keyword in text for keyword in START_KEYWORDS)


def is_toc_line(text: str) -> bool:
    stripped = text.strip()
    if "\t" in stripped and re.search(r"\t\s*\d+\s*$", stripped):
        return True
    if any(keyword in stripped for keyword in START_KEYWORDS + END_KEYWORDS) and re.search(r"\d+\s*$", stripped) and len(stripped) < 80:
        return True
    if re.search(r"\.{2,}\s*\d+\s*$", stripped):
        return True
    if re.search(r"\s{2,}\d+\s*$", stripped) and len(stripped) < 80:
        return True
    return False


def is_end_text(text: str) -> bool:
    if is_toc_line(text):
        return False
    if is_start_text(text):
        return False
    if not any(keyword in text for keyword in END_KEYWORDS):
        return False
    return bool(re.search(r"第[一二三四五六七八九十0-9]+[部分章节]", text) or re.match(r"^\d+[\.\s、]", text))


def find_format_range(doc: Document) -> tuple[int, int, list[str]]:
    warnings: list[str] = []
    body = doc._body._element
    children = list(body)
    start_idx = None
    skipped_toc = False
    for idx, child in enumerate(children):
        text = element_text(child)
        if is_start_text(text) and is_toc_line(text):
            skipped_toc = True
            continue
        if is_start_text(text):
            start_idx = idx
            break
    if start_idx is None:
        warnings.append("未识别到“响应文件格式/投标文件格式”等章节，已使用全文作为项目级模板。")
        return 0, len(children), warnings
    if skipped_toc:
        warnings.append("已跳过目录页中的响应文件格式条目，使用正文中的响应文件格式章节作为模板起点。")

    end_idx = len(children)
    for idx in range(start_idx + 1, len(children)):
        text = element_text(children[idx])
        if is_end_text(text):
            end_idx = idx
            break
    if end_idx == len(children):
        warnings.append("已识别响应文件格式起点，但未识别到下一章节，模板抽取到文档末尾。")
    return start_idx, end_idx, warnings


def keep_body_range(doc: Document, start_idx: int, end_idx: int) -> None:
    body = doc._body._element
    children = list(body)
    for idx, child in enumerate(children):
        if local_name(child) == "sectPr":
            continue
        if idx < start_idx or idx >= end_idx:
            body.remove(child)


def paragraph_text(paragraph: Paragraph) -> str:
    return "".join(run.text for run in paragraph.runs).strip()


def heading_level(paragraph: Paragraph, text: str) -> int | None:
    style_name = (paragraph.style.name or "") if paragraph.style else ""
    match = re.search(r"(\d+)$", style_name)
    if "Heading" in style_name and match:
        return int(match.group(1))
    m = re.match(r"^\s*(\d+(?:\.\d+){0,4})[\s、.．]", text)
    if m:
        return m.group(1).count(".") + 1
    if re.match(r"^[一二三四五六七八九十]+[、.．]", text):
        return 1
    return None


def extract_sections(doc: Document) -> list[dict[str, Any]]:
    sections = []
    for idx, paragraph in enumerate(doc.paragraphs, 1):
        text = paragraph_text(paragraph)
        if not text:
            continue
        level = heading_level(paragraph, text)
        if level:
            sections.append({
                "序号": len(sections) + 1,
                "层级": level,
                "标题": text,
                "段落序号": idx,
                "样式": paragraph.style.name if paragraph.style else "",
            })
    return sections


def table_text(cell) -> str:
    return "\n".join(p.text.strip() for p in cell.paragraphs if p.text.strip())


def extract_tables(doc: Document) -> list[dict[str, Any]]:
    tables = []
    for idx, table in enumerate(doc.tables, 1):
        first_row = []
        if table.rows:
            first_row = [table_text(cell) for cell in table.rows[0].cells]
        tables.append({
            "表格序号": idx,
            "行数": len(table.rows),
            "列数": len(table.columns),
            "首行": first_row,
        })
    return tables


def get_table_headers(table) -> list[str]:
    if not table.rows:
        return []
    return [table_text(cell).replace("\n", " ").strip() for cell in table.rows[0].cells]


def detect_fields_in_text(text: str, location: str) -> list[dict[str, Any]]:
    fields = []
    for label in FIELD_LABELS:
        if label not in text:
            continue
        if re.search(rf"{re.escape(label)}\s*[:：]\s*($|[_\s　]+|/|年\s*月\s*日)", text):
            fields.append({
                "字段名": label,
                "位置": location,
                "原文": text[:180],
                "填充策略": "按冒号后的空白位置填充；无法确认时保留空白。",
            })
    return fields


def extract_field_mapping(doc: Document) -> list[dict[str, Any]]:
    fields: list[dict[str, Any]] = []
    seen = set()
    for idx, paragraph in enumerate(doc.paragraphs, 1):
        for item in detect_fields_in_text(paragraph_text(paragraph), f"段落{idx}"):
            key = (item["字段名"], item["位置"], item["原文"])
            if key not in seen:
                fields.append(item)
                seen.add(key)
    for table_idx, table in enumerate(doc.tables, 1):
        for row_idx, row in enumerate(table.rows, 1):
            for col_idx, cell in enumerate(row.cells, 1):
                for item in detect_fields_in_text(table_text(cell), f"表格{table_idx} 行{row_idx} 列{col_idx}"):
                    key = (item["字段名"], item["位置"], item["原文"])
                    if key not in seen:
                        fields.append(item)
                        seen.add(key)
    return fields


def write_json(path: Path, data: Any) -> Path:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def extract_tender_format_file(source_docx: Path, project_name: str | None = None) -> GenerateResponse:
    source_docx = Path(source_docx)
    if source_docx.suffix.lower() != ".docx":
        raise ValueError("V2招标原格式抽取目前仅支持DOCX文件。PDF建议先转换为DOCX或作为解标来源。")

    name = project_name or source_docx.stem
    job_id, output_dir = new_job_dir(name)
    saved_source = output_dir / source_docx.name
    shutil.copy2(source_docx, saved_source)

    doc = Document(saved_source)
    start_idx, end_idx, warnings = find_format_range(doc)
    keep_body_range(doc, start_idx, end_idx)

    template_path = output_dir / f"{safe_name(name)}_项目级响应模板.docx"
    doc.save(template_path)

    template_doc = Document(template_path)
    sections = extract_sections(template_doc)
    tables = extract_tables(template_doc)
    fields = extract_field_mapping(template_doc)

    section_path = write_json(output_dir / f"{safe_name(name)}_章节清单.json", sections)
    table_path = write_json(output_dir / f"{safe_name(name)}_表格清单.json", tables)
    field_path = write_json(output_dir / f"{safe_name(name)}_字段映射.json", fields)
    report_path = output_dir / f"{safe_name(name)}_格式抽取报告.md"
    report_path.write_text(
        "\n".join([
            "# 招标原格式抽取报告",
            "",
            f"源文件：{source_docx.name}",
            f"模板范围：body元素 {start_idx} 到 {end_idx}",
            f"识别章节数：{len(sections)}",
            f"识别表格数：{len(tables)}",
            f"识别字段数：{len(fields)}",
            "",
            "## 提醒",
            *(f"- {warning}" for warning in warnings),
            "- 正式提交前仍需人工复核签字盖章、报价、授权、附件完整性和目录页码。",
        ]) + "\n",
        encoding="utf-8",
    )

    return GenerateResponse(
        job_id=job_id,
        files=[
            file_response(job_id, template_path, "docx-template"),
            file_response(job_id, field_path, "json"),
            file_response(job_id, section_path, "json"),
            file_response(job_id, table_path, "json"),
            file_response(job_id, report_path, "markdown"),
        ],
        warnings=warnings,
    )


def iter_all_paragraphs(doc: Document):
    for paragraph in doc.paragraphs:
        yield paragraph
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    yield paragraph


def fill_label_text(text: str, label: str, fill_value: str) -> str:
    if not fill_value:
        return text
    patterns = [
        rf"({re.escape(label)}\s*[:：])\s*[_\s　/年月日.-]*$",
        rf"({re.escape(label)}\s*[:：])\s*[_\s　/.-]+",
    ]
    for pattern in patterns:
        if re.search(pattern, text):
            return re.sub(pattern, rf"\1{fill_value}", text)
    return text


def fill_fields(doc: Document, fields: dict[str, Any]) -> list[str]:
    warnings = []
    normalized = {str(k): "" if v is None else str(v) for k, v in fields.items()}
    for paragraph in iter_all_paragraphs(doc):
        original = paragraph.text
        updated = original
        for key, val in normalized.items():
            updated = updated.replace(f"{{{{{key}}}}}", val)
            updated = updated.replace(f"{{{{ {key} }}}}", val)
            updated = fill_label_text(updated, key, val)
        if updated != original:
            paragraph.text = updated
    for key, val in normalized.items():
        if val and not any(key in p.text for p in iter_all_paragraphs(doc)):
            warnings.append(f"未在模板中找到可直接填充的字段：{key}")
    return warnings


def insert_paragraph_after(paragraph: Paragraph, text: str, style: str | None = None) -> Paragraph:
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    new_para.text = sanitize_generated_text(text)
    if style:
        try:
            new_para.style = style
        except Exception:
            pass
    return new_para


def add_paragraph_safe(doc: Document, text: str, style: str | None = None) -> Paragraph:
    text = sanitize_generated_text(text)
    if not style:
        return doc.add_paragraph(text)
    try:
        return doc.add_paragraph(text, style=style)
    except KeyError:
        return doc.add_paragraph(text)


def add_heading_safe(doc: Document, text: str, level: int = 1) -> Paragraph:
    text = sanitize_generated_text(text)
    try:
        return doc.add_heading(text, level=level)
    except KeyError:
        return doc.add_paragraph(text)


def section_matches(text: str, section: dict[str, Any]) -> bool:
    title = value(section, "标题") or value(section, "title")
    number = value(section, "章节编号") or value(section, "number")
    if title and title in text:
        return True
    if number and text.strip().startswith(number):
        return True
    return False


def section_level(section: dict[str, Any], default: int = 2) -> int:
    raw = section.get("层级", section.get("level", default))
    try:
        return max(1, min(4, int(raw)))
    except Exception:
        return default


def section_text_lines(content: str) -> list[str]:
    lines = []
    for raw in sanitize_generated_text(content).splitlines():
        line = raw.strip()
        if not line or line in {"```", "```markdown"}:
            continue
        lines.append(line)
    return lines


def insert_sections(doc: Document, sections: list[dict[str, Any]]) -> list[str]:
    warnings = []
    paragraphs = list(doc.paragraphs)
    unmatched = []
    for section in sections:
        content = value(section, "正文") or value(section, "content")
        points = section.get("编制要点", section.get("writing_points", []))
        if isinstance(points, str):
            points = [points]
        matched = None
        for paragraph in paragraphs:
            if section_matches(paragraph.text, section):
                matched = paragraph
                break
        lines = []
        if content:
            lines.extend(section_text_lines(content))
        for point in points:
            point_text = sanitize_generated_text(point)
            if point_text:
                lines.append(f"编制要点：{point_text}")
        if not lines:
            continue
        if matched:
            current = matched
            for line in reversed(lines):
                current = insert_paragraph_after(matched, line)
        else:
            unmatched.append(section)
    if unmatched:
        add_heading_safe(doc, "自动补充方案内容", level=1)
        for section in unmatched:
            heading = f"{value(section, '章节编号') or value(section, 'number')} {value(section, '标题') or value(section, 'title')}".strip()
            add_heading_safe(doc, heading or "未命名章节", level=section_level(section))
            content = value(section, "正文") or value(section, "content")
            if content:
                for line in section_text_lines(content):
                    add_paragraph_safe(doc, line)
            points = section.get("编制要点", section.get("writing_points", []))
            if isinstance(points, str):
                points = [points]
            for point in points:
                point_text = sanitize_generated_text(point)
                if point_text:
                    add_paragraph_safe(doc, point_text, style="List Bullet")
        warnings.append(f"{len(unmatched)}个章节未匹配到招标原格式标题，已追加到文末“自动补充方案内容”。")
    return warnings


def match_table(doc: Document, table_index: int | None, match_headers: list[str]):
    if table_index is not None:
        idx = table_index - 1
        if 0 <= idx < len(doc.tables):
            return doc.tables[idx], f"表格{table_index}"
        return None, f"未找到序号为{table_index}的表格"
    expected = [h.strip() for h in match_headers if h.strip()]
    if not expected:
        return None, "未提供表格序号或表头关键词"
    for idx, table in enumerate(doc.tables, 1):
        header_text = " ".join(get_table_headers(table))
        if all(keyword in header_text for keyword in expected):
            return table, f"表格{idx}"
    return None, "未找到匹配表头关键词的表格：" + "、".join(expected)


def clear_cell(cell) -> None:
    for paragraph in cell.paragraphs:
        paragraph.text = ""


def set_cell(cell, text: Any) -> None:
    clear_cell(cell)
    cell.text = sanitize_generated_text(text)


def delete_row(table, row_idx: int) -> None:
    row = table.rows[row_idx]
    table._tbl.remove(row._tr)


def fill_table_rows(table, rows: list[dict[str, Any]], mode: str = "replace_data_rows") -> None:
    if not rows:
        return
    headers = get_table_headers(table)
    if not headers:
        headers = list(rows[0].keys())
        new_row = table.add_row()
        for idx, header in enumerate(headers):
            if idx < len(new_row.cells):
                set_cell(new_row.cells[idx], header)
    if mode == "replace_data_rows" and len(table.rows) > 1:
        for idx in range(len(table.rows) - 1, 0, -1):
            delete_row(table, idx)
    for row_data in rows:
        row = table.add_row()
        for idx, header in enumerate(headers):
            if idx >= len(row.cells):
                break
            candidates = [header, header.replace(" ", ""), header.split("\n")[0]]
            cell_value = ""
            for key in candidates:
                if key in row_data:
                    cell_value = row_data[key]
                    break
            if cell_value == "":
                # 兼容常见表头别名
                for key, val in row_data.items():
                    if key in header or header in key:
                        cell_value = val
                        break
            set_cell(row.cells[idx], cell_value)


def fill_tables(doc: Document, table_fills) -> list[str]:
    warnings = []
    for spec in table_fills:
        table, note = match_table(doc, spec.table_index, spec.match_headers)
        if table is None:
            warnings.append(note)
            continue
        fill_table_rows(table, spec.rows, spec.mode)
        warnings.append(f"{note}已填充{len(spec.rows)}行数据。")
    return warnings


def generate_from_tender_format(req: TenderFormatFillRequest) -> GenerateResponse:
    template_dir = OUTPUT_ROOT / safe_name(req.template_job_id)
    template_name = req.template_file_name or "项目级响应模板.docx"
    candidates = [template_dir / template_name]
    if req.template_file_name is None:
        candidates.extend(template_dir.glob("*项目级响应模板.docx"))
    template_path = next((p for p in candidates if p.exists()), None)
    if template_path is None:
        raise FileNotFoundError(f"未找到项目级模板：{template_dir} / {template_name}")

    job_id, output_dir = new_job_dir(req.project_name)
    doc = Document(template_path)
    fields = dict(req.fields)
    if req.project_name:
        fields.setdefault("项目名称", req.project_name)
    if req.bidder_name:
        fields.setdefault("供应商名称", req.bidder_name)
        fields.setdefault("投标人名称", req.bidder_name)
        fields.setdefault("投标人", req.bidder_name)

    warnings = fill_fields(doc, fields)
    warnings.extend(fill_tables(doc, req.table_fills))
    warnings.extend(insert_sections(doc, req.sections))

    output_docx = output_dir / f"{safe_name(req.project_name)}_按招标原格式响应文件初稿.docx"
    doc.save(output_docx)

    package_req = BidPackageRequest(
        project_name=req.project_name,
        bidder_name=req.bidder_name,
        tender_name=req.tender_name,
        template_id=req.template_job_id,
        metadata={},
        response_matrix=req.response_matrix,
        material_gaps=req.material_gaps,
        sections=req.sections,
        checklist=req.checklist,
    )
    xlsx = generate_xlsx(package_req, output_dir)
    report = generate_report(package_req, output_dir, output_docx)

    return GenerateResponse(
        job_id=job_id,
        files=[
            file_response(job_id, output_docx, "docx"),
            file_response(job_id, xlsx, "xlsx"),
            file_response(job_id, report, "markdown"),
        ],
        warnings=warnings,
    )
