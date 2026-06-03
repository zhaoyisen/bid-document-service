from __future__ import annotations

from docx import Document

from bid_document_service.tender_format import element_text, find_format_range


def test_find_format_range_skips_table_of_contents_entry():
    doc = Document()
    doc.add_paragraph("目 录")
    doc.add_paragraph("第三部分 响应文件格式\t18")
    doc.add_paragraph("第四部分 采购项目内容及要求\t31")
    doc.add_paragraph("第一部分 磋商邀请")
    doc.add_paragraph("第二部分 参与磋商供应商须知")
    doc.add_paragraph("第三部分 响应文件格式")
    doc.add_paragraph("一、报价表")
    doc.add_paragraph("第四部分 采购项目内容及要求")

    start_idx, end_idx, warnings = find_format_range(doc)
    children = list(doc._body._element)

    assert start_idx > 2
    assert element_text(children[start_idx]) == "第三部分 响应文件格式"
    assert element_text(children[end_idx]) == "第四部分 采购项目内容及要求"
    assert any("已跳过目录页" in warning for warning in warnings)
