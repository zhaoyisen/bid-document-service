from __future__ import annotations

from typing import Any

from pydantic import BaseModel, Field


class BidPackageRequest(BaseModel):
    project_name: str = Field(..., description="项目名称")
    bidder_name: str | None = Field(None, description="投标人/供应商名称")
    tender_name: str | None = Field(None, description="招标文件或采购文件名称")
    template_id: str | None = Field(None, description="模板ID，V1仅记录不强制使用")
    metadata: dict[str, Any] = Field(default_factory=dict, description="项目元数据")
    response_matrix: list[dict[str, Any]] = Field(default_factory=list, description="响应矩阵")
    material_gaps: list[dict[str, Any]] = Field(default_factory=list, description="资料缺口")
    sections: list[dict[str, Any]] = Field(default_factory=list, description="方案章节")
    checklist: list[dict[str, Any]] = Field(default_factory=list, description="提交前检查项")


class GeneratedFile(BaseModel):
    name: str
    type: str
    path: str
    url: str


class GenerateResponse(BaseModel):
    job_id: str
    files: list[GeneratedFile]
    warnings: list[str] = Field(default_factory=list)


class TableFillSpec(BaseModel):
    table_index: int | None = Field(None, description="表格序号，从1开始")
    match_headers: list[str] = Field(default_factory=list, description="用于匹配表格的表头关键词")
    rows: list[dict[str, Any]] = Field(default_factory=list, description="需要填充的数据行")
    mode: str = Field("replace_data_rows", description="replace_data_rows或append")


class TenderFormatFillRequest(BaseModel):
    project_name: str = Field(..., description="项目名称")
    template_job_id: str = Field(..., description="项目级模板所在任务ID")
    template_file_name: str | None = Field(None, description="项目级模板文件名，默认使用任务目录中的项目级响应模板.docx")
    bidder_name: str | None = Field(None, description="投标人/供应商名称")
    tender_name: str | None = Field(None, description="招标文件或采购文件名称")
    fields: dict[str, Any] = Field(default_factory=dict, description="需要填入招标原格式的字段")
    table_fills: list[TableFillSpec] = Field(default_factory=list, description="需要填入招标表格的数据")
    response_matrix: list[dict[str, Any]] = Field(default_factory=list, description="响应矩阵")
    material_gaps: list[dict[str, Any]] = Field(default_factory=list, description="资料缺口")
    sections: list[dict[str, Any]] = Field(default_factory=list, description="需要插入到对应章节的方案内容")
    checklist: list[dict[str, Any]] = Field(default_factory=list, description="提交前检查项")
