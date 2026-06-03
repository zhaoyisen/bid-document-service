from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
WORKSPACE = ROOT.parent
sys.path.insert(0, str(ROOT))

from bid_document_service.models import TenderFormatFillRequest
from bid_document_service.tender_format import extract_tender_format_file, generate_from_tender_format


def find_tender_docx() -> Path:
    preferred = WORKSPACE / "成都银行-招标文件.docx"
    if preferred.exists():
        return preferred
    candidates = sorted(WORKSPACE.glob("*.docx"), key=lambda p: p.stat().st_size, reverse=True)
    for path in candidates:
        if "招标" in path.name:
            return path
    raise FileNotFoundError("未找到可用于V2测试的招标DOCX文件")


def main() -> int:
    tender = find_tender_docx()
    extract_result = extract_tender_format_file(tender, "成都银行统一加密平台升级项目")
    template_file = next(item for item in extract_result.files if item.type == "docx-template")

    sample = json.loads((ROOT / "examples" / "sample_request.json").read_text(encoding="utf-8"))
    fields = {
        "项目名称": sample["project_name"],
        "供应商名称": sample["bidder_name"],
        "投标人名称": sample["bidder_name"],
        "项目编号": sample["metadata"].get("项目编号", ""),
        "日期": "",
    }
    fill_req = TenderFormatFillRequest(
        project_name=sample["project_name"],
        template_job_id=extract_result.job_id,
        template_file_name=template_file.name,
        bidder_name=sample["bidder_name"],
        tender_name=sample["tender_name"],
        fields=fields,
        response_matrix=sample["response_matrix"],
        material_gaps=sample["material_gaps"],
        sections=sample["sections"],
        checklist=sample["checklist"],
    )
    fill_result = generate_from_tender_format(fill_req)
    print(json.dumps({
        "extract_job_id": extract_result.job_id,
        "extract_files": [item.model_dump() if hasattr(item, "model_dump") else item.dict() for item in extract_result.files],
        "fill_job_id": fill_result.job_id,
        "fill_files": [item.model_dump() if hasattr(item, "model_dump") else item.dict() for item in fill_result.files],
        "warnings": fill_result.warnings,
    }, ensure_ascii=False, indent=2))
    for item in extract_result.files + fill_result.files:
        path = Path(item.path)
        if not path.exists() or path.stat().st_size <= 0:
            raise RuntimeError(f"输出文件异常: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
