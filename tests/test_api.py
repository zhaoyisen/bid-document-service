from __future__ import annotations

import json
from pathlib import Path

from docx import Document
from fastapi.testclient import TestClient

from bid_document_service.main import app
from bid_document_service.settings import get_settings


ROOT = Path(__file__).resolve().parents[1]
WORKSPACE = ROOT.parent


def sample_payload() -> dict:
    return json.loads((ROOT / "examples" / "sample_request.json").read_text(encoding="utf-8"))


def test_generate_bid_package_success(monkeypatch):
    monkeypatch.setenv("APP_ENV", "test")
    get_settings.cache_clear()
    client = TestClient(app)
    response = client.post("/generate-bid-package", json=sample_payload())
    assert response.status_code == 200
    data = response.json()
    assert data["job_id"]
    assert len(data["files"]) == 3
    assert {item["type"] for item in data["files"]} == {"docx", "xlsx", "markdown"}


def test_generated_text_artifacts_are_sanitized(monkeypatch):
    monkeypatch.setenv("APP_ENV", "test")
    get_settings.cache_clear()
    client = TestClient(app)
    payload = sample_payload()
    payload["sections"] = [
        {
            "章节编号": "10.1",
            "层级": 2,
            "标题": "服务方案",
            "正文": "<think>内部推理不应进入正文</think>本项目服务方案正文。",
            "编制要点": ["<think>内部要点</think>按招标要求细化实施路径。"],
        }
    ]
    response = client.post("/generate-bid-package", json=payload)
    assert response.status_code == 200
    data = response.json()
    docx_file = next(item for item in data["files"] if item["type"] == "docx")
    text = "\n".join(p.text for p in Document(docx_file["path"]).paragraphs)
    assert "<think>" not in text
    assert "内部推理不应进入正文" not in text
    assert "本项目服务方案正文。" in text


def test_api_key_required_in_production(monkeypatch):
    monkeypatch.setenv("APP_ENV", "production")
    monkeypatch.setenv("DOCUMENT_SERVICE_API_KEY", "secret")
    get_settings.cache_clear()
    client = TestClient(app)
    response = client.post("/generate-bid-package", json=sample_payload())
    assert response.status_code == 401
    response = client.post("/generate-bid-package", json=sample_payload(), headers={"X-API-Key": "secret"})
    assert response.status_code == 200


def test_public_signed_download_link(monkeypatch):
    monkeypatch.setenv("APP_ENV", "production")
    monkeypatch.setenv("DOCUMENT_SERVICE_API_KEY", "secret")
    monkeypatch.setenv("DOCUMENT_SERVICE_DOWNLOAD_SECRET", "download-secret")
    get_settings.cache_clear()
    client = TestClient(app)
    response = client.post("/generate-bid-package", json=sample_payload(), headers={"X-API-Key": "secret"})
    assert response.status_code == 200
    data = response.json()
    first_url = data["files"][0]["url"]
    assert first_url.startswith("/public-files/")
    download = client.get(first_url)
    assert download.status_code == 200
    tampered = client.get(first_url.replace("token=", "token=x"))
    assert tampered.status_code == 401


def test_extract_and_fill_tender_format(monkeypatch):
    monkeypatch.setenv("APP_ENV", "test")
    get_settings.cache_clear()
    client = TestClient(app)
    tender = WORKSPACE / "成都银行-招标文件.docx"
    if not tender.exists():
        tender = next(p for p in WORKSPACE.glob("*.docx") if p.stat().st_size == 237928)
    with tender.open("rb") as file:
        response = client.post(
            "/extract-tender-format",
            data={"project_name": "成都银行统一加密平台升级项目"},
            files={"tender_file": (tender.name, file, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
        )
    assert response.status_code == 200
    extracted = response.json()
    template = next(item for item in extracted["files"] if item["type"] == "docx-template")
    payload = sample_payload()
    fill_payload = {
        "project_name": payload["project_name"],
        "template_job_id": extracted["job_id"],
        "template_file_name": template["name"],
        "bidder_name": payload["bidder_name"],
        "tender_name": payload["tender_name"],
        "fields": {"项目名称": payload["project_name"], "项目编号": payload["metadata"]["项目编号"]},
        "table_fills": [
            {
                "table_index": 1,
                "rows": [{"序号": "1", "名称": "示例产品", "数量": "1"}],
                "mode": "append",
            }
        ],
        "response_matrix": payload["response_matrix"],
        "material_gaps": payload["material_gaps"],
        "sections": payload["sections"],
        "checklist": payload["checklist"],
    }
    response = client.post("/generate-from-tender-format", json=fill_payload)
    assert response.status_code == 200
    data = response.json()
    assert len(data["files"]) == 3
    assert any(item["type"] == "docx" for item in data["files"])
