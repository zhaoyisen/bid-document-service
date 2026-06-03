from __future__ import annotations

import shutil
import secrets
import time
import uuid
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import Depends, FastAPI, File, Form, Header, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse

from .generator import OUTPUT_ROOT, generate_docx, generate_package, generate_xlsx, new_job_dir, safe_name, file_response, verify_download_token
from .maintenance import cleanup_old_outputs
from .models import BidPackageRequest, GenerateResponse, TenderFormatFillRequest
from .settings import get_settings
from .tender_format import extract_tender_format_file, generate_from_tender_format


@asynccontextmanager
async def lifespan(app_: FastAPI):
    settings = get_settings()
    removed = cleanup_old_outputs(settings.output_root, settings.output_retention_hours)
    if removed:
        print(f"cleanup_removed_output_dirs={removed}")
    yield


app = FastAPI(
    title="投标文档生成服务",
    description="接收 Dify 生成的投标结构化 JSON，输出 Word、Excel 和提交前检查报告。",
    version="1.0.0",
    lifespan=lifespan,
)


@app.middleware("http")
async def request_log_middleware(request: Request, call_next):
    request_id = request.headers.get("X-Request-ID") or uuid.uuid4().hex
    start = time.perf_counter()
    response = await call_next(request)
    elapsed_ms = int((time.perf_counter() - start) * 1000)
    response.headers["X-Request-ID"] = request_id
    print(f"{request_id} {request.method} {request.url.path} {response.status_code} {elapsed_ms}ms")
    return response


def verify_api_key(
    x_api_key: str | None = Header(default=None, alias="X-API-Key"),
    authorization: str | None = Header(default=None),
) -> None:
    settings = get_settings()
    expected = settings.api_key
    if not settings.require_api_key:
        return
    if not expected:
        raise HTTPException(status_code=500, detail="生产环境必须设置DOCUMENT_SERVICE_API_KEY")
    bearer = None
    if authorization and authorization.lower().startswith("bearer "):
        bearer = authorization[7:].strip()
    provided = [token for token in (x_api_key, bearer) if token]
    if not any(secrets.compare_digest(token, expected) for token in provided):
        raise HTTPException(status_code=401, detail="API Key无效")


@app.get("/health")
def health() -> dict[str, str]:
    settings = get_settings()
    return {"status": "ok", "version": app.version, "env": settings.app_env}


@app.post("/generate-bid-package", response_model=GenerateResponse, dependencies=[Depends(verify_api_key)])
def generate_bid_package(req: BidPackageRequest) -> GenerateResponse:
    return generate_package(req)


@app.post("/generate-docx", response_model=GenerateResponse, dependencies=[Depends(verify_api_key)])
def generate_docx_only(req: BidPackageRequest) -> GenerateResponse:
    job_id, output_dir = new_job_dir(req.project_name)
    docx = generate_docx(req, output_dir)
    return GenerateResponse(job_id=job_id, files=[file_response(job_id, docx, "docx")])


@app.post("/generate-xlsx", response_model=GenerateResponse, dependencies=[Depends(verify_api_key)])
def generate_xlsx_only(req: BidPackageRequest) -> GenerateResponse:
    job_id, output_dir = new_job_dir(req.project_name)
    xlsx = generate_xlsx(req, output_dir)
    return GenerateResponse(job_id=job_id, files=[file_response(job_id, xlsx, "xlsx")])


@app.post("/extract-tender-format", response_model=GenerateResponse, dependencies=[Depends(verify_api_key)])
async def extract_tender_format(
    tender_file: UploadFile = File(..., description="招标文件DOCX。V2优先支持DOCX原格式抽取。"),
    project_name: str | None = Form(default=None, description="项目名称，未传时使用文件名。"),
) -> GenerateResponse:
    if not tender_file.filename or not tender_file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="V2招标原格式抽取目前仅支持DOCX文件")
    name = project_name or Path(tender_file.filename).stem
    content = await tender_file.read()
    if len(content) > get_settings().max_upload_bytes:
        raise HTTPException(status_code=413, detail=f"文件超过大小限制：{get_settings().max_upload_mb}MB")
    job_id, output_dir = new_job_dir(name)
    source_path = output_dir / tender_file.filename
    source_path.write_bytes(content)
    try:
        result = extract_tender_format_file(source_path, name)
        shutil.rmtree(output_dir, ignore_errors=True)
        return result
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.post("/generate-from-tender-format", response_model=GenerateResponse, dependencies=[Depends(verify_api_key)])
def generate_from_tender_format_endpoint(req: TenderFormatFillRequest) -> GenerateResponse:
    try:
        return generate_from_tender_format(req)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.get("/files/{job_id}/{filename}")
def download_file(job_id: str, filename: str, _: None = Depends(verify_api_key)) -> FileResponse:
    safe_job = safe_name(job_id)
    if Path(filename).name != filename:
        raise HTTPException(status_code=400, detail="文件名非法")
    base = (OUTPUT_ROOT / safe_job).resolve()
    path = (base / filename).resolve()
    if not str(path).startswith(str(base)) or not path.exists() or not path.is_file():
        raise HTTPException(status_code=404, detail="文件不存在")
    return FileResponse(path=path, filename=filename)


@app.get("/public-files/{job_id}/{filename}")
def public_download_file(job_id: str, filename: str, expires: int, token: str) -> FileResponse:
    if not verify_download_token(job_id, filename, expires, token):
        raise HTTPException(status_code=401, detail="下载链接无效或已过期")
    safe_job = safe_name(job_id)
    if Path(filename).name != filename:
        raise HTTPException(status_code=400, detail="文件名非法")
    base = (OUTPUT_ROOT / safe_job).resolve()
    path = (base / filename).resolve()
    if not str(path).startswith(str(base)) or not path.exists() or not path.is_file():
        raise HTTPException(status_code=404, detail="文件不存在")
    return FileResponse(path=path, filename=filename)


@app.get("/openapi.json", include_in_schema=False)
def openapi_json() -> dict:
    return app.openapi()
