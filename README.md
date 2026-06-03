# 投标文档生成服务

该服务用于接收 Dify Workflow 生成的结构化投标内容，并生成可下载的 Word、Excel 和 Markdown 检查报告。

当前版本：`1.0.1`

## 核心定位

Dify 负责理解和生成：

- 解标
- 响应矩阵
- 资料缺口
- 3-4级方案目录
- 正文要点
- 提交前检查

本服务负责落成文件：

- `响应文件初稿.docx`
- `投标工作底稿.xlsx`
- `提交前检查报告.md`

## 启动

```powershell
cd C:\Users\Administrator\Desktop\标书\bid-document-service
python -m uvicorn bid_document_service.main:app --host 0.0.0.0 --port 8010
```

启动后访问：

```text
http://127.0.0.1:8010/docs
```

Dify 中可以通过 Tool / Custom Tool 调用：

```text
POST http://服务地址:8010/generate-bid-package
```

Docker Compose 部署见 [DEPLOY.md](DEPLOY.md)。

## 可选鉴权

如果设置环境变量：

```powershell
$env:DOCUMENT_SERVICE_API_KEY="your-secret-key"
```

请求时需要携带：

```text
X-API-Key: your-secret-key
```

或：

```text
Authorization: Bearer your-secret-key
```

生产环境必须设置 `DOCUMENT_SERVICE_API_KEY`。仅当 `APP_ENV=dev/test` 时允许无密钥调用。

## 下载链接

生成接口返回的文件链接默认使用签名下载地址：

```text
/public-files/{job_id}/{filename}?expires=...&token=...
```

这样 Dify 最终答案中的链接可以直接点击下载，不需要浏览器额外携带 `X-API-Key`。

生产环境建议配置：

```powershell
$env:PUBLIC_BASE_URL="https://你的文档服务域名"
$env:DOCUMENT_SERVICE_DOWNLOAD_SECRET="another-secret"
$env:PUBLIC_DOWNLOAD_TTL_SECONDS="86400"
```

说明：

- `PUBLIC_BASE_URL` 用于返回完整外部下载链接。
- `DOCUMENT_SERVICE_DOWNLOAD_SECRET` 用于签名下载链接；未设置时默认使用 `DOCUMENT_SERVICE_API_KEY`。
- `PUBLIC_DOWNLOAD_TTL_SECONDS` 控制下载链接有效期，默认 86400 秒。
- 原 `/files/{job_id}/{filename}` 仍保留，需要 API Key，适合内部系统调用。

## 接口

### `GET /health`

健康检查。

### `POST /extract-tender-format`

V2接口。上传招标 DOCX，自动识别“响应文件格式/投标文件格式/附件格式”等章节，并抽取为项目级模板。

输入：

- `tender_file`: 招标 DOCX 文件。
- `project_name`: 项目名称，可选。

输出：

- 项目级响应模板 DOCX
- 字段映射 JSON
- 章节清单 JSON
- 表格清单 JSON
- 格式抽取报告 Markdown

说明：

- 当前 V2 优先支持 DOCX。PDF 可用于解标，但不适合做原格式模板抽取。
- 如果未识别到“响应文件格式”章节，会退化为使用全文作为模板，并在报告中提示。

### `POST /generate-bid-package`

生成完整投标交付包。

输入：

- 项目名称
- 投标人名称
- 招标文件名称
- 响应矩阵
- 资料缺口
- 方案章节
- 提交前检查项

输出：

- 任务 ID
- Word 下载链接
- Excel 下载链接
- Markdown 检查报告下载链接

### `POST /generate-docx`

只生成 Word 初稿。

### `POST /generate-xlsx`

只生成 Excel 工作底稿。

### `POST /generate-from-tender-format`

V2接口。使用 `/extract-tender-format` 生成的项目级模板，按招标原格式填充字段和方案章节。

输入：

- `project_name`: 项目名称。
- `template_job_id`: `/extract-tender-format` 返回的任务 ID。
- `template_file_name`: 项目级模板文件名。
- `fields`: 需要填入招标原格式的字段，例如项目名称、供应商名称、项目编号、日期。
- `table_fills`: 需要填入招标表格的数据，例如报价表、人员表、案例表、偏离表。
- `sections`: Dify 生成的方案章节。
- `response_matrix`: 响应矩阵。
- `material_gaps`: 资料缺口。
- `checklist`: 提交前检查项。

输出：

- 按招标原格式生成的响应文件初稿 DOCX
- 投标工作底稿 XLSX
- 提交前检查报告 Markdown

### `GET /files/{job_id}/{filename}`

下载生成文件。

## 示例

```powershell
python scripts\smoke_test.py
```

该脚本会调用服务生成一套示例交付包到：

```text
outputs/
```

V2 招标原格式测试：

```powershell
python scripts\smoke_test_v2.py
```

该脚本会使用当前工作区的招标 DOCX，先抽取项目级响应模板，再按该模板生成“按招标原格式响应文件初稿”。

## 当前版本边界

V1 使用通用生成模板，重点是稳定生成 Word、Excel 和检查报告。

V2 已支持：

- 从招标 DOCX 中抽取“响应文件格式”章节。
- 生成项目级模板。
- 识别字段、章节和表格清单。
- 按招标原格式填充字段和章节。
- 按表格序号或表头关键词填充表格数据。
- 生成确定性检查结果，包括响应矩阵字段、标题层级和正文残留。

后续可继续扩展：

- 更多客户专用表格规则，例如复杂合并单元格报价表。
- 接入对象存储，返回外网可访问文件链接。
- 做成 Dify Tool Plugin。
