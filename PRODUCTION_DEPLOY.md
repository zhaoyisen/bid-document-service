# 生产部署说明

## 推荐方式：Git 拉取更新

第一次部署：

```bash
git clone https://github.com/zhaoyisen/bid-document-service.git
cd bid-document-service
cp .env.example .env
vi .env
docker compose build --no-cache bid-document-service
docker compose up -d
```

后续更新：

```bash
cd bid-document-service
git pull
docker compose build --no-cache bid-document-service
docker compose up -d
```

说明：

- 仓库只提交 `.env.example`，不提交真实 `.env`。
- 首次部署时复制 `.env.example` 为 `.env`，并把 `DOCUMENT_SERVICE_API_KEY` 和 `DOCUMENT_SERVICE_DOWNLOAD_SECRET` 改成真实长随机密钥。
- `outputs/` 是生成文件目录，不进入 Git。

## 部署方式一：服务器直接构建

适合服务器可以访问 Python 包源的情况。

```bash
cp .env.example .env
```

修改 `.env`：

```text
DOCUMENT_SERVICE_API_KEY=替换为足够长的随机密钥
PORT=8010
MAX_UPLOAD_MB=100
OUTPUT_RETENTION_HOURS=168
PUBLIC_BASE_URL=http://服务器IP:8010
```

启动：

```bash
docker compose up -d --build
```

检查：

```bash
curl http://127.0.0.1:8010/health
```

## 部署方式二：先打镜像包再上传

适合服务器不能联网安装依赖，或者你希望先在构建机上完成镜像。

Windows PowerShell：

```powershell
.\scripts\build_image_tar.ps1
```

Linux：

```bash
sh scripts/build_image_tar.sh
```

得到：

```text
bid-document-service_1.0.0.tar
```

上传到服务器后导入：

```bash
docker load -i bid-document-service_1.0.0.tar
docker compose up -d
```

## Dify 调用

HTTP Tool 或 Custom Tool 配置：

```text
POST http://服务器IP:8010/generate-bid-package
Header:
  X-API-Key: 你的 DOCUMENT_SERVICE_API_KEY
  Content-Type: application/json
```

按招标原格式生成时：

1. 调用 `/extract-tender-format` 上传招标 DOCX，取得 `template_job_id` 和模板文件名。
2. 调用 `/generate-from-tender-format`，传入字段、表格、章节、响应矩阵、资料缺口和检查项。
3. 使用返回的 `/files/{job_id}/{filename}` 下载 Word、Excel 和检查报告。

## 已具备的生产能力

- Docker Compose 部署。
- 生产环境 API Key 鉴权。
- 上传文件大小限制。
- 输出目录持久化。
- 启动时清理过期输出目录。
- Word、Excel、Markdown 交付物生成。
- 招标 DOCX 原格式抽取和项目级模板填充。
- 表格按序号或表头关键词填充。
- 响应矩阵字段、标题层级、正文污染词确定性检查。
- `/health` 健康检查。

## 当前边界

- 当前优先支持 DOCX 招标文件的原格式抽取；PDF 可进入解标流程，但不建议直接作为原格式模板。
- 复杂合并单元格、嵌套表格、报价表公式等，需要结合具体招标格式继续补充规则。
- 服务负责“文件落地和确定性检查”，不负责替代 Dify/LLM 的解标、内容生成和人工复核。
- 正式上线建议放在 Nginx/HTTPS 后面，并限制只允许 Dify 服务器或内网调用。
