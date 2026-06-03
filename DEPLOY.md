# Docker Compose 部署说明

## 1. 准备环境变量

```bash
cp .env.example .env
```

修改 `.env`：

```text
DOCUMENT_SERVICE_API_KEY=替换为足够长的随机密钥
PORT=8010
MAX_UPLOAD_MB=100
OUTPUT_RETENTION_HOURS=168
PUBLIC_BASE_URL=http://你的服务器:8010
```

## 2. 构建并启动

```bash
docker compose up -d --build
```

## 3. 验证

```bash
curl http://127.0.0.1:8010/health
```

接口文档：

```text
http://服务器IP:8010/docs
```

## 4. Dify 调用方式

在 Dify 中创建 Custom Tool 或 HTTP Tool：

- URL: `http://服务器IP:8010/generate-bid-package`
- Method: `POST`
- Header:

```text
X-API-Key: 你的DOCUMENT_SERVICE_API_KEY
Content-Type: application/json
```

V2 原格式流程：

1. 调用 `/extract-tender-format` 上传招标 DOCX，拿到 `template_job_id` 和模板文件名。
2. 调用 `/generate-from-tender-format`，传入 `template_job_id`、字段、表格数据、章节内容、响应矩阵和检查项。
3. 使用返回的 `/files/...` 链接下载 Word/Excel/报告。

## 5. 文件持久化

生成文件保存在宿主机：

```text
./outputs
```

默认保留 168 小时，启动时自动清理过期任务目录。可通过 `OUTPUT_RETENTION_HOURS` 调整。

## 6. 生产注意事项

- 必须设置 `DOCUMENT_SERVICE_API_KEY`。
- 建议在前面加 Nginx，开启 HTTPS。
- `PUBLIC_BASE_URL` 应设置成 Dify 能访问的服务地址。
- 如果部署在内网，Dify 服务器必须能访问该地址。
