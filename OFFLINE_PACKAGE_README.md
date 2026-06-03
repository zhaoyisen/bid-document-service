# 投标文档服务离线部署包

## 包内文件

- `bid-document-service_1.0.0.tar`：Docker 镜像包。
- `docker-compose.offline.yml`：离线部署用 Compose 文件，不在服务器上重新构建镜像。
- `.env.example`：环境变量模板。
- `scripts/load_and_start.sh`：Linux 一键加载镜像并启动。
- `scripts/load_and_start.ps1`：Windows PowerShell 一键加载镜像并启动。
- `README.md`、`PRODUCTION_DEPLOY.md`：服务接口和部署说明。

## Linux 服务器部署

```bash
unzip bid-document-service-offline-1.0.0.zip -d bid-document-service
cd bid-document-service
cp .env.example .env
vi .env
sh scripts/load_and_start.sh
```

至少修改：

```text
DOCUMENT_SERVICE_API_KEY=替换为足够长的随机密钥
DOCUMENT_SERVICE_DOWNLOAD_SECRET=替换为另一个足够长的随机密钥
PUBLIC_BASE_URL=http://服务器IP:8010
```

## Windows 服务器部署

```powershell
Expand-Archive .\bid-document-service-offline-1.0.0.zip -DestinationPath .\bid-document-service -Force
cd .\bid-document-service
Copy-Item .env.example .env
notepad .env
.\scripts\load_and_start.ps1
```

## 验证

```bash
curl http://127.0.0.1:8010/health
```

Dify 调用时请求头：

```text
X-API-Key: .env 中的 DOCUMENT_SERVICE_API_KEY
```

生成结果中的 `public_url` 是带签名和有效期的下载链接，默认有效期由
`PUBLIC_DOWNLOAD_TTL_SECONDS` 控制。`DOCUMENT_SERVICE_DOWNLOAD_SECRET` 用于签名下载链接，
建议不要和 `DOCUMENT_SERVICE_API_KEY` 使用同一个值。

## 关闭 2375

本机为了构建曾临时打开 Docker `tcp://localhost:2375`。构建完成后建议关闭：

Docker Desktop -> Settings -> General -> 取消 `Expose daemon on tcp://localhost:2375 without TLS`。
