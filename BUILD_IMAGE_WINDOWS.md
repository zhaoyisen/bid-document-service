# Windows 构建镜像

当前机器 Docker Desktop 已安装到：

```text
E:\Docker\Docker
```

Docker WSL 数据目录已配置到：

```text
E:\Docker\wsl
```

镜像 tar 默认输出到：

```text
E:\DockerImages\bid-document-service_1.0.0.tar
```

## 构建命令

用管理员 PowerShell 进入服务目录：

```powershell
cd C:\Users\Administrator\Desktop\标书\bid-document-service
.\scripts\build_image_tar.ps1
```

脚本会执行：

```powershell
docker build -t bid-document-service:1.0.0 .
docker save -o E:\DockerImages\bid-document-service_1.0.0.tar bid-document-service:1.0.0
```

## 当前会话说明

如果普通 PowerShell 报：

```text
permission denied while trying to connect to the docker API
```

说明当前登录令牌没有 Docker pipe 权限。安装 Docker 后通常需要注销/重登，或者使用管理员 PowerShell 执行构建。
