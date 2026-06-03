$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$packageRoot = Split-Path -Parent $scriptRoot
$imageTar = Join-Path $packageRoot "bid-document-service_1.0.0.tar"
$composeFile = Join-Path $packageRoot "docker-compose.offline.yml"
$envFile = Join-Path $packageRoot ".env"

if (-not (Test-Path -LiteralPath $imageTar)) {
    throw "Image tar not found: $imageTar"
}

if (-not (Test-Path -LiteralPath $envFile)) {
    Copy-Item -LiteralPath (Join-Path $packageRoot ".env.example") -Destination $envFile
    Write-Host "Created .env from .env.example. Edit DOCUMENT_SERVICE_API_KEY before production use."
}

docker load -i $imageTar
docker compose -f $composeFile up -d
docker compose -f $composeFile ps
