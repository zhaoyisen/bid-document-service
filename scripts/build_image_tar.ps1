param(
    [string]$ImageName = "bid-document-service:1.0.0",
    [string]$OutputFile = "E:\DockerImages\bid-document-service_1.0.0.tar"
)

$ErrorActionPreference = "Stop"

$docker = "docker"
$bundledDocker = "E:\Docker\Docker\resources\bin\docker.exe"
if (Test-Path -LiteralPath $bundledDocker) {
    $docker = $bundledDocker
    $env:DOCKER_CLI_PLUGIN_EXTRA_DIRS = "E:\Docker\Docker\resources\cli-plugins"
}

New-Item -ItemType Directory -Path (Split-Path -Parent $OutputFile) -Force | Out-Null

& $docker build -t $ImageName .
& $docker save -o $OutputFile $ImageName
Write-Host "Saved Docker image to $OutputFile"
