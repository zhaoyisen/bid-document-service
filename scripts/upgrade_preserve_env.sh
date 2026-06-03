#!/usr/bin/env bash
set -euo pipefail

if [ $# -lt 1 ]; then
  echo "Usage: $0 /path/to/bid-document-service-deploy.zip [deploy_dir]" >&2
  exit 2
fi

ZIP_PATH="$1"
DEPLOY_DIR="${2:-$(pwd)}"

if [ ! -f "$ZIP_PATH" ]; then
  echo "Zip file not found: $ZIP_PATH" >&2
  exit 1
fi

mkdir -p "$DEPLOY_DIR"
ZIP_PATH="$(cd "$(dirname "$ZIP_PATH")" && pwd)/$(basename "$ZIP_PATH")"
DEPLOY_DIR="$(cd "$DEPLOY_DIR" && pwd)"

TMP_DIR="$(mktemp -d)"
cleanup() {
  rm -rf "$TMP_DIR"
}
trap cleanup EXIT

if [ -f "$DEPLOY_DIR/.env" ]; then
  cp "$DEPLOY_DIR/.env" "$TMP_DIR/.env.keep"
fi

unzip -q "$ZIP_PATH" -d "$TMP_DIR/package"

if command -v rsync >/dev/null 2>&1; then
  rsync -a \
    --exclude ".env" \
    --exclude "outputs/" \
    --exclude "__pycache__/" \
    --exclude ".pytest-tmp/" \
    "$TMP_DIR/package/" "$DEPLOY_DIR/"
else
  (cd "$TMP_DIR/package" && tar --exclude="./.env" --exclude="./outputs" -cf - .) | (cd "$DEPLOY_DIR" && tar -xf -)
fi

if [ -f "$TMP_DIR/.env.keep" ]; then
  cp "$TMP_DIR/.env.keep" "$DEPLOY_DIR/.env"
fi

echo "Updated $DEPLOY_DIR and preserved .env."
echo "Next:"
echo "  cd $DEPLOY_DIR"
echo "  docker compose build --no-cache bid-document-service"
echo "  docker compose up -d"
