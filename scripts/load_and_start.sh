#!/usr/bin/env sh
set -eu

SCRIPT_DIR="$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)"
PACKAGE_ROOT="$(dirname "$SCRIPT_DIR")"
IMAGE_TAR="$PACKAGE_ROOT/bid-document-service_1.0.0.tar"
COMPOSE_FILE="$PACKAGE_ROOT/docker-compose.offline.yml"
ENV_FILE="$PACKAGE_ROOT/.env"

if [ ! -f "$IMAGE_TAR" ]; then
  echo "Image tar not found: $IMAGE_TAR" >&2
  exit 1
fi

if [ ! -f "$ENV_FILE" ]; then
  cp "$PACKAGE_ROOT/.env.example" "$ENV_FILE"
  echo "Created .env from .env.example. Edit DOCUMENT_SERVICE_API_KEY before production use."
fi

docker load -i "$IMAGE_TAR"
docker compose -f "$COMPOSE_FILE" up -d
docker compose -f "$COMPOSE_FILE" ps
