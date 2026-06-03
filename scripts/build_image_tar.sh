#!/usr/bin/env sh
set -eu

IMAGE_NAME="${IMAGE_NAME:-bid-document-service:1.0.0}"
OUTPUT_FILE="${OUTPUT_FILE:-bid-document-service_1.0.0.tar}"

docker build -t "$IMAGE_NAME" .
docker save -o "$OUTPUT_FILE" "$IMAGE_NAME"
echo "Saved Docker image to $OUTPUT_FILE"
