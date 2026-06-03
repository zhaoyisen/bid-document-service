from __future__ import annotations

import os
from functools import lru_cache
from pathlib import Path


class Settings:
    def __init__(self) -> None:
        base_dir = Path(__file__).resolve().parents[1]
        self.app_env = os.getenv("APP_ENV", "production").lower()
        self.host = os.getenv("HOST", "0.0.0.0")
        self.port = int(os.getenv("PORT", "8010"))
        self.api_key = os.getenv("DOCUMENT_SERVICE_API_KEY", "")
        self.output_root = Path(os.getenv("OUTPUT_ROOT", str(base_dir / "outputs"))).resolve()
        self.max_upload_mb = int(os.getenv("MAX_UPLOAD_MB", "100"))
        self.output_retention_hours = int(os.getenv("OUTPUT_RETENTION_HOURS", "168"))
        self.public_base_url = os.getenv("PUBLIC_BASE_URL", "").rstrip("/")
        self.download_secret = os.getenv("DOCUMENT_SERVICE_DOWNLOAD_SECRET", self.api_key)
        self.public_download_ttl_seconds = int(os.getenv("PUBLIC_DOWNLOAD_TTL_SECONDS", "86400"))

    @property
    def require_api_key(self) -> bool:
        return self.app_env not in {"dev", "development", "test"} or bool(self.api_key)

    @property
    def max_upload_bytes(self) -> int:
        return self.max_upload_mb * 1024 * 1024


@lru_cache
def get_settings() -> Settings:
    settings = Settings()
    settings.output_root.mkdir(parents=True, exist_ok=True)
    return settings
