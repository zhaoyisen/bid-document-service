from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from bid_document_service.generator import generate_package
from bid_document_service.models import BidPackageRequest


def main() -> int:
    sample = ROOT / "examples" / "sample_request.json"
    data = json.loads(sample.read_text(encoding="utf-8"))
    result = generate_package(BidPackageRequest(**data))
    print(result.model_dump_json(indent=2) if hasattr(result, "model_dump_json") else result.json(indent=2))
    for item in result.files:
        path = Path(item.path)
        if not path.exists():
            raise RuntimeError(f"missing output: {path}")
        if path.stat().st_size <= 0:
            raise RuntimeError(f"empty output: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
