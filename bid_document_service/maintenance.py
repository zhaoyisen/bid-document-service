from __future__ import annotations

import shutil
import time
from pathlib import Path


def cleanup_old_outputs(output_root: Path, retention_hours: int) -> int:
    if retention_hours <= 0 or not output_root.exists():
        return 0
    cutoff = time.time() - retention_hours * 3600
    removed = 0
    for child in output_root.iterdir():
        if not child.is_dir():
            continue
        try:
            mtime = child.stat().st_mtime
            if mtime < cutoff:
                shutil.rmtree(child)
                removed += 1
        except FileNotFoundError:
            continue
    return removed
