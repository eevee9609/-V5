"""可回復的 JSON 儲存工具。"""

from __future__ import annotations

import json
import os
import shutil
import tempfile
from datetime import datetime
from pathlib import Path


def write_json_with_backup(data: object, file_path: Path, *, backup_limit: int = 20) -> Path | None:
    """先備份既有 JSON，再以同磁碟原子取代方式寫入新內容。

    回傳本次建立的備份路徑；首次建立檔案時回傳 ``None``。
    """
    target = Path(file_path)
    target.parent.mkdir(parents=True, exist_ok=True)

    backup_path: Path | None = None
    if target.exists():
        backup_dir = target.parent / "backups"
        backup_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        backup_path = backup_dir / f"{target.stem}_{timestamp}{target.suffix}"
        shutil.copy2(target, backup_path)
        _trim_backups(backup_dir, target.stem, target.suffix, backup_limit)

    descriptor, temp_name = tempfile.mkstemp(
        prefix=f".{target.stem}_",
        suffix=".tmp",
        dir=target.parent,
        text=True,
    )
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8") as file:
            json.dump(data, file, ensure_ascii=False, indent=4)
            file.flush()
            os.fsync(file.fileno())
        os.replace(temp_name, target)
    except Exception:
        try:
            Path(temp_name).unlink(missing_ok=True)
        except OSError:
            pass
        raise

    return backup_path


def _trim_backups(backup_dir: Path, stem: str, suffix: str, backup_limit: int) -> None:
    if backup_limit < 1:
        return
    backups = sorted(
        backup_dir.glob(f"{stem}_*{suffix}"),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    for old_backup in backups[backup_limit:]:
        try:
            old_backup.unlink()
        except OSError:
            pass
