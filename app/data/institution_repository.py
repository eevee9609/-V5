"""代檢院所資料庫的讀寫。"""

from __future__ import annotations

import json
from pathlib import Path

from app.config import INSTITUTION_DB_FILE
from app.data.json_storage import write_json_with_backup


def load_institution_database(file_path: Path = INSTITUTION_DB_FILE) -> dict:
    if not file_path.exists():
        return {}
    try:
        with file_path.open("r", encoding="utf-8") as file:
            return json.load(file)
    except Exception:
        return {}


def save_institution_database(data: dict, file_path: Path = INSTITUTION_DB_FILE) -> None:
    write_json_with_backup(data, file_path)
