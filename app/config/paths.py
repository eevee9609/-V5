"""集中管理原始資料檔、資源與輸出檔的路徑。"""

from __future__ import annotations

import sys
from pathlib import Path


def _application_root() -> Path:
    """取得原始碼或 PyInstaller 執行檔旁的應用程式根目錄。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[2]


APP_ROOT = _application_root()
DATA_DIR = APP_ROOT / "data"
ASSETS_DIR = APP_ROOT / "assets"

HEALTH_CHECK_ITEMS_FILE = DATA_DIR / "health_check_items.json"
INSTITUTION_DB_FILE = DATA_DIR / "institution_db.json"
SYSTEM_CONFIG_FILE = DATA_DIR / "system_config.json"
ICON_PATH = ASSETS_DIR / "icon.ico"
PANDA_IMAGE_PATH = ASSETS_DIR / "panda.png"
