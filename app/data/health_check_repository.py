"""團檢項目 JSON 資料庫的讀取、查找與排序。"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Callable

from app.config import HEALTH_CHECK_ITEMS_FILE
from app.utils.sorting import safe_chinese_sort_key, safe_natural_sort_key


master_db: dict = {}
all_item_names: list[str] = []


def find_db_item_robust(display_label):
    clean_label = display_label.split(" | ")[0].strip()
    search_key = clean_label.replace(" ", "").replace("\u3000", "").upper()
    if search_key in master_db:
        return master_db[search_key]

    org_match = re.search(r"^\[(.*?)\]", clean_label)
    org_name = org_match.group(1).strip() if org_match else ""
    en_match = re.search(r"\((.*?)\)$", clean_label.strip())
    english_name = en_match.group(1).strip() if en_match else ""

    internal_code = ""
    if "] " in clean_label:
        parts = clean_label.split("] ", 1)
        if " - " in parts[1]:
            internal_code = parts[1].split(" - ", 1)[0].strip()

    clean_name = clean_label
    if org_name:
        clean_name = re.sub(r"^\[.*?\]\s*", "", clean_name)
    if english_name:
        clean_name = re.sub(r"\s*\(.*?\)$", "", clean_name)
    if internal_code and clean_name.startswith(internal_code):
        clean_name = clean_name[len(internal_code):].strip()
        if clean_name.startswith("-"):
            clean_name = clean_name[1:].strip()

    if internal_code:
        for item in master_db.values():
            if item["internal_code"] == internal_code:
                return item
    if english_name:
        for item in master_db.values():
            if item["sys_code"] == english_name:
                if not org_name or org_name[0] in item["org_display"] or item["org_display"][0] in org_name:
                    return item
    if clean_name:
        for item in master_db.values():
            if item["full_name"] == clean_name:
                if not org_name or org_name[0] in item["org_display"] or item["org_display"][0] in org_name:
                    return item
    if clean_name:
        for item in master_db.values():
            if clean_name in item["full_name"] or item["full_name"] in clean_name:
                return item
    return None


# 過渡期間保留舊名稱，讓從舊單檔搬出的函式不需改動其查詢邏輯。
def find_db_item_robust_global(display_label):
    return find_db_item_robust(display_label)


def get_sorted_items_list(items_list, sort_by):
    if sort_by == "預設 (原始順序)":
        return sorted(items_list, key=lambda item: all_item_names.index(item) if item in all_item_names else 999999)

    def sort_key(label):
        item = find_db_item_robust(label)
        if not item:
            internal_code = ""
            clean_label = label.split(" | ")[0].strip()
            if "] " in clean_label:
                parts = clean_label.split("] ", 1)
                if " - " in parts[1]:
                    internal_code = parts[1].split(" - ", 1)[0].strip()
            clean_name = clean_label
            org_match = re.search(r"^\[(.*?)\]", clean_label)
            org_name = org_match.group(1).strip() if org_match else ""
            if org_name:
                clean_name = re.sub(r"^\[.*?\]\s*", "", clean_name)
            if internal_code:
                clean_name = re.sub(r"^[A-Za-z0-9]+\s*-\s*", "", clean_name)
            en_match = re.search(r"\((.*?)\)$", clean_label.strip())
            english_name = en_match.group(1).strip() if en_match else ""
            if english_name:
                clean_name = re.sub(r"\s*\(.*?\)$", "", clean_name)
            item = {
                "internal_code": internal_code,
                "full_name": clean_name.strip(),
                "sys_code": english_name,
                "org_display": org_name,
                "category": "",
            }

        internal_code = item.get("internal_code", "") or ""
        full_name = item.get("full_name", "") or ""
        sys_code = item.get("sys_code", "") or ""
        org_display = item.get("org_display", "") or ""
        category = item.get("category", "") or ""
        if sort_by == "系統代碼":
            return safe_natural_sort_key(internal_code), safe_chinese_sort_key(full_name)
        if sort_by in ["中文名稱", "中/英文名稱"]:
            return safe_chinese_sort_key(full_name), safe_natural_sort_key(internal_code)
        if sort_by == "英文簡寫":
            return safe_natural_sort_key(sys_code), safe_natural_sort_key(internal_code)
        if sort_by == "歸屬單位":
            return safe_natural_sort_key(org_display), safe_natural_sort_key(internal_code)
        if sort_by == "檢驗類別":
            return safe_chinese_sort_key(category), safe_natural_sort_key(internal_code)
        return safe_natural_sort_key(internal_code), safe_chinese_sort_key(label)

    return sorted(items_list, key=sort_key)


def load_health_check_database(
    silent: bool = True,
    *,
    file_path: Path = HEALTH_CHECK_ITEMS_FILE,
    on_warning: Callable[[str, str], None] | None = None,
    on_error: Callable[[str, str], None] | None = None,
):
    if not file_path.exists():
        message = "尚未建立大腦資料庫！請至【設定 > 資料庫維護】匯入 Excel 總表。"
        if not silent and on_warning:
            on_warning("提示", message)
        return False, message

    try:
        with file_path.open("r", encoding="utf-8") as file:
            raw_data = json.load(file)
        master_db.clear()
        all_item_names.clear()
        for sys_code, info in raw_data.items():
            full_name = info.get("中文名稱", "").strip()
            en_name = info.get("英文簡寫", "").strip()
            internal_code = info.get("系統代碼", "").strip()
            org_raw = info.get("歸屬單位", "C2").strip().upper()
            category = info.get("檢驗類別", "").strip()
            org_display = org_raw.replace("C", "杏").replace("B", "博")
            item_info = {
                "full_name": full_name,
                "sys_code": en_name,
                "internal_code": internal_code,
                "org_display": org_display,
                "category": category,
            }
            if internal_code:
                base_label = f"[{org_display}] {internal_code} - {full_name} ({en_name})"
            else:
                base_label = f"[{org_display}] {full_name} ({en_name})"
            display_label = f"{base_label} | {category}" if category else base_label
            if display_label not in all_item_names:
                all_item_names.append(display_label)
            key_with_category = display_label.replace(" ", "").replace("\u3000", "").upper()
            key_without_category = base_label.replace(" ", "").replace("\u3000", "").upper()
            master_db[key_with_category] = item_info
            master_db[key_without_category] = item_info
        return True, f"🎯 成功讀取 JSON 大腦 | 📊 共載入 {len(master_db)} 項目"
    except Exception as exc:
        if not silent and on_error:
            on_error("錯誤", f"讀取失敗：{str(exc)}")
        return False, f"無法成功讀取 JSON，錯誤：{exc}"
