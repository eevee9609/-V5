"""日期文字的正規化與民國年轉換。"""

from __future__ import annotations

import re

import pandas as pd


def normalize_date_text(date_value):
    if date_value is None:
        return ""
    if hasattr(date_value, "strftime"):
        try:
            return date_value.strftime("%Y/%m/%d")
        except Exception:
            pass
    if isinstance(date_value, (int, float)) and pd.isna(date_value):
        return ""
    value = str(date_value).strip()
    if not value or value.lower() in {"nan", "none"}:
        return ""
    value = value.replace(".", "/").replace("-", "/")
    value = re.sub(r"\s+", "", value)
    if re.fullmatch(r"\d{8}", value):
        return f"{value[:4]}/{value[4:6]}/{value[6:8]}"
    if re.fullmatch(r"\d{6,7}", value):
        return f"{value[:-4]}/{value[-4:-2]}/{value[-2:]}"
    return value


def convert_to_taiwan_date(date_str):
    normalized = normalize_date_text(date_str)
    if not normalized:
        return ""
    value = normalized.strip().replace(".", "/").replace("-", "/")
    parts = value.split("/")
    if len(parts) >= 3:
        try:
            year, month, day = int(parts[0]), int(parts[1]), int(parts[2])
            if year > 1911:
                return f"{year - 1911}{str(month).zfill(2)}{str(day).zfill(2)}"
            return f"{year}{str(month).zfill(2)}{str(day).zfill(2)}"
        except ValueError:
            pass
    matches = re.findall(r"\d+", value)
    if len(matches) >= 3:
        try:
            year, month, day = int(matches[0]), int(matches[1]), int(matches[2])
            if year > 1911:
                return f"{year - 1911}{str(month).zfill(2)}{str(day).zfill(2)}"
            return f"{year}{str(month).zfill(2)}{str(day).zfill(2)}"
        except ValueError:
            pass
    return normalized
