"""與介面無關的排序工具。"""

from __future__ import annotations

import re


def safe_natural_sort_key(value):
    if not value:
        return []
    parts = []
    for token in re.split(r"(\d+)", str(value)):
        if not token:
            continue
        if token.isdigit():
            parts.append((0, int(token)))
        else:
            parts.append((1, token.lower()))
    return parts


def safe_chinese_sort_key(value):
    if not value:
        return (0,)
    key = []
    for char in str(value):
        try:
            key.extend(list(char.encode("big5")))
        except Exception:
            key.extend([255, 255, ord(char)])
    return tuple(key)
