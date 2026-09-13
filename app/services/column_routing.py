"""院所專屬分派規則的解析與欄位比對。"""

from __future__ import annotations

import re
from collections.abc import Iterable, Mapping


def parse_routing_rules(raw_value: object) -> list[str]:
    """把使用者輸入的逗號清單轉成去重、保留順序的規則列表。"""
    rules: list[str] = []
    seen: set[str] = set()
    for value in str(raw_value or "").replace("，", ",").split(","):
        rule = value.strip()
        key = _normalize(rule)
        if rule and key and key not in seen:
            rules.append(rule)
            seen.add(key)
    return rules


def column_routing_values(worksheet, column_index: int, base_header_row: int, top_tag_row: int) -> dict[str, str]:
    """讀出動態欄位可供規則比對的四種資料。"""
    return {
        "item_name": _cell_text(worksheet, max(1, base_header_row - 2), column_index),
        "system_code": _cell_text(worksheet, max(1, base_header_row - 1), column_index),
        "english_code": _cell_text(worksheet, base_header_row, column_index),
        "org_tag": _cell_text(worksheet, top_tag_row, column_index),
    }


def find_matching_rule(rules: Iterable[str], values: Mapping[str, str]) -> str | None:
    """找出第一個命中的規則。

    「系統:105」、「英文:GLU」、「名稱:血糖」這類精準規則會只對指定
    資料庫欄位做完整比對。舊版自由文字規則仍可讀取，但可能部分命中多個
    項目，設定畫面會要求改用精準規則後才能儲存。
    """
    normalized_values = {field: _normalize(value) for field, value in values.items()}
    compact_values = {field: _compact(value) for field, value in values.items()}

    for rule in rules:
        normalized_rule = _normalize(rule)
        compact_rule = _compact(rule)
        if not normalized_rule:
            continue

        explicit_field, explicit_value = _parse_explicit_rule(rule)
        if explicit_field:
            if (
                explicit_value == normalized_values.get(explicit_field, "")
                or _compact(explicit_value) == compact_values.get(explicit_field, "")
            ):
                return rule
            continue

        if normalized_rule in normalized_values.values() or compact_rule in compact_values.values():
            return rule

        if _allows_partial_match(compact_rule):
            for field in ("english_code", "item_name"):
                if compact_rule in compact_values.get(field, ""):
                    return rule
    return None


def resolve_forced_destination(
    force_to_xing: Iterable[str],
    force_to_boren: Iterable[str],
    values: Mapping[str, str],
) -> tuple[str | None, str | None, str | None]:
    """依院所規則決定項目目的地；同時命中時杏聯優先。"""
    xing_rule = find_matching_rule(force_to_xing, values)
    boren_rule = find_matching_rule(force_to_boren, values)
    if xing_rule:
        return "xing", xing_rule, boren_rule
    if boren_rule:
        return "boren", None, boren_rule
    return None, None, None


def analyze_routing_rules(
    force_to_xing: object,
    force_to_boren: object,
    items: Iterable[Mapping[str, object]],
) -> dict[str, object]:
    """分析院所分派規則，供設定畫面在儲存前預覽與阻擋衝突。

    回傳值保留原始資料庫項目，UI 可自行用系統代碼、英文簡碼或完整名稱
    呈現。``unmatched_rules`` 與 ``conflicts`` 應視為錯誤；
    ``ambiguous_rules`` 表示部分規則同時命中多個項目，適合作為警告顯示。
    """
    rules_by_destination = {
        "xing": _coerce_routing_rules(force_to_xing),
        "boren": _coerce_routing_rules(force_to_boren),
    }
    source_items = list(items)
    rule_matches: dict[str, list[dict[str, object]]] = {"xing": [], "boren": []}
    unmatched_rules: list[dict[str, str]] = []
    ambiguous_rules: list[dict[str, object]] = []

    for destination, rules in rules_by_destination.items():
        for rule in rules:
            matches = [
                item
                for item in source_items
                if find_matching_rule([rule], database_item_routing_values(item))
            ]
            match_record: dict[str, object] = {
                "destination": destination,
                "rule": rule,
                "matches": matches,
            }
            rule_matches[destination].append(match_record)
            if not matches:
                unmatched_rules.append({"destination": destination, "rule": rule})
            elif len(matches) > 1:
                ambiguous_rules.append(match_record)

    item_destinations: list[dict[str, object]] = []
    conflicts: list[dict[str, object]] = []
    for item in source_items:
        values = database_item_routing_values(item)
        destination, xing_rule, boren_rule = resolve_forced_destination(
            rules_by_destination["xing"],
            rules_by_destination["boren"],
            values,
        )
        item_record: dict[str, object] = {
            "item": item,
            "destination": destination,
            "xing_rule": xing_rule,
            "boren_rule": boren_rule,
        }
        if destination:
            item_destinations.append(item_record)
        if xing_rule and boren_rule:
            conflicts.append(item_record)

    return {
        "rules": rules_by_destination,
        "rule_matches": rule_matches,
        "unmatched_rules": unmatched_rules,
        "ambiguous_rules": ambiguous_rules,
        "item_destinations": item_destinations,
        "conflicts": conflicts,
        "has_errors": bool(unmatched_rules or conflicts),
        "has_warnings": bool(ambiguous_rules),
    }


def database_item_routing_values(item: Mapping[str, object]) -> dict[str, str]:
    """將資料庫項目轉成與 Excel 欄位相同的可比對欄位。"""
    return {
        "item_name": str(item.get("full_name", "") or ""),
        "system_code": str(item.get("internal_code", "") or ""),
        "english_code": str(item.get("sys_code", "") or ""),
        "org_tag": str(item.get("org_display", "") or ""),
    }


def make_precise_routing_rule(item: Mapping[str, object]) -> str:
    """依資料庫項目建立可穩定保存的精準分派規則。

    優先使用系統代碼；沒有系統代碼時才使用英文簡碼，最後才使用完整名稱。
    這些前綴也能讓日後讀取設定時知道要比對哪一個欄位。
    """
    internal_code = str(item.get("internal_code", "") or "").strip()
    if internal_code:
        return f"系統:{internal_code}"
    english_code = str(item.get("sys_code", "") or "").strip()
    if english_code:
        return f"英文:{english_code}"
    return f"名稱:{str(item.get('full_name', '') or '').strip()}"


def override_org_display(original_org: object, destination: str | None) -> tuple[str, bool]:
    """依目的地覆寫杏／博標籤，並保留原本的分組數字（例如 杏2 → 博2）。"""
    original = str(original_org or "")
    if destination not in {"xing", "boren"} or not original:
        return original, False

    target_prefix = "杏" if destination == "xing" else "博"
    if original[0] not in {"杏", "博"}:
        return original, False

    effective = target_prefix + original[1:]
    return effective, effective != original


def _cell_text(worksheet, row: int, column: int) -> str:
    value = worksheet.cell(row=row, column=column).value
    return "" if value is None else str(value).strip()


def _normalize(value: object) -> str:
    text = str(value or "").strip().casefold()
    # Excel 有時會把數字系統代碼讀成 225.0；這應視為使用者輸入的 225。
    numeric_match = re.fullmatch(r"(\d+)\.0+", text)
    if numeric_match:
        return numeric_match.group(1)
    return re.sub(r"\s+", "", text)


def _compact(value: object) -> str:
    return re.sub(r"[\s\u3000_\-‐‑–—/\\()（）.]+", "", _normalize(value))


def _allows_partial_match(compact_rule: str) -> bool:
    if len(compact_rule) < 2:
        return False
    return any(char.isalpha() or "\u4e00" <= char <= "\u9fff" for char in compact_rule)


def _parse_explicit_rule(rule: object) -> tuple[str | None, str]:
    """解析精準規則的欄位前綴；未知前綴保留給舊版自由文字規則。"""
    text = str(rule or "").strip()
    if ":" not in text and "：" not in text:
        return None, ""

    prefix, value = re.split(r"[:：]", text, maxsplit=1)
    field_by_prefix = {
        "系統": "system_code",
        "代碼": "system_code",
        "system": "system_code",
        "code": "system_code",
        "英文": "english_code",
        "english": "english_code",
        "en": "english_code",
        "名稱": "item_name",
        "name": "item_name",
        "item": "item_name",
    }
    field = field_by_prefix.get(_normalize(prefix))
    return field, _normalize(value)


def _coerce_routing_rules(raw_value: object) -> list[str]:
    """兼容設定畫面的文字與呼叫端已解析好的規則清單。"""
    if raw_value is None or isinstance(raw_value, str):
        return parse_routing_rules(raw_value)
    if isinstance(raw_value, Iterable):
        return parse_routing_rules(",".join(str(value) for value in raw_value))
    return parse_routing_rules(raw_value)
