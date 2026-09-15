"""Exclusive user/feeder classification for the official statistics tables."""

from __future__ import annotations

from collections import Counter, defaultdict
from typing import Mapping

import pandas as pd


RULE_VERSION = "2026-09-16-priority-v1"
FREQUENT_RULES = (
    ("一年内停电次数超过5次", "年>5次"),
    ("连续60天停电次数超过3次", "60天>3次"),
    ("一年内预安排停电次数超过3次", "预安排>3次"),
)
WARNING_RULES = (
    ("近30天停电2次", "30天2次"),
    ("近50天停电3次", "50天3次"),
    ("一年内停电4-5次", "一年4-5次"),
)


def classify_text(value: object, rules: tuple[tuple[str, str], ...]) -> str:
    text = "" if pd.isna(value) else str(value)
    return next((category for token, category in rules if token in text), "")


def _normalized(value: object) -> str:
    if pd.isna(value):
        return ""
    result = str(value).strip()
    return result[:-2] if result.endswith(".0") and result[:-2].isdigit() else result


def classify_entity_groups(
    frequent: pd.DataFrame,
    warning: pd.DataFrame,
    *,
    key_column: str,
    city_column: str = "所属地市",
    nea_only: bool = False,
) -> dict[str, dict[str, int]]:
    """Assign each stable entity to one group/category, then count by city.

    All multi-hit labels remain in source sheets. A frequent entity is removed
    from the warning group. Conflicting cities use the most frequent observed
    city, breaking ties lexically; this is logged rather than counted twice.
    """
    entities: dict[str, dict] = {}
    for group_name, frame in (("frequent", frequent), ("warning", warning)):
        if frame.empty:
            continue
        if key_column not in frame or city_column not in frame:
            raise ValueError(f"统计清单缺少 {key_column} 或 {city_column}")
        for _, row in frame.iterrows():
            key = _normalized(row[key_column])
            if not key:
                continue
            entry = entities.setdefault(key, {
                "cities": Counter(), "frequent": set(), "warning": set(), "nea": False,
            })
            entry["cities"][_normalized(row[city_column]) or "未归属"] += 1
            entry["nea"] |= _normalized(row.get("是否纳入国家能源局台账", "")) == "是"
            label_col = "频繁停电类型" if group_name == "frequent" else "停电预警类型"
            value = row.get(label_col, "")
            text = "" if pd.isna(value) else str(value)
            for token, category in FREQUENT_RULES if group_name == "frequent" else WARNING_RULES:
                if token in text:
                    entry[group_name].add(category)

    result: dict[str, dict[str, int]] = defaultdict(lambda: defaultdict(int))
    conflict_count = 0
    for entry in entities.values():
        if nea_only and not entry["nea"]:
            continue
        candidates = FREQUENT_RULES if entry["frequent"] else WARNING_RULES
        group_name = "frequent" if entry["frequent"] else "warning"
        category = next((name for _, name in candidates if name in entry[group_name]), "")
        if not category:
            continue
        cities: Counter = entry["cities"]
        if len(cities) > 1:
            conflict_count += 1
        city = sorted(cities, key=lambda name: (-cities[name], name))[0]
        result[f"{group_name}:{category}"][city] += 1
    if conflict_count:
        print(f"[WARN] {key_column} 存在 {conflict_count} 个跨地市归属冲突，已按主归属计一次")
    return {category: dict(counts) for category, counts in result.items()}


def total_city_counts(*sections: Mapping[str, int]) -> dict[str, int]:
    totals: dict[str, int] = defaultdict(int)
    for section in sections:
        for city, count in section.items():
            totals[city] += int(count)
    return dict(totals)


def official_section(groups: Mapping[str, Mapping[str, int]], prefix: str, labels: tuple[str, ...]) -> list[dict[str, int]]:
    parts = [dict(groups.get(f"{prefix}:{label}", {})) for label in labels]
    return parts + [total_city_counts(*parts)]
