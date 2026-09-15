"""Compare a completed daily report with the preceding completed report."""

from __future__ import annotations

from collections import Counter, defaultdict
import json
from pathlib import Path
from typing import Any

import pandas as pd
import openpyxl

from statistics_priority import FREQUENT_RULES, WARNING_RULES, RULE_VERSION


DETAIL_COLUMNS = [
    "对象类型", "用户或馈线编码", "所属地市", "上次停电次数", "本次停电次数",
    "上次分类", "本次分类", "变化说明", "上次来源定位", "本次来源定位",
]


def _id(value: object) -> str:
    if pd.isna(value):
        return ""
    value = str(value).strip()
    return value[:-2] if value.endswith(".0") and value[:-2].isdigit() else value


def _number(value: object) -> int:
    parsed = pd.to_numeric(value, errors="coerce")
    return 0 if pd.isna(parsed) else int(parsed)


def _snapshot(tables: dict[str, pd.DataFrame], kind: str, *, legacy_warning_priority: bool = False) -> dict[str, dict[str, Any]]:
    key_col = "用户编码" if kind == "用户" else "所属馈线编码"
    frequent_name = "频繁停电用户清单" if kind == "用户" else "频繁停电线路清单"
    warning_name = "停电预警用户清单" if kind == "用户" else "停电预警线路清单"
    members: dict[str, dict[str, Any]] = {}
    for group, sheet_name in (("frequent", frequent_name), ("warning", warning_name)):
        frame = tables.get(sheet_name, pd.DataFrame())
        if frame.empty:
            continue
        for _, row in frame.iterrows():
            key = _id(row.get(key_col, ""))
            if not key:
                continue
            entry = members.setdefault(key, {
                "categories": {"frequent": set(), "warning": set()},
                "counts": {}, "cities": Counter(),
            })
            entry["cities"][_id(row.get("所属地市", "")) or "未归属"] += 1
            subkey = _id(row.get("所属馈线编码", "")) if kind == "用户" else key
            entry["counts"][subkey] = max(entry["counts"].get(subkey, 0), _number(row.get("停电总次数", 0)))
            label = "频繁停电类型" if group == "frequent" else "停电预警类型"
            text = _id(row.get(label, ""))
            for token, category in FREQUENT_RULES if group == "frequent" else WARNING_RULES:
                if token in text:
                    entry["categories"][group].add(category)
    snapshot: dict[str, dict[str, Any]] = {}
    for key, entry in members.items():
        frequent = entry["categories"]["frequent"]
        warning = entry["categories"]["warning"]
        if not frequent and not warning:
            continue
        priority = FREQUENT_RULES if frequent else (tuple(reversed(WARNING_RULES)) if legacy_warning_priority else WARNING_RULES)
        active = frequent if frequent else warning
        category = next((name for _, name in priority if name in active), "")
        city = sorted(entry["cities"], key=lambda name: (-entry["cities"][name], name))[0]
        snapshot[key] = {
            "category": ("频繁/" if frequent else "预警/") + category,
            "count": sum(entry["counts"].values()), "city": city,
            "frequent": bool(frequent),
            "multi_hit": len(frequent) > 1 or len(warning) > 1 or bool(frequent and warning),
        }
    return snapshot


def _source_locations(raw: pd.DataFrame, kind: str, keys: set[str]) -> dict[str, str]:
    if raw.empty or not keys:
        return {}
    key_col = "用户编码" if kind == "用户" else "所属馈线编码"
    if key_col not in raw.columns:
        return {}
    lookup = raw[key_col].map(_id)
    subset = raw.loc[lookup.isin(keys)]
    result: dict[str, list[str]] = defaultdict(list)
    for _, row in subset.iterrows():
        key = _id(row.get(key_col, ""))
        locator = "工单{order}/用户{user}/馈线{feeder}/停电{start}".format(
            order=_id(row.get("工单号", "")), user=_id(row.get("用户编码", "")),
            feeder=_id(row.get("所属馈线编码", "")),
            start=_id(row.get("去重后停电开始时间", row.get("停电开始时间", ""))),
        )
        if locator not in result[key] and len(result[key]) < 5:
            result[key].append(locator)
    return {key: "；".join(values) for key, values in result.items()}


def completed_baseline(path: Path | None, data_year: int) -> bool:
    if not path or not path.is_file():
        return False
    stem = path.stem.replace("停电用户_处理结果", "")
    required = [
        path.with_name(f"{stem}停电用户_处理结果_发出版.xlsx"),
        path.with_name(f"{stem}统计表.xlsx"),
        path.with_name(f"{stem}停电摘要（简版）.txt"),
        path.with_name(f"{stem}停电摘要（全量版）.txt"),
    ]
    if not all(file.is_file() and file.stat().st_size > 0 for file in required):
        return False
    try:
        dates = pd.read_excel(path, sheet_name="用户停电总次数统计表", usecols=lambda c: c in ("去重后停电开始时间", "停电开始时间"), nrows=1000)
        for col in ("去重后停电开始时间", "停电开始时间"):
            if col in dates:
                years = pd.to_datetime(dates[col], errors="coerce").dropna().dt.year.unique()
                return len(years) == 0 or all(int(year) == data_year for year in years)
    except (ValueError, OSError):
        return False
    return False


def compare_reports(
    current: dict[str, pd.DataFrame], previous_file: Path | None,
    current_raw: pd.DataFrame, previous_raw: pd.DataFrame,
    *, data_year: int, current_source: str,
    current_stats_dicts: list[dict[str, int]] | None = None,
) -> tuple[dict[str, Any], pd.DataFrame, dict[str, str]]:
    if not completed_baseline(previous_file, data_year):
        return ({"available": False, "reason": "没有同年份已完成的上次日报", "rule_version": RULE_VERSION},
                pd.DataFrame(columns=DETAIL_COLUMNS), {})
    previous: dict[str, pd.DataFrame] = {}
    try:
        with pd.ExcelFile(previous_file) as excel:
            for sheet in ("频繁停电用户清单", "停电预警用户清单", "频繁停电线路清单", "停电预警线路清单"):
                previous[sheet] = pd.read_excel(excel, sheet_name=sheet, dtype=str)
    except (ValueError, OSError) as exc:
        return ({"available": False, "reason": f"上次日报清单不可读取：{exc}", "rule_version": RULE_VERSION},
                pd.DataFrame(columns=DETAIL_COLUMNS), {})
    old_sidecar = previous_file.with_name(f"{previous_file.stem}_比较.json")
    previous_rule_version = "未记录（历史报告）"
    if old_sidecar.is_file():
        try:
            previous_rule_version = json.loads(old_sidecar.read_text(encoding="utf-8")).get(
                "rule_version", previous_rule_version
            )
        except (OSError, ValueError):
            pass
    old_rule_is_legacy = previous_rule_version != RULE_VERSION
    stats_columns: list[dict[str, Any]] = []
    if current_stats_dicts is not None:
        stem = previous_file.stem.replace("停电用户_处理结果", "")
        stats_path = previous_file.with_name(f"{stem}统计表.xlsx")
        try:
            workbook = openpyxl.load_workbook(stats_path, read_only=True, data_only=True)
            try:
                sheet = workbook["2026年统计表"]
                old_totals = [_number(sheet.cell(3, col).value) for col in range(2, 22)]
            finally:
                workbook.close()
            column_names = (
                "预警用户年4-5", "预警用户50天3", "预警用户30天2", "预警用户合计",
                "预警线路年4-5", "预警线路50天3", "预警线路30天2", "预警线路合计",
                "频繁用户年>5", "频繁用户60天>3", "频繁用户预安排>3", "频繁用户合计",
                "频繁线路年>5", "频繁线路60天>3", "频繁线路预安排>3", "频繁线路合计",
                "NEA线路年>5", "NEA线路60天>3", "NEA线路预安排>3", "NEA线路合计",
            )
            for index, (label, prior, section) in enumerate(zip(column_names, old_totals, current_stats_dicts)):
                now = sum(int(value) for value in section.values())
                stats_columns.append({"column": chr(66 + index), "label": label,
                                      "previous": prior, "current": now, "delta": now - prior})
        except (OSError, ValueError, KeyError):
            stats_columns = []
    rows: list[dict[str, Any]] = []
    user_explanations: dict[str, str] = {}
    metrics: dict[str, Any] = {}
    for kind in ("用户", "线路"):
        old = _snapshot(previous, kind, legacy_warning_priority=old_rule_is_legacy)
        new = _snapshot(current, kind)
        freq_old = {key for key, value in old.items() if value["frequent"]}
        freq_new = {key for key, value in new.items() if value["frequent"]}
        metrics[kind] = {
            "previous_frequent": len(freq_old), "current_frequent": len(freq_new),
            "net_change": len(freq_new) - len(freq_old),
            "entered_frequent": len(freq_new - freq_old),
            "left_frequent": len(freq_old - freq_new),
            "decreased_outage_count": sum(
                new[key]["count"] < old[key]["count"] for key in old.keys() & new.keys()
            ),
        }
        total_index = 11 if kind == "用户" else 15
        if stats_columns:
            total = stats_columns[total_index]
            metrics[kind].update({
                "report_total_previous": total["previous"],
                "report_total_current": total["current"],
                "report_total_delta": total["delta"],
            })
        changed_keys = {
            key for key in old.keys() | new.keys()
            if key not in old or key not in new
            or old[key]["count"] != new[key]["count"]
            or old[key]["category"] != new[key]["category"]
            or (old_rule_is_legacy and old[key]["multi_hit"])
        }
        old_locs = _source_locations(previous_raw, kind, changed_keys)
        new_locs = _source_locations(current_raw, kind, changed_keys)
        for key in sorted(changed_keys):
            prior, after = old.get(key), new.get(key)
            if prior is None:
                description = "本次新增；请核对新增日期或重导记录"
            elif after is None:
                description = "本次减少；上次记录或分类已不在本次结果，需核对重导及剔除依据"
            elif after["count"] < prior["count"]:
                description = f"停电次数减少 {prior['count'] - after['count']} 次；请核对记录修订、剔除及口径"
            elif after["count"] > prior["count"]:
                description = f"停电次数增加 {after['count'] - prior['count']} 次；请核对新增日期及修订记录"
            else:
                description = "停电次数未变，分类变化；请核对统计规则版本"
            if prior and after and prior["category"] != after["category"]:
                description += f"；{prior['category']}→{after['category']}"
            if prior and after and old_rule_is_legacy and prior["multi_hit"]:
                description += "；上次多规则命中可能重复计列，本次按最高优先级只计一次"
            rows.append({
                "对象类型": kind, "用户或馈线编码": key,
                "所属地市": (after or prior)["city"],
                "上次停电次数": prior["count"] if prior else 0,
                "本次停电次数": after["count"] if after else 0,
                "上次分类": prior["category"] if prior else "",
                "本次分类": after["category"] if after else "",
                "变化说明": description,
                "上次来源定位": old_locs.get(key) or (
                    f"旧报告{previous_file.name}/" +
                    ("频繁停电用户清单" if kind == "用户" else "频繁停电线路清单") +
                    f"/{key}" if prior else ""
                ),
                "本次来源定位": new_locs.get(key) or (f"本次用户或馈线编码/{key}" if after else ""),
            })
            if kind == "用户" and after:
                user_explanations[key] = description
    comparison = {
        "available": True, "previous_report": str(previous_file),
        "current_source": current_source, "rule_version": RULE_VERSION,
        "previous_rule_version": previous_rule_version,
        "metrics": metrics, "changed_entities": len(rows),
        "statistics_columns": stats_columns,
    }
    return comparison, pd.DataFrame(rows, columns=DETAIL_COLUMNS), user_explanations
