"""Utilities for historical period statistics (year + month/day range)."""

from __future__ import annotations

import calendar
import re
import shutil
from datetime import date
from pathlib import Path

HISTORICAL_YEARS = (2025, 2026)

EXCLUDE_2025_GLOBS = (
    "表11重大事件日统计_20260629091427858*.xlsx",
    "表11重大事件日统计_20260629091427858*",
)
EXCLUDE_2026_NAME = "表11重大事件日统计表_20260629091136914.xlsx"

ANNUAL_2025_REL = Path("input") / "Annual Summary" / "2025" / "2025年停电用户数据.xlsx"


def resolve_period_dates(
    year: int,
    start_month: int,
    start_day: int,
    end_month: int,
    end_day: int,
) -> tuple[date, date]:
    """Resolve period dates; allow cross-year when end month/day is before start.

    Rules:
    - Start year is always ``year``.
    - If end (month, day) < start (month, day), end year = year + 1 (cross-year).
    - Otherwise end year = year.
    - Explicit cross-year via end after Dec still uses year for start.
    """
    if year not in HISTORICAL_YEARS:
        raise ValueError(f"仅支持年份 {HISTORICAL_YEARS}，收到：{year}")
    for label, month, day in (
        ("起始", start_month, start_day),
        ("结束", end_month, end_day),
    ):
        if month < 1 or month > 12:
            raise ValueError(f"{label}月份无效：{month}")
        max_day = calendar.monthrange(year if label == "起始" else year, month)[1]
        # For end month in next year leap-check, recompute below after end_year known
        if day < 1 or day > 31:
            raise ValueError(f"{label}日期无效：{day}")

    start_year = year
    if (end_month, end_day) < (start_month, start_day):
        end_year = year + 1
    else:
        end_year = year

    start_max = calendar.monthrange(start_year, start_month)[1]
    end_max = calendar.monthrange(end_year, end_month)[1]
    if start_day > start_max:
        raise ValueError(f"起始日期无效：{start_year}-{start_month}-{start_day}")
    if end_day > end_max:
        raise ValueError(f"结束日期无效：{end_year}-{end_month}-{end_day}")

    start = date(start_year, start_month, start_day)
    end = date(end_year, end_month, end_day)
    if end < start:
        raise ValueError("结束日期不得早于起始日期")
    return start, end


def annual_2025_target(stats_base: Path) -> Path:
    return (stats_base / ANNUAL_2025_REL).resolve()


def ensure_annual_2025_file(stats_base: Path) -> Path:
    """Ensure Annual Summary/2025/2025年停电用户数据.xlsx exists; search & copy if needed."""
    target = annual_2025_target(stats_base)
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.is_file() and target.stat().st_size > 0:
        return target

    preferred = stats_base / "2025年停电用户数据.xlsx"
    candidates: list[Path] = []
    if preferred.is_file():
        candidates.append(preferred)

    for path in stats_base.rglob("*.xlsx"):
        if path.name.startswith("~$"):
            continue
        if path.resolve() == target:
            continue
        name = path.name
        if "2025" in name and "停电" in name:
            candidates.append(path)

    # Prefer exact name, then shorter path under stats_base root
    def sort_key(p: Path) -> tuple[int, int, str]:
        exact = 0 if p.name == "2025年停电用户数据.xlsx" else 1
        return (exact, len(p.parts), str(p))

    candidates = sorted({c.resolve() for c in candidates}, key=sort_key)
    if not candidates:
        raise FileNotFoundError(
            "未找到 2025 年停电用户数据.xlsx，请放入 统计材料/ 或 "
            "统计材料/input/Annual Summary/2025/"
        )

    source = candidates[0]
    shutil.copy2(source, target)
    return target


def find_latest_newdata_file(stats_base: Path) -> Path | None:
    newdata = stats_base / "input" / "Newdata" / "historical"
    if not newdata.is_dir():
        return None
    files = [
        f
        for f in list(newdata.glob("*.xlsx")) + list(newdata.glob("*.xls"))
        if not f.name.startswith("~$") and f.stat().st_size > 0
    ]
    if not files:
        return None

    date_re = re.compile(r"(20\d{2})[.\-](\d{1,2})[.\-](\d{1,2})")

    def file_key(path: Path) -> tuple[int, float]:
        match = date_re.search(path.stem)
        if match:
            try:
                d = date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
                return (1, d.toordinal())
            except ValueError:
                pass
        return (0, path.stat().st_mtime)

    return max(files, key=file_key)


def resolve_exclude_file(stats_base: Path, year: int) -> Path | None:
    """Resolve default table-11 exclude file for historical run."""
    year_dir = stats_base / "input" / "exclude" / str(year)
    if not year_dir.is_dir():
        return None

    if year == 2025:
        for pattern in EXCLUDE_2025_GLOBS:
            matches = [
                p
                for p in year_dir.glob(pattern)
                if p.is_file() and not p.name.startswith("~$")
            ]
            # also match without extension in glob result for .xlsx
            if matches:
                return max(matches, key=lambda p: p.stat().st_mtime)
        # fallback: any 表11 with timestamp fragment
        for p in year_dir.iterdir():
            if p.is_file() and "20260629091427858" in p.name and p.suffix.lower() in {".xlsx", ".xls"}:
                return p
    elif year == 2026:
        preferred = year_dir / EXCLUDE_2026_NAME
        if preferred.is_file():
            return preferred
        for p in year_dir.iterdir():
            if p.is_file() and "20260629091136914" in p.name and p.suffix.lower() in {".xlsx", ".xls"}:
                return p

    # Fallback: latest table11 in year dir
    from exclude_utils import select_latest_table11_file

    return select_latest_table11_file(year_dir)


def build_historical_config(stats_base: Path) -> dict:
    """Config payload for GET /api/historical/config."""
    years_info: dict[str, dict] = {}

    # 2025
    annual_path = annual_2025_target(stats_base)
    annual_ready = annual_path.is_file() and annual_path.stat().st_size > 0
    if not annual_ready:
        # Probe whether bootstrap can succeed without copying yet
        try:
            probe = list(stats_base.glob("2025年停电用户数据.xlsx"))
            probe += [
                p
                for p in stats_base.rglob("*2025*停电*.xlsx")
                if not p.name.startswith("~$")
            ]
            can_bootstrap = any(p.is_file() and p.stat().st_size > 0 for p in probe)
        except OSError:
            can_bootstrap = False
    else:
        can_bootstrap = True

    exclude_2025 = resolve_exclude_file(stats_base, 2025)
    years_info["2025"] = {
        "data_ready": annual_ready or can_bootstrap,
        "data_path": str(annual_path),
        "data_filename": annual_path.name if annual_ready else None,
        "exclude_ready": exclude_2025 is not None,
        "exclude_filename": exclude_2025.name if exclude_2025 else None,
        "source_label": "Annual Summary/2025/2025年停电用户数据.xlsx",
    }

    newdata = find_latest_newdata_file(stats_base)
    exclude_2026 = resolve_exclude_file(stats_base, 2026)
    years_info["2026"] = {
        "data_ready": newdata is not None,
        "data_path": str(newdata) if newdata else str(stats_base / "input" / "Newdata"),
        "data_filename": newdata.name if newdata else None,
        "exclude_ready": exclude_2026 is not None,
        "exclude_filename": exclude_2026.name if exclude_2026 else None,
        "source_label": "input/Newdata 最新 Excel",
    }

    return {
        "years": list(HISTORICAL_YEARS),
        "year_info": years_info,
        "allow_cross_year": True,
        "hint": "预计耗时 20–60 分钟；长时间无日志输出不等于卡死",
    }
