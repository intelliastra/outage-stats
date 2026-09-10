"""Utilities for table-11 exclude file management."""

from __future__ import annotations

import re
from pathlib import Path

_TABLE11_TIMESTAMP_RE = re.compile(r"_(\d{14,17})(?:\.xlsx)?$", re.IGNORECASE)

MIN_EXCLUDE_YEAR = 2024
MAX_EXCLUDE_YEAR = 2030
DEFAULT_EXCLUDE_YEAR = 2026


def validate_exclude_year(year: int) -> int:
    if year < MIN_EXCLUDE_YEAR or year > MAX_EXCLUDE_YEAR:
        raise ValueError(
            f"年份须在 {MIN_EXCLUDE_YEAR}–{MAX_EXCLUDE_YEAR} 之间，收到：{year}"
        )
    return year


def list_exclude_years() -> list[int]:
    return list(range(MIN_EXCLUDE_YEAR, MAX_EXCLUDE_YEAR + 1))


def _table11_filename_timestamp(path: Path) -> int | None:
    match = _TABLE11_TIMESTAMP_RE.search(path.stem)
    if not match:
        return None
    return int(match.group(1))


def select_latest_table11_file(folder: Path) -> Path | None:
    """从目录中选取最新的表11 xlsx；目录为空或不存在时返回 None。"""
    if not folder.is_dir():
        return None

    candidates = [
        f
        for f in list(folder.glob("*.xlsx")) + list(folder.glob("*.xls"))
        if not f.name.startswith("~$")
    ]
    if not candidates:
        return None

    preferred = [
        f for f in candidates
        if "表11" in f.name or "重大事件" in f.name
    ]
    pool = preferred or candidates

    def sort_key(path: Path) -> tuple[int, float, str]:
        ts = _table11_filename_timestamp(path)
        return (
            1 if ts is not None else 0,
            float(ts) if ts is not None else path.stat().st_mtime,
            path.name,
        )

    return max(pool, key=sort_key)


def get_year_latest_file_info(year_dir: Path) -> dict[str, object] | None:
    latest = select_latest_table11_file(year_dir)
    if latest is None:
        return None
    return {
        "filename": latest.name,
        "size": latest.stat().st_size,
    }
