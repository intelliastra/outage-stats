"""Filter raw outage rows by the exported AE responsibility-reason code."""

from __future__ import annotations

import pandas as pd


PREFIXES = ("513", "514", "515")
CODE_COLUMN = "责任原因代码"
CODE_COLUMN_POSITION = 30  # Excel AE, zero based


def exclude_external_damage_rows(raw: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, int]]:
    columns = [str(name).strip() for name in raw.columns]
    if len(columns) <= CODE_COLUMN_POSITION or columns[CODE_COLUMN_POSITION] != CODE_COLUMN:
        raise ValueError("输入表AE列必须是“责任原因代码”，不能运行未过滤的外力统计")
    frame = raw.copy()
    frame.columns = columns
    codes = frame[CODE_COLUMN].fillna("").astype(str).str.strip()
    matches = {prefix: codes.str.startswith(prefix, na=False) for prefix in PREFIXES}
    remove = matches["513"] | matches["514"] | matches["515"]
    filtered = frame.loc[~remove].copy().reset_index(drop=True)
    counts = {
        "original_rows": len(frame),
        "removed_513": int(matches["513"].sum()),
        "removed_514": int(matches["514"].sum()),
        "removed_515": int(matches["515"].sum()),
        "removed_total": int(remove.sum()),
        "remaining_rows": len(filtered),
        "blank_code_rows": int(codes.eq("").sum()),
    }
    assert counts["original_rows"] == counts["remaining_rows"] + counts["removed_total"]
    return filtered, counts
