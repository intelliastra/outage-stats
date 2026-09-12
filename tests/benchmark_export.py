"""Reproducible synthetic benchmark for the 254,683 x 48 export target."""

from __future__ import annotations

import argparse
import sys
import tempfile
import time
from pathlib import Path

import pandas as pd


SCRIPT_DIR = Path(__file__).resolve().parents[1] / "统计材料" / "运行脚本"
sys.path.insert(0, str(SCRIPT_DIR))

from report_io import write_tables_streaming


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--rows", type=int, default=254_683)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    rows = args.rows
    data = {
        "工单号": [f"WO-{index:08d}" for index in range(rows)],
        "用户编码": [f"{index:012d}" for index in range(rows)],
        "所属馈线编码": [f"F-{index % 5000:05d}" for index in range(rows)],
        "停电总次数": [index % 7 for index in range(rows)],
        "停电开始时间": pd.date_range("2026-01-01", periods=rows, freq="min"),
    }
    for index in range(43):
        data[f"业务字段{index + 1:02d}"] = [f"值-{index}-{row % 100}" for row in range(rows)]
    frame = pd.DataFrame(data)
    destination = args.output or Path(tempfile.gettempdir()) / "outage-stats-export-benchmark.xlsx"
    started = time.perf_counter()
    write_tables_streaming({"用户停电总次数统计表": frame}, destination)
    elapsed = time.perf_counter() - started
    print(f"BENCHMARK rows={rows} cols={len(frame.columns)} elapsed={elapsed:.2f}s file={destination}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
