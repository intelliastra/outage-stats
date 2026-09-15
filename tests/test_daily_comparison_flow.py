from __future__ import annotations

import importlib.util
import io
import json
import os
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from datetime import datetime, timedelta
from pathlib import Path

import openpyxl
import pandas as pd

SCRIPT_DIR = Path(__file__).resolve().parents[1] / "统计材料" / "运行脚本"
sys.path.insert(0, str(SCRIPT_DIR))

from report_io import write_tables_streaming


def make_raw(year: int, *, drop_u2_last: bool = False) -> pd.DataFrame:
    rows = []
    for day in range(1, 7):
        start = datetime(year, 1, day, 8)
        for user in ("U1", "U2", "U3"):
            if drop_u2_last and user == "U2" and day >= 4:
                continue
            rows.append({
                "工单号": f"WO{day}", "用户编码": user, "中电联用户名称": user,
                "用户性质": "公用", "停电类型": "故障",
                "去重后停电开始时间": start.strftime("%Y-%m-%d %H:%M:%S"),
                "去重后停电结束时间": (start + timedelta(hours=1)).strftime("%Y-%m-%d %H:%M:%S"),
                "所属馈线编码": f"L{user[-1]}", "所属馈线名称": f"线路{user[-1]}",
                "所属供电所": "供电所", "所属区局": "区局", "所属地市": "广州",
            })
    return pd.DataFrame(rows)


class DailyComparisonFlowTest(unittest.TestCase):
    def test_correction_reduces_frequent_user_and_preserves_previous(self):
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory)
            newdata = base / "input/Newdata/daily"
            newdata.mkdir(parents=True)
            output = base / "output/daily"
            baseline = base / "2025年停电用户数据.xlsx"
            write_tables_streaming({"停电用户数据": make_raw(2025)}, baseline)
            template = base / "统计表（模板）.xlsx"
            workbook = openpyxl.Workbook()
            workbook.active.title = "2026年统计表"
            workbook.create_sheet("2025年统计表")
            for sheet in workbook:
                sheet["A3"] = "公司"
                sheet["A4"] = "广州"
            workbook.save(template)
            major = base / "表11.xlsx"
            write_tables_streaming({"表11": pd.DataFrame({"重大事件日期": ["2026-12-31"]})}, major)
            first = newdata / "停电用户数据_20260916000000001.xlsx"
            second = newdata / "停电用户数据_20260916000000002.xlsx"
            write_tables_streaming({"停电用户数据": make_raw(2026)}, first)
            spec = importlib.util.spec_from_file_location("daily_flow_test", SCRIPT_DIR / "process_outage_tables_v1.0.8.py")
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            module.STATS_TEMPLATE_PATH = str(template)
            module.NEA_LEDGER_PATH = str(base / "missing-ledger.xlsx")
            module.RAW_2025_DATA_PATH = str(baseline)
            module.ANNUAL_2025_PATH = str(baseline)
            module.OUTPUT_BASE_DIR = str(output)
            module.EXISTINGDATA_FOLDER = str(output)
            module.NEWDATA_FOLDER = str(newdata)
            old_argv = sys.argv
            old_cache = os.environ.get("OUTAGE_CACHE_DIR")
            os.environ["OUTAGE_CACHE_DIR"] = str(base / "cache")
            try:
                sys.argv = ["daily", "--non-interactive", "--exclude-2025-file", str(major), "--exclude-2026-file", str(major)]
                with redirect_stdout(io.StringIO()):
                    self.assertEqual(module.main(), 0)
                original = list(output.glob("*/*停电用户_处理结果.xlsx"))
                self.assertEqual(len(original), 1)
                write_tables_streaming({"停电用户数据": make_raw(2026, drop_u2_last=True)}, second)
                with redirect_stdout(io.StringIO()):
                    self.assertEqual(module.main(), 0)
            finally:
                sys.argv = old_argv
                if old_cache is None:
                    os.environ.pop("OUTAGE_CACHE_DIR", None)
                else:
                    os.environ["OUTAGE_CACHE_DIR"] = old_cache
            results = sorted(output.glob("*/*停电用户_处理结果.xlsx"))
            self.assertEqual(len(results), 2)
            self.assertTrue(original[0].is_file())
            newer = next(path for path in results if path != original[0])
            comparison = json.loads(newer.with_name(f"{newer.stem}_比较.json").read_text(encoding="utf-8"))
            self.assertTrue(comparison["available"])
            self.assertEqual(comparison["metrics"]["用户"]["left_frequent"], 1, comparison)
            self.assertEqual(comparison["metrics"]["用户"]["report_total_delta"], -1, comparison)
            self.assertEqual(comparison["statistics_columns"][11]["column"], "M")
            detail = pd.read_excel(newer, sheet_name="变化明细", dtype=str)
            self.assertIn("U2", detail["用户或馈线编码"].tolist())
            summary = newer.with_name(newer.stem.replace("停电用户_处理结果", "") + "停电摘要（简版）.txt").read_text(encoding="utf-8")
            self.assertIn("较上次成功日报统计表减少：频繁停电用户 1 户", summary)


if __name__ == "__main__":
    unittest.main()
