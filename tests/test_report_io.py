from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

import openpyxl
import pandas as pd


SCRIPT_DIR = Path(__file__).resolve().parents[1] / "统计材料" / "运行脚本"
sys.path.insert(0, str(SCRIPT_DIR))

from report_io import write_tables_streaming, write_text_atomic


class ReportIoTest(unittest.TestCase):
    def test_atomic_text_writer(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "摘要.txt"
            write_text_atomic(target, "第一行\n第二行")
            self.assertEqual(target.read_text(encoding="utf-8"), "第一行\n第二行\n")
            self.assertFalse(list(target.parent.glob("*.tmp")))

    def test_streaming_writer_preserves_contract(self):
        tables = {
            "用户停电总次数统计表": pd.DataFrame(
                {
                    "用户编码": ["00123", 456.0],
                    "停电总次数": [2.0, "3"],
                    "停电开始时间": pd.to_datetime(["2026-09-01 08:30", "2026-09-02 09:45"]),
                }
            ),
            "停电预警线路清单": pd.DataFrame({"所属馈线编码": ["0007"]}),
        }
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "report.xlsx"
            write_tables_streaming(tables, target)
            self.assertTrue(target.is_file())
            self.assertFalse(list(target.parent.glob("*.tmp")))

            workbook = openpyxl.load_workbook(target, read_only=False, data_only=True)
            self.assertEqual(workbook.sheetnames, list(tables))
            sheet = workbook["用户停电总次数统计表"]
            self.assertEqual(sheet.freeze_panes, "A2")
            self.assertEqual(sheet.auto_filter.ref, "A1:C3")
            self.assertEqual(sheet["A2"].value, "00123")
            self.assertEqual(sheet["A2"].number_format, "@")
            self.assertEqual(sheet["B3"].value, 3)
            self.assertEqual(sheet["B3"].number_format, "0")
            workbook.close()


if __name__ == "__main__":
    unittest.main()
