from __future__ import annotations

import sys
import importlib.util
import tempfile
import unittest
from pathlib import Path

import pandas as pd
import openpyxl

SCRIPT_DIR = Path(__file__).resolve().parents[1] / "统计材料" / "运行脚本"
sys.path.insert(0, str(SCRIPT_DIR))

from external_damage_filter import exclude_external_damage_rows
from report_comparison import compare_reports
from report_io import write_tables_streaming
from statistics_priority import classify_entity_groups, official_section


def load_script(name: str, file_name: str):
    spec = importlib.util.spec_from_file_location(name, SCRIPT_DIR / file_name)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class StatisticsUpdatesTest(unittest.TestCase):
    def test_warning_and_frequent_are_exclusive(self):
        frequent = pd.DataFrame([
            {"用户编码": "U1", "所属地市": "广州", "频繁停电类型": "一年内停电次数超过5次；连续60天停电次数超过3次"},
            {"用户编码": "U1", "所属地市": "广州", "频繁停电类型": "连续60天停电次数超过3次"},
            {"用户编码": "U3", "所属地市": "广州", "频繁停电类型": "一年内预安排停电次数超过3次"},
        ])
        warning = pd.DataFrame([
            {"用户编码": "U1", "所属地市": "广州", "停电预警类型": "近30天停电2次"},
            {"用户编码": "U2", "所属地市": "广州", "停电预警类型": "一年内停电4-5次；近50天停电3次；近30天停电2次"},
        ])
        groups = classify_entity_groups(frequent, warning, key_column="用户编码")
        self.assertEqual(groups["frequent:年>5次"], {"广州": 1})
        self.assertEqual(groups["frequent:预安排>3次"], {"广州": 1})
        self.assertEqual(groups["warning:30天2次"], {"广州": 1})
        self.assertNotIn("warning:一年4-5次", groups)
        warning_section = official_section(groups, "warning", ("一年4-5次", "50天3次", "30天2次"))
        self.assertEqual([part.get("广州", 0) for part in warning_section], [0, 0, 1, 1])

    def test_feeder_deduplicates_and_nea_filters(self):
        frequent = pd.DataFrame([
            {"所属馈线编码": "L1", "所属地市": "东莞", "频繁停电类型": "连续60天停电次数超过3次", "是否纳入国家能源局台账": "是"},
            {"所属馈线编码": "L1", "所属地市": "东莞", "频繁停电类型": "一年内停电次数超过5次", "是否纳入国家能源局台账": "是"},
        ])
        groups = classify_entity_groups(frequent, pd.DataFrame(), key_column="所属馈线编码", nea_only=True)
        self.assertEqual(groups["frequent:年>5次"], {"东莞": 1})
        self.assertNotIn("frequent:60天>3次", groups)

    def test_both_scripts_populate_2026_and_2025_template_columns(self):
        current = {
            "频繁停电用户清单": pd.DataFrame([
                {"用户编码": "U1", "所属地市": "广州", "频繁停电类型": "一年内停电次数超过5次；连续60天停电次数超过3次"},
            ]),
            "停电预警用户清单": pd.DataFrame([
                {"用户编码": "U1", "所属地市": "广州", "停电预警类型": "近30天停电2次"},
                {"用户编码": "U2", "所属地市": "广州", "停电预警类型": "一年内停电4-5次；近30天停电2次"},
            ]),
            "频繁停电线路清单": pd.DataFrame([
                {"所属馈线编码": "L1", "所属地市": "广州", "频繁停电类型": "一年内停电次数超过5次；连续60天停电次数超过3次", "是否纳入国家能源局台账": "是"},
            ]),
            "停电预警线路清单": pd.DataFrame([
                {"所属馈线编码": "L1", "所属地市": "广州", "停电预警类型": "近30天停电2次"},
                {"所属馈线编码": "L2", "所属地市": "广州", "停电预警类型": "一年内停电4-5次；近50天停电3次"},
            ]),
        }
        older = {key: current[key] for key in ("频繁停电用户清单", "频繁停电线路清单")}
        for name, script in (
            ("daily_script_test", "process_outage_tables_v1.0.8.py"),
            ("external_script_test", "exclude_external_damage_v1.0.0.py"),
        ):
            module = load_script(name, script)
            values = module.compute_2026_stats_dicts(current)
            self.assertEqual(len(values), 20)
            self.assertEqual(values[3].get("广州"), 1)   # E: only U2 warning
            self.assertEqual(values[7].get("广州"), 1)   # I: only L2 warning
            self.assertEqual(values[11].get("广州"), 1)  # M: U1 once
            self.assertEqual(values[15].get("广州"), 1)  # Q: L1 once
            self.assertEqual(values[19].get("广州"), 1)  # U: NEA L1 once
            prior_values = module.compute_2025_stats_dicts(older)
            self.assertEqual(len(prior_values), 12)
            self.assertEqual(prior_values[3].get("广州"), 1)

    def test_external_damage_filters_prefixes_and_checks_ae(self):
        columns = [f"列{i}" for i in range(31)]
        columns[30] = "责任原因代码"
        raw = pd.DataFrame([
            [""] * 30 + [value] for value in ("5132", "5142", "5150", "5160", None, 5139)
        ], columns=columns)
        result, counts = exclude_external_damage_rows(raw)
        self.assertEqual(counts["removed_total"], 4)
        self.assertEqual(counts["remaining_rows"], 2)
        self.assertEqual(counts["blank_code_rows"], 1)
        self.assertEqual(result["责任原因代码"].iloc[0], "5160")
        with self.assertRaisesRegex(ValueError, "AE列"):
            exclude_external_damage_rows(raw.rename(columns={"责任原因代码": "错误列"}))

    def test_external_script_rejects_an_invalid_second_sheet(self):
        columns = [f"列{i}" for i in range(31)]
        columns[30] = "责任原因代码"
        good = pd.DataFrame([[""] * 30 + ["5160"]], columns=columns)
        bad = good.rename(columns={"责任原因代码": "错误列"})
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "two_sheets.xlsx"
            write_tables_streaming({"good": good, "bad": bad}, source)
            module = load_script("external_all_sheets_test", "exclude_external_damage_v1.0.0.py")
            with self.assertRaisesRegex(RuntimeError, "AE列"):
                module.read_excel_all_sheets(source)

    def test_exclude_endpoint_rejects_a_different_daily_file(self):
        backend_dir = Path(__file__).resolve().parents[1] / "web" / "backend"
        sys.path.insert(0, str(backend_dir))
        from upload_match import latest_matching_daily_file
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            daily, exclude = root / "daily", root / "exclude"
            daily.mkdir(); exclude.mkdir()
            (daily / "same.xlsx").write_bytes(b"daily original")
            (exclude / "same.xlsx").write_bytes(b"manually changed")
            with self.assertRaisesRegex(ValueError, "内容不一致"):
                latest_matching_daily_file(exclude / "same.xlsx", daily)
            (exclude / "same.xlsx").write_bytes(b"daily original")
            self.assertEqual(latest_matching_daily_file(exclude / "same.xlsx", daily).name, "same.xlsx")

    def test_comparison_keeps_disappeared_entity(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            old_path = root / "2026.1.1-9.10停电用户_处理结果.xlsx"
            old_tables = {
                "用户停电总次数统计表": pd.DataFrame([{"用户编码": "U1", "停电开始时间": "2026-09-10"}]),
                "频繁停电用户清单": pd.DataFrame([
                    {"用户编码": "U1", "所属馈线编码": "L1", "所属地市": "广州", "停电总次数": 6, "频繁停电类型": "一年内停电次数超过5次；连续60天停电次数超过3次"},
                    {"用户编码": "U2", "所属馈线编码": "L2", "所属地市": "广州", "停电总次数": 4, "频繁停电类型": "连续60天停电次数超过3次"},
                ]),
                "停电预警用户清单": pd.DataFrame(),
                "频繁停电线路清单": pd.DataFrame([
                    {"所属馈线编码": "L2", "所属地市": "广州", "停电总次数": 4, "频繁停电类型": "连续60天停电次数超过3次"},
                ]),
                "停电预警线路清单": pd.DataFrame(),
            }
            write_tables_streaming(old_tables, old_path)
            (root / "2026.1.1-9.10停电用户_处理结果_发出版.xlsx").write_bytes(old_path.read_bytes())
            old_stats = openpyxl.Workbook()
            old_stats.active.title = "2026年统计表"
            old_stats.active["A3"] = "公司"
            old_stats.active["M3"] = 4  # Old report display was overcounted.
            old_stats.save(root / "2026.1.1-9.10统计表.xlsx")
            for suffix in ("停电摘要（简版）.txt", "停电摘要（全量版）.txt"):
                (root / f"2026.1.1-9.10{suffix}").write_text("完成", encoding="utf-8")
            current = {
                "频繁停电用户清单": pd.DataFrame([
                    {"用户编码": "U1", "所属馈线编码": "L1", "所属地市": "广州", "停电总次数": 5, "频繁停电类型": "一年内停电次数超过5次"},
                ]),
                "停电预警用户清单": pd.DataFrame(),
                "频繁停电线路清单": pd.DataFrame(),
                "停电预警线路清单": pd.DataFrame(),
            }
            old_raw = pd.DataFrame([{"工单号": "A2", "用户编码": "U2", "所属馈线编码": "L2", "停电开始时间": "2026-09-09"}])
            new_raw = pd.DataFrame([{"工单号": "A1", "用户编码": "U1", "所属馈线编码": "L1", "停电开始时间": "2026-09-10"}])
            new_sections = [{} for _ in range(20)]
            new_sections[11] = {"广州": 1}
            info, detail, explanations = compare_reports(
                current, old_path, new_raw, old_raw, data_year=2026,
                current_source="fixture.xlsx", current_stats_dicts=new_sections,
            )
            self.assertTrue(info["available"])
            self.assertEqual(info["metrics"]["用户"]["left_frequent"], 1)
            self.assertEqual(info["metrics"]["用户"]["report_total_delta"], -3)
            self.assertEqual(info["metrics"]["用户"]["change_detail_count"], 2)
            self.assertEqual(info["metrics"]["用户"]["decreased_outage_count"], 2)
            self.assertEqual(info["metrics"]["用户"]["same_count_type_changed"], 0)
            lost = detail[(detail["对象类型"] == "用户") & (detail["用户或馈线编码"] == "U2")].iloc[0]
            self.assertIn("本次减少", lost["变化说明"])
            self.assertIn("工单A2", lost["上次来源定位"])
            self.assertIn("停电次数减少", explanations["U1"])
            self.assertIn("多规则命中", explanations["U1"])
            self.assertIn("频繁/年>5次；频繁/60天>3次", detail.loc[detail["用户或馈线编码"] == "U1", "上次分类"].iloc[0])
            self.assertIn("类型变化", detail.loc[detail["用户或馈线编码"] == "U1", "变化说明"].iloc[0])

    def test_comparison_only_reduced_or_same_count_type_changed(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            stem = "2026.1.1-9.10"
            previous = root / f"{stem}停电用户_处理结果.xlsx"
            old = {
                "用户停电总次数统计表": pd.DataFrame([{"停电开始时间": "2026-09-10"}]),
                "频繁停电用户清单": pd.DataFrame([
                    {"用户编码": "TYPE", "所属地市": "广州", "停电总次数": 6, "频繁停电类型": "一年内停电次数超过5次；连续60天停电次数超过3次"},
                    {"用户编码": "LESS", "所属地市": "广州", "停电总次数": 6, "频繁停电类型": "一年内停电次数超过5次"},
                    {"用户编码": "MORE", "所属地市": "广州", "停电总次数": 6, "频繁停电类型": "一年内停电次数超过5次"},
                ]),
                "停电预警用户清单": pd.DataFrame(),
                "频繁停电线路清单": pd.DataFrame(),
                "停电预警线路清单": pd.DataFrame(),
            }
            write_tables_streaming(old, previous)
            (root / f"{stem}停电用户_处理结果_发出版.xlsx").write_bytes(previous.read_bytes())
            stats = openpyxl.Workbook()
            stats.active.title = "2026年统计表"
            stats.save(root / f"{stem}统计表.xlsx")
            for suffix in ("停电摘要（简版）.txt", "停电摘要（全量版）.txt"):
                (root / f"{stem}{suffix}").write_text("完成", encoding="utf-8")
            now = {
                "频繁停电用户清单": pd.DataFrame([
                    {"用户编码": "TYPE", "所属地市": "广州", "停电总次数": 6, "频繁停电类型": "一年内停电次数超过5次"},
                    {"用户编码": "LESS", "所属地市": "广州", "停电总次数": 5, "频繁停电类型": "一年内停电次数超过5次；连续60天停电次数超过3次"},
                    {"用户编码": "MORE", "所属地市": "广州", "停电总次数": 7, "频繁停电类型": "一年内停电次数超过5次"},
                    {"用户编码": "NEW", "所属地市": "广州", "停电总次数": 6, "频繁停电类型": "一年内停电次数超过5次"},
                ]),
                "停电预警用户清单": pd.DataFrame(),
                "频繁停电线路清单": pd.DataFrame(),
                "停电预警线路清单": pd.DataFrame(),
            }
            info, detail, _ = compare_reports(now, previous, pd.DataFrame(), pd.DataFrame(),
                                              data_year=2026, current_source="new.xlsx")
            self.assertEqual(set(detail["用户或馈线编码"]), {"TYPE", "LESS"})
            changed_type = detail.set_index("用户或馈线编码").loc["TYPE"]
            self.assertEqual(changed_type["上次停电次数"], changed_type["本次停电次数"])
            self.assertIn("频繁/60天>3次", changed_type["上次分类"])
            self.assertNotIn("频繁/60天>3次", changed_type["本次分类"])
            self.assertIn("类型变化", changed_type["变化说明"])
            decreased = detail.set_index("用户或馈线编码").loc["LESS"]
            self.assertIn("停电次数减少", decreased["变化说明"])
            self.assertIn("类型变化", decreased["变化说明"])
            self.assertIn("原因待核", decreased["备注"])
            self.assertEqual(info["changed_entities"], 2)
            self.assertEqual(info["metrics"]["用户"]["change_detail_count"], 2)
            self.assertEqual(info["metrics"]["用户"]["decreased_outage_count"], 1)
            self.assertEqual(info["metrics"]["用户"]["same_count_type_changed"], 1)


if __name__ == "__main__":
    unittest.main()
