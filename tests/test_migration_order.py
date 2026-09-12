from __future__ import annotations

import sys
import unittest
from pathlib import Path


BACKEND_DIR = Path(__file__).resolve().parents[1] / "web" / "backend"
sys.path.insert(0, str(BACKEND_DIR))

from migrate_history import sort_key


class MigrationOrderTest(unittest.TestCase):
    def test_export_timestamp(self):
        key = sort_key(Path("停电用户数据_20260910175401704.xlsx"))
        self.assertEqual(key[0], 0)
        self.assertEqual(key[1].date().isoformat(), "2026-09-10")

    def test_short_date_range(self):
        key = sort_key(Path("2026.1.1-6.15停电用户.xlsx"))
        self.assertEqual(key[0], 0)
        self.assertEqual(key[1].date().isoformat(), "2026-06-15")

    def test_full_date_range(self):
        key = sort_key(Path("2026.1.1-2026.9.1.xlsx"))
        self.assertEqual(key[1].date().isoformat(), "2026-09-01")

    def test_compact_date_range(self):
        key = sort_key(Path("20260601-0706停电数据.xlsx"))
        self.assertEqual(key[1].date().isoformat(), "2026-07-06")

    def test_unknown_name_requires_confirmation(self):
        self.assertEqual(sort_key(Path("工作簿2.xlsx"))[0], 1)


if __name__ == "__main__":
    unittest.main()
