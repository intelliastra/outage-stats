from __future__ import annotations

import sys
import unittest
from datetime import datetime, timezone
from pathlib import Path


BACKEND_DIR = Path(__file__).resolve().parents[1] / "web" / "backend"
sys.path.insert(0, str(BACKEND_DIR))

import database


class DatabaseKeyTest(unittest.TestCase):
    def setUp(self):
        self.started = datetime(2026, 9, 10, 8, 30, tzinfo=timezone.utc)

    def test_record_id_has_highest_priority(self):
        key, source = database._record_key(
            {"停电记录id": "R-1", "事件id": "E-1", "用户id": "U-1"}, self.started
        )
        self.assertEqual(key, "record:R-1")
        self.assertEqual(source, "停电记录id")

    def test_event_and_user_are_second_priority(self):
        key, source = database._record_key({"事件id": "E-1", "用户id": "U-1"}, self.started)
        self.assertEqual(key, "event-user:E-1|U-1")
        self.assertEqual(source, "事件id+用户id")

    def test_fallback_key_is_deterministic(self):
        payload = {"工单号": "W1", "用户编码": "001", "所属馈线编码": "F1"}
        first = database._record_key(payload, self.started)
        second = database._record_key(payload, self.started)
        self.assertEqual(first, second)
        self.assertTrue(first[0].startswith("fallback:W1|001|F1|"))


if __name__ == "__main__":
    unittest.main()

