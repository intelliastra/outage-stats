from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


class UploadDestinationTest(unittest.TestCase):
    def test_exclude_page_targets_its_own_newdata_directory(self):
        source = (ROOT / "web" / "frontend" / "js" / "exclude.js").read_text(encoding="utf-8")
        self.assertIn('/api/upload?dest=exclude_newdata', source)

    def test_chunk_helper_preserves_explicit_destination(self):
        source = (ROOT / "web" / "frontend" / "js" / "common.js").read_text(encoding="utf-8")
        self.assertIn('match(/[?&]dest=([^&#]+)/)', source)
        self.assertIn('decodeURIComponent(destination[1])', source)

    def test_external_damage_script_reads_all_sheets(self):
        source = (
            ROOT / "统计材料" / "运行脚本" / "exclude_external_damage_v1.0.0.py"
        ).read_text(encoding="utf-8")
        self.assertIn("df = read_excel_all_sheets(input_path)", source)
        self.assertNotIn("df = pd.read_excel(input_path, sheet_name=0)", source)


if __name__ == "__main__":
    unittest.main()
