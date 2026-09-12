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


if __name__ == "__main__":
    unittest.main()
