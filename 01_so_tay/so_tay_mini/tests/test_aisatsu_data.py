import json
import unittest
from pathlib import Path


ROOT_DIR = Path(__file__).resolve().parents[1]


class AisatsuDataTest(unittest.TestCase):
    def test_returning_home_section_starts_on_new_page(self):
        data = json.loads((ROOT_DIR / "data" / "topics" / "aisatsu.json").read_text(encoding="utf-8"))
        section = next(item for item in data if item["tieu_de_tieng_viet"] == "6. Trước khi về nước")

        self.assertTrue(section["ngat_trang_truoc"])


if __name__ == "__main__":
    unittest.main()
