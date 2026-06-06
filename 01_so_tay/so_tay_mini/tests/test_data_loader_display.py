import sys
import unittest
import json
from pathlib import Path


ROOT_DIR = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT_DIR / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from data_loader import (
    expand_vocabulary_item,
    get_vocabulary_display_text,
    load_vocabulary_data,
    meanings_are_near_duplicate,
)


class VocabularyDisplayTextTest(unittest.TestCase):
    def test_parallel_vocabulary_and_meaning_are_split(self):
        item = {
            "phan_nhom": "Bài 26 - Sự kiện & Địa điểm",
            "tu_vung": "おおさかべん (大阪弁) / ほうげん (方言)",
            "y_nghia": "Tiếng Osaka, phương ngữ Osaka",
        }

        expanded = expand_vocabulary_item(item)

        self.assertEqual(
            [(row["tu_vung"], row["y_nghia"]) for row in expanded],
            [
                ("おおさかべん (大阪弁)", "Tiếng Osaka"),
                ("ほうげん (方言)", "Phương ngữ Osaka"),
            ],
        )

    def test_slash_variants_with_one_meaning_are_not_split(self):
        item = {
            "tu_vung": "さら / おさら (皿 / お皿)",
            "y_nghia": "Đĩa",
        }

        self.assertEqual(expand_vocabulary_item(item), [item])

    def test_similar_vietnamese_meanings_are_near_duplicate(self):
        self.assertTrue(
            meanings_are_near_duplicate(
                "Chán, ghét, không chấp nhận được",
                "Chán ghét, không chấp nhận được",
            )
        )
        self.assertTrue(meanings_are_near_duplicate("Chán, ghét, không chấp nhận được", "Chán, không thích"))

    def test_different_meanings_are_not_near_duplicate(self):
        self.assertFalse(meanings_are_near_duplicate("Bao nhiêu tiền", "Cho dù/thế nào"))

    def test_n4_level_suffix_keeps_n_prefix(self):
        item = {"tu_vung": "おさきに (お先に)", "cap_do": "N4"}

        self.assertEqual(get_vocabulary_display_text(item), "おさきに (お先に) - N4")

    def test_n4_vocabulary_uses_dictionary_form_outside_parentheses(self):
        data = json.loads(
            (ROOT_DIR / "data" / "vocabulary" / "n4" / "bai33_rules_signs_and_telegrams.json").read_text(
                encoding="utf-8"
            )
        )
        target = next(item for item in data if item["y_nghia"] == "Truyền đạt")

        self.assertEqual(target["tu_vung"], "つたえる (伝えます)")

    def test_n4_dictionary_conversion_keeps_masu_inside_parentheses(self):
        all_words, _json_data_map = load_vocabulary_data(ROOT_DIR / "data")
        target = next(item for item in all_words if item["tu_vung"] == "つたえる (伝えます)")

        self.assertEqual(target["tu_vung_hien_thi"], "つたえる (伝えます) - N4")

    def test_n4_outside_masu_only_remains_in_question_phrases(self):
        bad_forms = {"さがする", "だする", "おとする", "さする", "もどする"}

        for path in (ROOT_DIR / "data" / "vocabulary" / "n4").glob("*.json"):
            data = json.loads(path.read_text(encoding="utf-8"))
            for item in data:
                outside = item["tu_vung"].split("(", 1)[0].split("（", 1)[0].strip()
                with self.subTest(path=path.name, vocabulary=item["tu_vung"]):
                    self.assertNotIn(outside, bad_forms)
                    if "ます" in outside:
                        self.assertIn("ますか", outside)

    def test_long_parenthesis_variant_is_compacted(self):
        item = {
            "tu_vung": "おげんきでいらっしゃいますか (お元気でいらっしゃいますか)",
            "cap_do": "N4",
        }

        self.assertEqual(
            get_vocabulary_display_text(item),
            "おげんきでいらっしゃいますか (お元気...か) - N4",
        )

    def test_short_okurigana_after_kanji_is_kept_when_compacted(self):
        item = {
            "tu_vung": "ちょっとおねがいがあるんですが (ちょっとお願いがあるんですが)",
            "cap_do": "N4",
        }

        self.assertEqual(
            get_vocabulary_display_text(item),
            "ちょっとおねがいがあるんですが (ちょっとお願い...が) - N4",
        )

    def test_real_vocabulary_data_uses_compact_display_text(self):
        all_words, _json_data_map = load_vocabulary_data(ROOT_DIR / "data")
        target = next(
            item
            for item in all_words
            if item["tu_vung"] == "おげんきでいらっしゃいますか (お元気でいらっしゃいますか)"
        )

        self.assertEqual(
            target["tu_vung_hien_thi"],
            "おげんきでいらっしゃいますか (お元気...か) - N4",
        )

    def test_real_vocabulary_data_splits_osaka_entry(self):
        all_words, json_data_map = load_vocabulary_data(ROOT_DIR / "data")
        bai26_data = next(
            file_data
            for file_data in json_data_map
            if file_data["file_name"] == "bai26_requests_events_and_public_services.json"
        )["data"]

        self.assertNotIn(
            "おおさかべん (大阪弁) / ほうげん (方言)",
            [item["tu_vung"] for item in bai26_data],
        )
        self.assertIn(
            ("おおさかべん (大阪弁) - N4", "Tiếng Osaka"),
            [(item["tu_vung_hien_thi"], item["y_nghia"]) for item in bai26_data],
        )
        self.assertIn(
            ("ほうげん (方言) - N4", "Phương ngữ Osaka"),
            [(item["tu_vung_hien_thi"], item["y_nghia"]) for item in all_words],
        )

    def test_real_vocabulary_index_deduplicates_iya(self):
        all_words, _json_data_map = load_vocabulary_data(ROOT_DIR / "data")
        iya_rows = [item for item in all_words if item["tu_vung"] == "いや (嫌)"]

        self.assertEqual(len(iya_rows), 1)
        self.assertEqual(iya_rows[0]["y_nghia"], "Chán, ghét, không chấp nhận được")

    def test_real_vocabulary_index_keeps_different_meanings(self):
        all_words, _json_data_map = load_vocabulary_data(ROOT_DIR / "data")
        ikura_meanings = [item["y_nghia"] for item in all_words if item["tu_vung"] == "いくら"]

        self.assertEqual(sorted(ikura_meanings), ["Bao nhiêu tiền", "Cho dù/thế nào"])


if __name__ == "__main__":
    unittest.main()
