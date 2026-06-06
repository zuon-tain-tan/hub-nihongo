import sys
import unittest
from pathlib import Path


ROOT_DIR = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT_DIR / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from formatting import setup_document
from sections import (
    INDEX_EMPTY_RIGHT_PLACEHOLDER,
    add_index_paragraph,
    get_vietnamese_index_meaning_text,
    get_vietnamese_index_vocabulary_text,
    split_index_meaning_if_long,
)


class IndexParagraphTest(unittest.TestCase):
    def test_long_meaning_without_comma_moves_to_leader_continuation(self):
        left_text = "おさきにどうぞ (お先にどうぞ) - N4"
        right_text = "Mời anh/chị cứ về trước"

        self.assertEqual(split_index_meaning_if_long(left_text, right_text), ("", right_text))

    def test_index_paragraph_uses_tab_leader_on_continuation_line(self):
        doc = setup_document()
        left_text = "おさきにどうぞ (お先にどうぞ) - N4"
        right_text = "Mời anh/chị cứ về trước"

        add_index_paragraph(doc, left_text, right_text, split_long_meaning=True)

        self.assertEqual(doc.paragraphs[-2].text, f"{left_text}\t{INDEX_EMPTY_RIGHT_PLACEHOLDER}")
        self.assertEqual(doc.paragraphs[-1].text, f"\t{right_text}")
        self.assertTrue(doc.paragraphs[-2].paragraph_format.keep_with_next)

    def test_long_meaning_with_comma_moves_whole_meaning_to_continuation(self):
        left_text = "どうもおつかれさまでした (どうもお疲れ...た) - N4"
        right_text = "Xong rồi ạ, cảm ơn quý khách"

        self.assertEqual(
            split_index_meaning_if_long(left_text, right_text),
            ("", right_text),
        )

    def test_short_meaning_moves_to_continuation_when_left_side_is_too_long(self):
        left_text = "どちらさまでしょうか (どちら様でしょうか) - N4"
        right_text = "Ai đấy ạ?"

        self.assertEqual(split_index_meaning_if_long(left_text, right_text), ("", right_text))

    def test_medium_meaning_stays_with_very_long_left_side(self):
        left_text = "よろしくおつたえください (よろしくお伝えください) - N4"
        right_text = "Cho tôi gửi lời hỏi thăm"

        self.assertEqual(
            split_index_meaning_if_long(left_text, right_text, keep_medium_right_with_long_left=True),
            (right_text, ""),
        )

    def test_medium_meaning_moves_to_continuation_when_left_side_is_not_long_enough(self):
        left_text = "おさきにどうぞ (お先にどうぞ) - N4"
        right_text = "Mời anh/chị cứ về trước"

        self.assertEqual(split_index_meaning_if_long(left_text, right_text), ("", right_text))

    def test_vietnamese_index_moves_long_meaning_before_japanese_to_continuation(self):
        doc = setup_document()
        left_text = "Akihabara, khu bán đồ điện tử nổi tiếng ở Tokyo - N4"
        right_text = "あきはばら (秋葉原)"

        add_index_paragraph(doc, left_text, right_text, split_long_meaning=True)

        self.assertEqual(doc.paragraphs[-2].text, f"{left_text}\t{INDEX_EMPTY_RIGHT_PLACEHOLDER}")
        self.assertEqual(doc.paragraphs[-1].text, f"\t{right_text}")

    def test_vietnamese_index_moves_level_to_meaning_side(self):
        item = {
            "tu_vung": "あきはばら (秋葉原)",
            "tu_vung_hien_thi": "あきはばら (秋葉原) - N4",
            "y_nghia": "Akihabara, khu bán đồ điện tử nổi tiếng ở Tokyo",
            "cap_do": "N4",
        }

        self.assertEqual(
            get_vietnamese_index_meaning_text(item),
            "Akihabara, khu bán đồ điện tử nổi tiếng ở Tokyo - N4",
        )
        self.assertEqual(get_vietnamese_index_vocabulary_text(item), "あきはばら (秋葉原)")

    def test_vietnamese_index_does_not_show_n5_level(self):
        item = {
            "tu_vung": "せんせい (先生)",
            "tu_vung_hien_thi": "せんせい (先生)",
            "y_nghia": "Giáo viên",
            "cap_do": "N5",
        }

        self.assertEqual(get_vietnamese_index_meaning_text(item), "Giáo viên")
        self.assertEqual(get_vietnamese_index_vocabulary_text(item), "せんせい (先生)")


if __name__ == "__main__":
    unittest.main()
