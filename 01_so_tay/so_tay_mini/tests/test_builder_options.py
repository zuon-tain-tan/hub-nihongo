import sys
import unittest
from pathlib import Path
from unittest.mock import Mock, patch


ROOT_DIR = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT_DIR / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

import builder
from builder import ask_export_options, parse_export_options


class ExportOptionParsingTest(unittest.TestCase):
    def test_all_inputs_return_default_order(self):
        for raw in ["", "all", "a", "tat ca", "tất cả", "  ALL  "]:
            with self.subTest(raw=raw):
                self.assertEqual(parse_export_options(raw), ["1", "2", "3", "4", "5", "6", "7"])

    def test_number_lists_keep_input_order(self):
        cases = {
            "1": ["1"],
            "1,2,3,4,5,6": ["1", "2", "3", "4", "5", "6"],
            "1 2 3 4 5 6": ["1", "2", "3", "4", "5", "6"],
            "2,4,3,5": ["2", "4", "3", "5"],
            "1, 2; 3": ["1", "2", "3"],
        }
        for raw, expected in cases.items():
            with self.subTest(raw=raw):
                self.assertEqual(parse_export_options(raw), expected)

    def test_duplicate_numbers_are_removed_without_reordering(self):
        self.assertEqual(parse_export_options("2,4,2,3,4,5"), ["2", "4", "3", "5"])

    def test_full_width_input_is_accepted(self):
        self.assertEqual(parse_export_options("１，２，３"), ["1", "2", "3"])
        self.assertEqual(parse_export_options("1、2、3"), ["1", "2", "3"])

    def test_invalid_inputs_are_rejected(self):
        for raw in ["0", "8", "1,8", "1-3", "abc", "1.2"]:
            with self.subTest(raw=raw):
                self.assertIsNone(parse_export_options(raw))

    def test_prompt_retries_after_invalid_input(self):
        with patch("builtins.input", side_effect=["1,8", "1,2,3"]):
            with patch("builtins.print"):
                self.assertEqual(ask_export_options(), ["1", "2", "3"])

    def test_build_document_uses_only_selected_sections(self):
        doc = Mock()
        added_sections = []

        def record_section(_doc, add_func, *_args):
            added_sections.append(add_func.__name__)

        with patch("builder.ask_export_options", return_value=["1", "2", "3", "4", "5", "6"]):
            with patch("builder.setup_document", return_value=doc):
                with patch("builder.load_vocabulary_data", return_value=([], {})):
                    with patch("builder.load_5s_data", return_value={}):
                        with patch(
                            "builder.load_special_topics",
                            return_value={"horenso": {}, "aisatsu": {}, "garbage_sorting": {}},
                        ):
                            with patch("builder.add_toc_section"):
                                with patch("builder.add_section_with_break", side_effect=record_section):
                                    with patch("builder.prepare_cover_pdf", return_value=Path("missing-cover.pdf")):
                                        with patch("builder.update_docx_fields_and_export_pdf"):
                                            with patch("builtins.print"):
                                                builder.build_document()

        self.assertEqual(
            added_sections,
            [
                "add_5s_section",
                "add_horenso_section",
                "add_aisatsu_section",
                "add_garbage_section",
                "add_japanese_index",
                "add_vietnamese_index",
                "add_online_version_page",
                "add_completion_image_page",
            ],
        )
        self.assertNotIn("add_grouped_vocabulary", added_sections)


if __name__ == "__main__":
    unittest.main()
