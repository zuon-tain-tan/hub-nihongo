import sys
import unittest
from pathlib import Path


ROOT_DIR = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT_DIR / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from image_vocabulary_builder import (
    IMAGE_VOCABULARY_FULL_DOCX_NAME_TEMPLATE,
    IMAGE_VOCABULARY_FULL_PDF_NAME_TEMPLATE,
    IMAGE_VOCABULARY_DOCX_NAME_TEMPLATE,
    IMAGE_VOCABULARY_PDF_NAME_TEMPLATE,
    create_image_vocabulary_full_document,
    create_image_vocabulary_document,
    enrich_image_vocabulary_metadata,
    format_lesson_selection_slug,
    get_vocabulary_reading,
    split_langoal_vocabulary_terms,
    parse_lesson_number,
    parse_lesson_selection,
)
from image_vocabulary_scraper import (
    format_lesson_dir_name,
    get_langoal_lesson_url,
    parse_langoal_vocabulary_items,
)


SAMPLE_HTML = """
<main>
  <h2>新出語彙導入イラスト</h2>
  <div class="w-layout-grid category-grid">
    <div>
      <img src="https://cdn.langoal.com/images/mnn-l1-1.png"
           srcset="https://cdn.langoal.com/images/mnn-l1-1-p-500.png 500w, https://cdn.langoal.com/images/mnn-l1-1.png 512w"
           alt="「教師」「先生」">
      <h3 class="image-sample-title">「教師」「先生」</h3>
    </div>
    <div>
      <img src="https://cdn.langoal.com/images/mnn-l1-2-p-500.png"
           srcset="https://cdn.langoal.com/images/mnn-l1-2-p-500.png 500w, https://cdn.langoal.com/images/mnn-l1-2.png 512w"
           alt="「学生」">
    </div>
    <div>
      <img src="https://cdn.langoal.com/images/mnn-l1-sheet.png" alt="まとめ">
    </div>
  </div>
  <h2>新出語彙シート</h2>
  <div class="w-layout-grid category-grid">
    <img src="https://cdn.langoal.com/images/mnn-l1-sheet-1.png" alt="sheet">
  </div>
</main>
"""


class ImageVocabularyTest(unittest.TestCase):
    def test_langoal_parser_reads_intro_vocabulary_images_only(self):
        items = parse_langoal_vocabulary_items(SAMPLE_HTML, lesson=1)

        self.assertEqual(
            items,
            [
                {
                    "index": 1,
                    "vocabulary": "「教師」「先生」",
                    "image_url": "https://cdn.langoal.com/images/mnn-l1-1.png",
                },
                {
                    "index": 2,
                    "vocabulary": "「学生」",
                    "image_url": "https://cdn.langoal.com/images/mnn-l1-2.png",
                },
            ],
        )

    def test_lesson_helpers_format_paths_and_urls(self):
        self.assertEqual(format_lesson_dir_name(1), "lesson_01")
        self.assertEqual(get_langoal_lesson_url(12), "https://langoal.com/vocbs/lesson-12")
        self.assertEqual(parse_lesson_number("1"), 1)
        self.assertEqual(parse_lesson_number(" 12 "), 12)
        self.assertIsNone(parse_lesson_number("0"))
        self.assertIsNone(parse_lesson_number("abc"))

    def test_lesson_selection_accepts_ranges_and_lists(self):
        self.assertEqual(parse_lesson_selection("1"), [1])
        self.assertEqual(parse_lesson_selection("1-3"), [1, 2, 3])
        self.assertEqual(parse_lesson_selection("1,3,5"), [1, 3, 5])
        self.assertEqual(parse_lesson_selection("1 2 3"), [1, 2, 3])
        self.assertEqual(parse_lesson_selection("1 - 3"), [1, 2, 3])
        self.assertEqual(parse_lesson_selection("１-３"), [1, 2, 3])
        self.assertEqual(parse_lesson_selection("1-3,3,5"), [1, 2, 3, 5])
        self.assertEqual(parse_lesson_selection("1-50")[0], 1)
        self.assertEqual(parse_lesson_selection("1-50")[-1], 50)
        self.assertEqual(len(parse_lesson_selection("1-50")), 50)

    def test_lesson_selection_rejects_invalid_ranges(self):
        for raw in ["0", "51", "1-51", "0-1", "3-1", "abc", "1--3", ""]:
            with self.subTest(raw=raw):
                self.assertIsNone(parse_lesson_selection(raw))

    def test_document_uses_a4_and_grid_table(self):
        metadata = {
            "lesson": 1,
            "source_url": "https://langoal.com/vocbs/lesson-1",
            "items": [
                {
                    "index": 1,
                    "vocabulary": "「教師」「先生」",
                    "reading": "きょうし / せんせい",
                    "image_file": "assets/images/qr/so_tay_a5_online_qr.png",
                }
            ],
        }

        doc = create_image_vocabulary_document(metadata, project_root=ROOT_DIR)
        section = doc.sections[0]

        self.assertAlmostEqual(section.page_width.cm, 21.0, places=1)
        self.assertAlmostEqual(section.page_height.cm, 29.7, places=1)
        self.assertEqual(len(doc.tables), 1)
        self.assertEqual(len(doc.inline_shapes), 1)
        self.assertIn("きょうし / せんせい", [paragraph.text for paragraph in doc.tables[0].cell(0, 0).paragraphs])

    def test_vocabulary_reading_is_added_from_existing_vocab_index(self):
        reading_index = {
            "教師": "きょうし",
            "先生": "せんせい",
            "エンジニア": "エンジニア",
        }
        metadata = {
            "lesson": 1,
            "items": [
                {"index": 1, "vocabulary": "「教師」「先生」"},
                {"index": 2, "vocabulary": "「エンジニア」"},
            ],
        }

        changed = enrich_image_vocabulary_metadata(metadata, reading_index)

        self.assertTrue(changed)
        self.assertEqual(metadata["items"][0]["reading"], "きょうし / せんせい")
        self.assertEqual(metadata["items"][1]["reading"], "エンジニア")

    def test_vocabulary_reading_uses_kana_when_word_is_already_kana(self):
        self.assertEqual(get_vocabulary_reading("「えんぴつ」", {}), "えんぴつ")
        self.assertEqual(get_vocabulary_reading("「カード」", {}), "カード")

    def test_langoal_terms_drop_notes_and_optional_prefixes(self):
        self.assertEqual(split_langoal_vocabulary_terms("「吸います［たばこを］」"), ["吸います"])
        self.assertEqual(split_langoal_vocabulary_terms("「（お）国」"), ["(お)国", "国", "お国"])

    def test_vocabulary_reading_matches_terms_with_notes(self):
        reading_index = {
            "吸います": "すう",
            "国": "くに",
            "号室": "ごうしつ",
            "CD": "シーディー",
        }

        self.assertEqual(get_vocabulary_reading("「吸います［たばこを］」", reading_index), "すう")
        self.assertEqual(get_vocabulary_reading("「（お）国」", reading_index), "くに")
        self.assertEqual(get_vocabulary_reading("「ー号室」", reading_index), "ごうしつ")
        self.assertEqual(get_vocabulary_reading("「CD」", reading_index), "シーディー")

    def test_full_document_adds_each_lesson(self):
        metadata_list = [
            {
                "lesson": 1,
                "source_url": "https://langoal.com/vocbs/lesson-1",
                "items": [
                    {
                        "index": 1,
                        "vocabulary": "「教師」「先生」",
                        "image_file": "assets/images/qr/so_tay_a5_online_qr.png",
                    }
                ],
            },
            {
                "lesson": 2,
                "source_url": "https://langoal.com/vocbs/lesson-2",
                "items": [
                    {
                        "index": 1,
                        "vocabulary": "「これ」",
                        "image_file": "assets/images/qr/so_tay_a5_online_qr.png",
                    }
                ],
            },
        ]

        doc = create_image_vocabulary_full_document(metadata_list, project_root=ROOT_DIR)
        paragraph_texts = [paragraph.text for paragraph in doc.paragraphs]

        self.assertIn("HÌNH ẢNH TỪ VỰNG - BÀI 1", paragraph_texts)
        self.assertIn("HÌNH ẢNH TỪ VỰNG - BÀI 2", paragraph_texts)
        self.assertEqual(len(doc.tables), 2)
        self.assertEqual(len(doc.inline_shapes), 2)

    def test_output_names_include_lesson_number(self):
        self.assertEqual(
            IMAGE_VOCABULARY_DOCX_NAME_TEMPLATE.format(lesson=1),
            "hinh_anh_tu_vung_lesson_01.docx",
        )
        self.assertEqual(
            IMAGE_VOCABULARY_PDF_NAME_TEMPLATE.format(lesson=1),
            "hinh_anh_tu_vung_lesson_01.pdf",
        )

    def test_full_output_names_include_selection(self):
        self.assertEqual(format_lesson_selection_slug([1]), "lesson_01")
        self.assertEqual(format_lesson_selection_slug([1, 2, 3]), "lessons_01_03")
        self.assertEqual(format_lesson_selection_slug([1, 3, 5]), "lessons_01_03_05")
        self.assertEqual(
            IMAGE_VOCABULARY_FULL_DOCX_NAME_TEMPLATE.format(selection="lessons_01_50"),
            "hinh_anh_tu_vung_lessons_01_50.docx",
        )
        self.assertEqual(
            IMAGE_VOCABULARY_FULL_PDF_NAME_TEMPLATE.format(selection="lessons_01_50"),
            "hinh_anh_tu_vung_lessons_01_50.pdf",
        )


if __name__ == "__main__":
    unittest.main()
