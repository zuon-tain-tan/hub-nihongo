import re
import sys
from datetime import datetime

from config import ASSETS_DIR, DATA_DIR, OUTPUT_DIR
from data_loader import load_5s_data, load_special_topics, load_vocabulary_data
from export import merge_pdfs, prepare_cover_pdf, update_docx_fields_and_export_pdf
from formatting import setup_document
from sections import (
    add_5s_section,
    add_aisatsu_section,
    add_garbage_section,
    add_grouped_vocabulary,
    add_horenso_section,
    add_japanese_index,
    add_online_version_page,
    add_section_with_break,
    add_toc_section,
    add_vietnamese_index,
)


def configure_stdio():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stdin, "reconfigure"):
        sys.stdin.reconfigure(encoding="utf-8")


def ask_export_options():
    valid_options = {"1", "2", "3", "4", "5", "6", "7"}
    all_options = ["1", "2", "3", "4", "5", "6", "7"]
    print("Chọn phần muốn tạo trong sổ tay:")
    print("1. 5S-K-C")
    print("2. HORENSO")
    print("3. Aisatsu")
    print("4. Phân loại rác")
    print("5. A-z chữ cái Hira,Kata")
    print("6. A-z chữ cái Vie")
    print("7. Gom cụm từ vựng")
    print("Nhấn Enter hoặc nhập all để tạo toàn bộ theo thứ tự mặc định.")
    print("Ví dụ nhập: 1 hoặc 1,3 hoặc 1 2 3 hoặc 2,4,3,5")
    print(
        "Lưu ý: Thứ tự nhập sẽ là thứ tự xuất hiện trong tài liệu. "
        "Ví dụ 2,4,3,5 sẽ xuất HORENSO trước, rồi Phân loại rác, "
        "Aisatsu, rồi TRA NHANH Hira/Kata."
    )

    while True:
        raw = input("Nhập lựa chọn: ").strip().lower()
        if raw in {"", "all", "a", "tat ca", "tất cả"}:
            return all_options

        choices = []
        for part in [part for part in re.split(r"[\s,;]+", raw) if part]:
            if part not in choices:
                choices.append(part)

        if choices and set(choices).issubset(valid_options):
            return choices

        print("Lựa chọn không hợp lệ. Ví dụ đúng: 1,2 hoặc 1 3 hoặc 2,4,3,5 hoặc all.")


def get_display_date():
    return datetime.now().strftime("%d/%m/%Y")


def ensure_output_dirs():
    docx_dir = OUTPUT_DIR / "docx" / "notebook"
    pdf_dir = OUTPUT_DIR / "pdf" / "notebook"
    docx_dir.mkdir(parents=True, exist_ok=True)
    pdf_dir.mkdir(parents=True, exist_ok=True)
    return docx_dir, pdf_dir


def build_document():
    choices = ask_export_options()
    display_date = get_display_date()
    docx_dir, pdf_dir = ensure_output_dirs()
    doc = setup_document()

    all_words, json_data_map = load_vocabulary_data(DATA_DIR)
    five_s_data = load_5s_data(DATA_DIR)
    special_topics = load_special_topics(DATA_DIR)

    add_toc_section(doc)

    section_map = {
        "1": (add_5s_section, five_s_data),
        "2": (add_horenso_section, special_topics["horenso"]),
        "3": (add_aisatsu_section, special_topics["aisatsu"]),
        "4": (add_garbage_section, special_topics["garbage_sorting"]),
        "5": (add_japanese_index, all_words),
        "6": (add_vietnamese_index, all_words),
        "7": (add_grouped_vocabulary, json_data_map),
    }

    for choice in choices:
        add_func, data = section_map[choice]
        add_section_with_break(doc, add_func, data)

    qr_path = ASSETS_DIR / "images" / "qr" / "so_tay_a5_online_qr.png"
    add_section_with_break(doc, add_online_version_page, qr_path, display_date)

    docx_path = docx_dir / "so_tay_a5_content.docx"
    content_pdf_path = pdf_dir / "so_tay_a5_content.pdf"
    final_pdf_path = pdf_dir / "so_tay_a5_full.pdf"
    cover_pdf_path = prepare_cover_pdf(display_date)

    doc.save(docx_path)
    print(f"Đã tạo file Word: {docx_path}")

    update_docx_fields_and_export_pdf(docx_path, content_pdf_path)
    print(f"Đã tạo PDF nội dung: {content_pdf_path}")

    if cover_pdf_path.exists():
        merge_pdfs([cover_pdf_path, content_pdf_path], final_pdf_path)
        print(f"Đã ghép bìa và tạo PDF hoàn chỉnh: {final_pdf_path}")
    else:
        print(f"Không tìm thấy file bìa: {cover_pdf_path}")


def main():
    configure_stdio()
    build_document()
