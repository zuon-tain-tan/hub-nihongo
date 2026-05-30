import logging
import re

from docx import Document

from config import (
    DOCUMENT_FONT_SIZE_PT,
    JAPANESE_FONT,
    OUTPUT_DIR,
    TEMPLATES_DIR,
    VIETNAMESE_FONT,
)
from docx_utils import set_paragraph_text_keep_first_run


def update_docx_fields_and_export_pdf(docx_path, pdf_path):
    import win32com.client

    if pdf_path.exists():
        pdf_path.unlink()

    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(str(docx_path.resolve()))
        try:
            doc.Fields.Update()
            for toc in doc.TablesOfContents:
                toc.Update()
                toc.Range.ParagraphFormat.SpaceBefore = 0
                toc.Range.ParagraphFormat.SpaceAfter = 0
                toc.Range.ParagraphFormat.LineSpacingRule = 1

            for style_name in ["TOC 1", "TOC 2", "TOC 3", "Normal"]:
                try:
                    style = doc.Styles(style_name)
                    style.Font.Name = VIETNAMESE_FONT
                    style.Font.NameFarEast = JAPANESE_FONT
                    style.Font.Size = DOCUMENT_FONT_SIZE_PT
                    style.ParagraphFormat.SpaceBefore = 0
                    style.ParagraphFormat.SpaceAfter = 0
                    style.ParagraphFormat.LineSpacingRule = 1
                except Exception:
                    pass
        except Exception:
            pass

        doc.Save()
        doc.ExportAsFixedFormat(str(pdf_path.resolve()), 17)
        doc.Close(False)
    finally:
        word.Quit()


def update_cover_date(cover_docx_path, display_date):
    doc = Document(cover_docx_path)
    changed = False
    for paragraph in doc.paragraphs:
        text = paragraph.text
        if not text:
            continue

        if re.search(r"Last Updated:", text, re.IGNORECASE):
            set_paragraph_text_keep_first_run(paragraph, f"Last Updated: {display_date}")
            changed = True
            continue

        if re.search(r"Cập nhật lần cuối", text, re.IGNORECASE):
            set_paragraph_text_keep_first_run(paragraph, f"(Cập nhật lần cuối {display_date})")
            changed = True

    if changed:
        doc.save(cover_docx_path)


def prepare_cover_pdf(display_date):
    cover_docx_path = TEMPLATES_DIR / "docx" / "so_tay_a5_cover.docx"
    cover_pdf_path = OUTPUT_DIR / "pdf" / "cover" / "so_tay_a5_cover.pdf"
    cover_pdf_path.parent.mkdir(parents=True, exist_ok=True)

    if cover_docx_path.exists():
        update_cover_date(cover_docx_path, display_date)
        update_docx_fields_and_export_pdf(cover_docx_path, cover_pdf_path)
        print(f"Đã cập nhật ngày bìa và xuất PDF bìa: {cover_pdf_path}")
        return cover_pdf_path

    return cover_pdf_path


def merge_pdfs(input_paths, output_path):
    from pypdf import PdfWriter

    logging.getLogger("pypdf").setLevel(logging.ERROR)
    writer = PdfWriter()
    for input_path in input_paths:
        writer.append(str(input_path))

    with output_path.open("wb") as file:
        writer.write(file)
