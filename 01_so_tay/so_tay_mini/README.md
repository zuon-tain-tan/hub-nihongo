# So Tay Mini

Du an tao so tay A5 tieng Nhat - tieng Viet tu du lieu JSON, xuat DOCX va PDF bang `python-docx` va Microsoft Word COM tren Windows.

## Cau Truc Thu Muc

```text
.
|-- assets/                 # Static assets and images
|   `-- images/qr/
|-- data/                   # Source data
|   |-- kana/
|   |-- topics/
|   `-- vocabulary/
|       |-- n5/
|       `-- n4/
|-- output/                 # Generated DOCX/PDF/PPTX files
|   |-- docx/kanji/
|   |-- docx/notebook/
|   |-- pdf/cover/
|   |-- pdf/kanji/
|   |-- pdf/notebook/
|   `-- pptx/
|-- main.py                 # Main runner
|-- tieu_chi.md             # Document formatting criteria
|-- src/                    # Python source modules
`-- templates/              # Editable document templates
    `-- docx/
```

## Chay Du An

```powershell
pip install -r requirements.txt
python main.py
```

## Cai Selenium Cho Tinh Nang 4

Tinh nang `4. Hinh anh tu vung` dung Selenium de mo trang Langoal va tai anh tu vung. May can co Chrome hoac Edge.

```powershell
pip install -r requirements.txt
```

Neu Selenium bao khong tim thay trinh duyet, hay cai Google Chrome hoac Microsoft Edge ban moi nhat roi chay lai `python main.py`.

Khi chay tinh nang 4, co the nhap mot bai hoac nhieu bai:

```text
1
1-50
1,3,5
1-3,7
```

Gioi han hop le la bai 1 den bai 50. Bai nao da co `data/image_vocabulary/lesson_xx.json` thi chuong trinh se dung lai metadata san co va bo qua buoc tai lai.

Neu cai editable package:

```powershell
pip install -e .
so-tay-mini
```

Output mac dinh:

- `output/docx/kanji/*.docx`
- `output/pdf/kanji/*.pdf`
- `output/docx/notebook/so_tay_a5_content.docx`
- `output/pdf/cover/so_tay_a5_cover.pdf`
- `output/pdf/notebook/so_tay_a5_content.pdf`
- `output/pdf/notebook/so_tay_a5_full.pdf`
