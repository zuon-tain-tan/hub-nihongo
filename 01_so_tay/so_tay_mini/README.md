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
|   `-- vocabulary/n5/
|-- output/                 # Generated DOCX/PDF/PPTX files
|   |-- docx/notebook/
|   |-- pdf/cover/
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

Neu cai editable package:

```powershell
pip install -e .
so-tay-mini
```

Output mac dinh:

- `output/docx/notebook/so_tay_a5_content.docx`
- `output/pdf/cover/so_tay_a5_cover.pdf`
- `output/pdf/notebook/so_tay_a5_content.pdf`
- `output/pdf/notebook/so_tay_a5_full.pdf`
