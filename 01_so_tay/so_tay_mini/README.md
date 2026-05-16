# Sổ tay mini

## Cấu trúc

- `scripts/`: script thao tác dự án.
- `assets/`: tài nguyên dùng khi tạo tài liệu, ví dụ QR.
- `tuvung/`: dữ liệu JSON.
- `pdf/bia/`: file bìa PDF.
- `word/bia/`: file bìa Word dùng để xuất lại PDF bìa.
- `pdf/sotay/`: PDF kết quả.
- `word/sotay/`: Word kết quả.
- `word/legacy/`: file Word cũ trước khi tái cấu trúc.
- `mp3/`, `ppt/`: tài nguyên học liệu.

```text
so_tay_mini/
├─ a.py                         # launcher tạo sổ tay
├─ c.py                         # launcher xóa mp3 theo giọng
├─ scripts/
│  ├─ build_sotay.py
│  └─ delete_mp3_by_voice.py
├─ assets/
│  └─ qr/so_tay_a5_online_qr.png
├─ tuvung/
│  ├─ 5S/
│  ├─ chu_cai/
│  └─ tu_vungN5/
├─ word/
│  ├─ bia/so_tay_a5_bia.docx
│  ├─ sotay/so_tay_a5_no_bia.docx
│  └─ legacy/
└─ pdf/
   ├─ bia/so_tay_a5_bia.pdf
   └─ sotay/
      ├─ so_tay_a5_no_bia.pdf
      └─ so_tay_a5.pdf
```

## Tạo sổ tay

Chạy từ thư mục gốc:

```powershell
python a.py
```

Hoặc chạy script chính:

```powershell
python scripts/build_sotay.py
```

Khi được hỏi, nhập ví dụ:

- `1`
- `1,3`
- `1 2 3`
- `all`

Kết quả:

- `word/sotay/so_tay_a5_no_bia.docx`
- `pdf/sotay/so_tay_a5_no_bia.pdf`
- `pdf/sotay/so_tay_a5.pdf`

Khi chạy, script tự cập nhật ngày trên `word/bia/so_tay_a5_bia.docx`, xuất lại `pdf/bia/so_tay_a5_bia.pdf`, rồi ghép bìa vào PDF sổ tay.

Trang cuối dùng QR tại `assets/qr/so_tay_a5_online_qr.png`.

## Phụ thuộc

```powershell
python -m pip install -r requirements.txt
```

Script xuất PDF và cập nhật mục lục bằng Microsoft Word qua COM, nên cần chạy trên Windows có Word.
