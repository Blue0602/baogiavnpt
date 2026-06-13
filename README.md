# Tool tạo Báo giá FiberVNN - Streamlit

## 1. Cài thư viện

```bash
pip install -r requirements.txt
```

## 2. Chạy app

```bash
streamlit run app.py
```

## 3. Xuất PDF

App luôn xuất được DOCX. Muốn xuất PDF tự động, máy cần cài LibreOffice.

- Windows: cài LibreOffice rồi khởi động lại terminal.
- Ubuntu/Debian:

```bash
sudo apt update
sudo apt install libreoffice -y
```

## 4. Cách sửa template

File mẫu nằm ở:

```text
templates/bao_gia_fiber_template.docx
```

Các placeholder đang dùng:

```text
{{ ten_khach_hang }}
{{ dia_chi_lap_dat }}
{{ account }}
{{ ma_thanh_toan }}
{{ tieu_de_cuoc_su_dung }}
{{ noi_dung_dich_vu }}
{{ don_gia }}
{{ so_luong }}
{{ thanh_tien }}
{{ ghi_chu_vat }}
{{ bang_chu }}
{{ ten_tai_khoan }}
{{ so_tai_khoan }}
{{ ngan_hang }}
{{ noi_dung_thanh_toan }}
{{ nhan_vien }}
{{ sdt_nhan_vien }}
{{ email_nhan_vien }}
```

Lưu ý: không tách placeholder thành nhiều font/style khác nhau trong Word, ví dụ không để `{{ ten_` một kiểu chữ và `khach_hang }}` kiểu chữ khác.
