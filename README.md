# Tool Streamlit tạo Báo giá FiberVNN

Luồng dùng:

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Template

File template chính nằm tại:

```text
templates/bao_gia_fiber_template.docx
```

Template này được làm từ file báo giá gốc có sẵn banner VNPT và dấu mộc đỏ. App chỉ thay các placeholder text, không xử lý/chèn ảnh bằng code.

Các placeholder đang dùng:

```text
{{ ten_khach_hang }}
{{ dia_chi_lap_dat }}
{{ account }}
{{ ma_thanh_toan }}
{{ tieu_de_cuoc_su_dung }}
{{ noi_dung_dich_vu }}
{{ so_luong }}
{{ don_gia }}
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

## PDF

Xuất PDF cần cài LibreOffice trên máy/server. Nếu chưa cài, app vẫn tạo được DOCX.
