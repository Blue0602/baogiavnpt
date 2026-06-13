# -*- coding: utf-8 -*-
"""
Tool tạo Báo giá FiberVNN bằng Streamlit.
Luồng dùng:
1) Mở app
2) Điền form
3) Bấm "Tạo báo giá"
4) Tải DOCX/PDF

Chạy app:
    streamlit run app.py

Yêu cầu:
- Có file template DOCX trong templates/bao_gia_fiber_template.docx
- Nếu muốn xuất PDF, máy/server cần cài LibreOffice.
"""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import tempfile
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, Any, Optional

import streamlit as st
from dateutil.relativedelta import relativedelta
from docxtpl import DocxTemplate

# =========================
# CẤU HÌNH ĐƯỜNG DẪN
# =========================
BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_DIR = BASE_DIR / "templates"
OUTPUT_DOCX_DIR = BASE_DIR / "outputs" / "docx"
OUTPUT_PDF_DIR = BASE_DIR / "outputs" / "pdf"
DEFAULT_TEMPLATE = TEMPLATE_DIR / "bao_gia_fiber_template.docx"

OUTPUT_DOCX_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_PDF_DIR.mkdir(parents=True, exist_ok=True)


# =========================
# HÀM TIỆN ÍCH FORMAT
# =========================
def format_vnd(value: int | float | str) -> str:
    """Format số tiền kiểu Việt Nam: 10296000 -> 10.296.000"""
    try:
        number = int(float(str(value).replace(".", "").replace(",", "")))
    except Exception:
        number = 0
    return f"{number:,}".replace(",", ".")


def clean_filename(text: str) -> str:
    """Làm sạch tên file để tránh lỗi ký tự đặc biệt."""
    text = text.strip().lower()
    replacements = {
        "đ": "d", "Đ": "D",
        "á": "a", "à": "a", "ả": "a", "ã": "a", "ạ": "a",
        "ă": "a", "ắ": "a", "ằ": "a", "ẳ": "a", "ẵ": "a", "ặ": "a",
        "â": "a", "ấ": "a", "ầ": "a", "ẩ": "a", "ẫ": "a", "ậ": "a",
        "é": "e", "è": "e", "ẻ": "e", "ẽ": "e", "ẹ": "e",
        "ê": "e", "ế": "e", "ề": "e", "ể": "e", "ễ": "e", "ệ": "e",
        "í": "i", "ì": "i", "ỉ": "i", "ĩ": "i", "ị": "i",
        "ó": "o", "ò": "o", "ỏ": "o", "õ": "o", "ọ": "o",
        "ô": "o", "ố": "o", "ồ": "o", "ổ": "o", "ỗ": "o", "ộ": "o",
        "ơ": "o", "ớ": "o", "ờ": "o", "ở": "o", "ỡ": "o", "ợ": "o",
        "ú": "u", "ù": "u", "ủ": "u", "ũ": "u", "ụ": "u",
        "ư": "u", "ứ": "u", "ừ": "u", "ử": "u", "ữ": "u", "ự": "u",
        "ý": "y", "ỳ": "y", "ỷ": "y", "ỹ": "y", "ỵ": "y",
    }
    for k, v in replacements.items():
        text = text.replace(k, v)
    text = re.sub(r"[^a-z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text[:80] or "bao_gia"


def format_date_vn(d: date) -> str:
    return d.strftime("%d/%m/%Y")


def calc_end_date(start_date: date, total_months: int) -> date:
    """Ngày kết thúc = ngày bắt đầu + tổng tháng - 1 ngày."""
    return start_date + relativedelta(months=total_months) - timedelta(days=1)


# =========================
# ĐỌC TIỀN BẰNG CHỮ TIẾNG VIỆT
# =========================
_DIGITS = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"]
_UNITS = ["", "ngàn", "triệu", "tỷ", "ngàn tỷ", "triệu tỷ"]


def _read_three_digits(n: int, full: bool = False) -> str:
    """Đọc số từ 0-999. full=True để đọc đủ hàng trăm khi nằm giữa số lớn."""
    assert 0 <= n <= 999
    hundred = n // 100
    ten = (n % 100) // 10
    unit = n % 10
    words = []

    if hundred > 0:
        words += [_DIGITS[hundred], "trăm"]
    elif full and (ten > 0 or unit > 0):
        words += ["không", "trăm"]

    if ten > 1:
        words += [_DIGITS[ten], "mươi"]
        if unit == 1:
            words.append("mốt")
        elif unit == 4:
            words.append("bốn")
        elif unit == 5:
            words.append("lăm")
        elif unit > 0:
            words.append(_DIGITS[unit])
    elif ten == 1:
        words.append("mười")
        if unit == 5:
            words.append("lăm")
        elif unit > 0:
            words.append(_DIGITS[unit])
    else:
        if unit > 0:
            if hundred > 0 or full:
                words.append("lẻ")
            words.append(_DIGITS[unit])

    return " ".join(words)


def number_to_vietnamese_words(number: int) -> str:
    """Đọc số nguyên dương sang tiếng Việt, dùng cho tiền VND."""
    number = int(number)
    if number == 0:
        return "Không đồng"
    if number < 0:
        return "Âm " + number_to_vietnamese_words(abs(number)).lower()

    groups = []
    while number > 0:
        groups.append(number % 1000)
        number //= 1000

    parts = []
    highest_idx = len(groups) - 1
    for idx in range(highest_idx, -1, -1):
        group = groups[idx]
        if group == 0:
            continue
        full = idx != highest_idx
        part = _read_three_digits(group, full=full)
        unit = _UNITS[idx] if idx < len(_UNITS) else ""
        if unit:
            part = f"{part} {unit}"
        parts.append(part)

    text = " ".join(parts).strip()
    return text[:1].upper() + text[1:] + " đồng"


# =========================
# RENDER DOCX / PDF
# =========================
def find_libreoffice() -> Optional[str]:
    """Tìm executable LibreOffice/soffice."""
    candidates = ["soffice", "libreoffice"]
    for name in candidates:
        path = shutil.which(name)
        if path:
            return path
    return None


def render_docx(template_path: Path, context: Dict[str, Any], output_docx: Path) -> Path:
    """Render DOCX bằng docxtpl."""
    doc = DocxTemplate(str(template_path))
    doc.render(context)
    doc.save(str(output_docx))
    return output_docx


def convert_docx_to_pdf(docx_path: Path, output_dir: Path) -> Optional[Path]:
    """Convert DOCX sang PDF bằng LibreOffice headless. Trả None nếu không có LibreOffice."""
    lo = find_libreoffice()
    if not lo:
        return None

    output_dir.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp_home:
        env = os.environ.copy()
        env["HOME"] = tmp_home
        cmd = [
            lo,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(output_dir),
            str(docx_path),
        ]
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, env=env)

    pdf_path = output_dir / (docx_path.stem + ".pdf")
    return pdf_path if pdf_path.exists() else None


def save_uploaded_template(uploaded_file) -> Path:
    """Lưu template DOCX người dùng upload tạm thời."""
    tmp_dir = tempfile.mkdtemp(prefix="template_")
    tmp_path = Path(tmp_dir) / uploaded_file.name
    tmp_path.write_bytes(uploaded_file.getbuffer())
    return tmp_path


# =========================
# STREAMLIT UI
# =========================
st.set_page_config(page_title="Tạo báo giá FiberVNN", page_icon="📄", layout="centered")

st.title("📄 Tool tạo Báo giá FiberVNN")
st.caption("Mở app → điền form → bấm tạo báo giá → tải DOCX/PDF")

with st.sidebar:
    st.header("⚙️ Cấu hình")
    uploaded_template = st.file_uploader(
        "Upload template DOCX riêng nếu có",
        type=["docx"],
        help="Nếu không upload, app sẽ dùng templates/bao_gia_fiber_template.docx",
    )
    st.info("PDF cần máy/server có cài LibreOffice. Nếu chưa có, app vẫn xuất DOCX bình thường.")

if not DEFAULT_TEMPLATE.exists() and uploaded_template is None:
    st.error("Không tìm thấy template mặc định: templates/bao_gia_fiber_template.docx")
    st.stop()

with st.form("quote_form"):
    st.subheader("1. Thông tin khách hàng")
    customer_name = st.text_input("Tên khách hàng", value="CÔNG TY TRÁCH NHIỆM HỮU HẠN HANGDO VINA")
    install_address = st.text_area(
        "Địa chỉ lắp đặt",
        value="Đường Số 6B KCN Nhơn Trạch 1, Xã Phú Hội, Nhơn Trạch, Thành phố Đồng Nai",
        height=70,
    )
    col1, col2 = st.columns(2)
    with col1:
        account_no = st.text_input("Acc/số máy", value="hangdovina12304524")
    with col2:
        payment_code = st.text_input("Mã thanh toán", value="DNI-13-0088123")

    st.subheader("2. Gói cước & thời gian")
    col1, col2 = st.columns(2)
    with col1:
        package_name = st.text_input("Gói cước", value="fiberEco3")
        start_date = st.date_input("Ngày bắt đầu", value=date(2026, 7, 10), format="DD/MM/YYYY")
    with col2:
        paid_months = st.number_input("Số tháng trả trước", min_value=1, max_value=60, value=12, step=1)
        bonus_months = st.number_input("Số tháng tặng", min_value=0, max_value=24, value=2, step=1)

    st.subheader("3. Tiền & thanh toán")
    price_mode = st.radio(
        "Cách tính tiền",
        ["Nhập tổng giá gói", "Tính theo giá/tháng x số tháng trả trước"],
        horizontal=False,
    )

    col1, col2 = st.columns(2)
    with col1:
        if price_mode == "Nhập tổng giá gói":
            price_input = st.number_input("Đơn giá/Tổng giá gói (VNĐ)", min_value=0, value=10296000, step=1000)
        else:
            price_input = st.number_input("Giá/tháng (VNĐ)", min_value=0, value=858000, step=1000)
    with col2:
        quantity = st.number_input("Số lượng", min_value=1, max_value=100, value=1, step=1)

    col1, col2 = st.columns(2)
    with col1:
        bank_account_name = st.text_input("Tên tài khoản", value="Viễn Thông Đồng Nai – Tập Đoàn Bưu Chính Viễn Thông Việt Nam")
        bank_account_no = st.text_input("Số tài khoản", value="0121000771395")
    with col2:
        bank_name = st.text_input("Ngân hàng", value="Vietcombank CN Đồng Nai")
        vat_note = st.text_input("Ghi chú VAT", value="Đã bao gồm thuế VAT")

    st.subheader("4. Nhân viên phụ trách")
    col1, col2 = st.columns(2)
    with col1:
        staff_name = st.text_input("Tên nhân viên", value="Vương Thanh Thuận")
        staff_phone = st.text_input("SĐT", value="0837892579")
    with col2:
        staff_email = st.text_input("Email", value="thuanvt.dni@vnpt.vn")
        staff_code = st.text_input("Mã NV/CTV", value="CTV092914")

    submitted = st.form_submit_button("🚀 Tạo báo giá DOCX/PDF", use_container_width=True)

if submitted:
    errors = []
    if not customer_name.strip():
        errors.append("Thiếu tên khách hàng")
    if not install_address.strip():
        errors.append("Thiếu địa chỉ lắp đặt")
    if not account_no.strip():
        errors.append("Thiếu Acc/số máy")
    if not payment_code.strip():
        errors.append("Thiếu mã thanh toán")
    if not package_name.strip():
        errors.append("Thiếu gói cước")
    if price_input <= 0:
        errors.append("Số tiền phải lớn hơn 0")

    if errors:
        st.error("Không thể tạo báo giá:\n- " + "\n- ".join(errors))
        st.stop()

    total_months = int(paid_months + bonus_months)
    end_date = calc_end_date(start_date, total_months)

    if price_mode == "Nhập tổng giá gói":
        unit_price = int(price_input)
        total_amount = int(price_input * quantity)
    else:
        unit_price = int(price_input * paid_months)
        total_amount = int(price_input * paid_months * quantity)

    amount_in_words = number_to_vietnamese_words(total_amount)
    service_content = (
        f"Gia hạn trả trước gói trả trước {total_months} tháng gói {package_name} "
        f"( tính từ {format_date_vn(start_date)} - {format_date_vn(end_date)})"
    )
    payment_content = f"Thanh toán trả trước {paid_months} tháng – {payment_code} – {staff_code}"
    usage_title = f"Cước sử dụng {total_months} tháng"

    context = {
        "ngay_tao": format_date_vn(date.today()),
        "ten_khach_hang": customer_name.strip(),
        "dia_chi_lap_dat": install_address.strip(),
        "account": account_no.strip(),
        "ma_thanh_toan": payment_code.strip(),
        "goi_cuoc": package_name.strip(),
        "so_thang_tra_truoc": int(paid_months),
        "so_thang_tang": int(bonus_months),
        "tong_so_thang": total_months,
        "ngay_bat_dau": format_date_vn(start_date),
        "ngay_ket_thuc": format_date_vn(end_date),
        "noi_dung_dich_vu": service_content,
        "tieu_de_cuoc_su_dung": usage_title,
        "don_gia": format_vnd(unit_price),
        "so_luong": int(quantity),
        "thanh_tien": format_vnd(total_amount),
        "bang_chu": amount_in_words,
        "ghi_chu_vat": vat_note.strip(),
        "ten_tai_khoan": bank_account_name.strip(),
        "so_tai_khoan": bank_account_no.strip(),
        "ngan_hang": bank_name.strip(),
        "noi_dung_thanh_toan": payment_content,
        "nhan_vien": staff_name.strip(),
        "sdt_nhan_vien": staff_phone.strip(),
        "email_nhan_vien": staff_email.strip(),
        "ma_nhan_vien": staff_code.strip(),
    }

    template_path = save_uploaded_template(uploaded_template) if uploaded_template else DEFAULT_TEMPLATE
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_customer = clean_filename(customer_name)
    output_docx = OUTPUT_DOCX_DIR / f"Bao_gia_Fiber_{safe_customer}_{timestamp}.docx"

    try:
        render_docx(template_path, context, output_docx)
        pdf_path = None
        try:
            pdf_path = convert_docx_to_pdf(output_docx, OUTPUT_PDF_DIR)
        except subprocess.CalledProcessError as e:
            st.warning("Đã tạo DOCX nhưng convert PDF lỗi. Kiểm tra LibreOffice hoặc template DOCX.")
        except Exception as e:
            st.warning(f"Đã tạo DOCX nhưng chưa tạo được PDF: {e}")

        st.success("Tạo báo giá thành công")

        st.subheader("Xem nhanh dữ liệu đã tính")
        st.write({
            "Tổng số tháng": total_months,
            "Ngày kết thúc": format_date_vn(end_date),
            "Thành tiền": format_vnd(total_amount),
            "Bằng chữ": amount_in_words,
            "Nội dung thanh toán": payment_content,
        })

        with open(output_docx, "rb") as f:
            st.download_button(
                label="⬇️ Tải file Word DOCX",
                data=f.read(),
                file_name=output_docx.name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

        if pdf_path and pdf_path.exists():
            with open(pdf_path, "rb") as f:
                st.download_button(
                    label="⬇️ Tải file PDF",
                    data=f.read(),
                    file_name=pdf_path.name,
                    mime="application/pdf",
                    use_container_width=True,
                )
        else:
            st.info("Chưa có PDF. Cài LibreOffice để app tự convert DOCX → PDF.")

    except Exception as e:
        st.exception(e)
