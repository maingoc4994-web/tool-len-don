import base64
import io
import json
import os
import re
import time
import unicodedata
from copy import copy
from datetime import date, datetime
from difflib import SequenceMatcher

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import requests
import streamlit as st
from google import genai
from PIL import Image

# ==============================================================================
# CẤU HÌNH GIAO DIỆN TỐI ƯU ĐIỆN THOẠI (MOBILE-FIRST)
# ==============================================================================
st.set_page_config(
    page_title="Tool hóa đơn MTT - Trái cây",
    page_icon="🍎",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    .main { background-color: #f4f6f8; }
    .stButton>button {
        width: 100%;
        background-color: #2e7d32;
        color: white;
        font-weight: bold;
        border-radius: 10px;
        padding: 12px;
        border: none;
        font-size: 16px;
    }
    .stButton>button:hover { background-color: #1b5e20; color: white; }
    h1 { color: #1b5e20; text-align: center; font-size: 1.65rem !important; }
    .stDownloadButton>button {
        width: 100%;
        background-color: #0277bd;
        color: white;
        font-weight: bold;
        border-radius: 10px;
        padding: 12px;
        font-size: 16px;
    }
    .stDownloadButton>button:hover { background-color: #01579b; color: white; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown("<h1>🍎 Tool Lên Hóa Đơn MTT (Trái Cây)</h1>", unsafe_allow_html=True)
st.markdown(
    "<p style='text-align: center; color: #555; font-size: 0.95rem;'>"
    "Quét ảnh đơn hàng bằng AI Gemini, tự động gợi ý khách hàng từ tên file, "
    "gom đơn theo khách và xuất file Excel giá trị số thuần chuẩn MTT.</p>",
    unsafe_allow_html=True,
)

# ==============================================================================
# BẢO MẬT API KEY & DANH SÁCH MODEL DỰ PHÒNG
# ==============================================================================
try:
    API_KEY = st.secrets["GEMINI_API_KEY"]
except Exception:
    st.error("⚠️ Chưa cấu hình GEMINI_API_KEY trong mục Secrets của Streamlit Cloud!")
    st.stop()

FALLBACK_MODELS = [
    "gemini-3.5-flash",
    "gemini-2.5-flash",
    "gemini-3.5-flash-lite",
]

CUSTOMER_COLS = [
    "STT",
    "Mã khách hàng",
    "Tên khách hàng",
    "MST/CCCD chủ hộ",
    "MSĐVCQHVNS",
    "Căn cước công dân",
    "Địa chỉ",
    "Người nhận HĐ",
    "Email",
    "Số điện thoại",
    "Ngừng sử dụng",
]

DEFAULT_CUSTOMER_PATHS = [
    "Danh_sach_khach_hang.xlsx",
    "Danh_sach_khach_hang.xls",
    "Danh_sach_khach_hang (2).xls",
    "Danh_sach_khach_hang (2).xlsx",
]

DEFAULT_TEMPLATE_PATHS = [
    "Mau_hoa_don_ban_hang_MTT_2.xls",
    "Mau_hoa_don_ban_hang_MTT_2.xlsx",
    "Mau_hoa_don_ban_hang_MTT.xlsx",
    "Mau_hoa_don_ban_hang_MTT.xls",
]

MTT_HEADERS = [
    "Số thứ tự hóa đơn (*)",
    "Ngày hóa đơn",
    "Tên đơn vị mua hàng",
    "Mã số thuế",
    "Địa chỉ",
    "Người mua hàng",
    "Email",
    "Số điện thoại",
    "Căn cước công dân",
    "Hình thức thanh toán (*)",
    "Tên hàng hóa/dịch vụ (*)",
    "ĐVT",
    "Số lượng",
    "Đơn giá",
    "Thành tiền",
]

MTT_COL_WIDTHS = [
    23.15, 20.71, 30.70, 20.86, 30.01,
    20.86, 30.70, 18.58, 21.43, 24.86,
    25.71, 15.71, 15.71, 25.71, 25.71,
]

ABBREV_MAP = {
    "vp": "van phong",
    "ubnd": "uy ban nhan dan",
    "hdnd": "hdnd",
    "dbqh": "dbqh",
    "th": "tieu hoc",
    "cd": "cao dang",
    "cty": "cong ty",
    "ct": "cong ty",
    "ks": "khach san",
    "tt": "trung tam",
    "bt": "bao tang",
}


# ==============================================================================
# TẠO FILE MẪU MTT TÍCH HỢP SẴN (KHÔNG CẦN TẢI FILE MẪU MỖI LẦN MỞ APP)
# ==============================================================================
def create_builtin_mtt_template() -> bytes:
    """Tạo file mẫu hóa đơn bán hàng MTT chuẩn 15 cột (sheet 'Hóa đơn bán hàng', dòng 1-8)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Hóa đơn bán hàng"

    font_title = Font(name="Times New Roman", size=14, bold=True)
    font_guide_bold = Font(name="Times New Roman", size=11, bold=True, italic=True)
    font_guide = Font(name="Times New Roman", size=11, italic=True)
    font_header = Font(name="Times New Roman", size=12, bold=True)
    font_data = Font(name="Times New Roman", size=12)

    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    thin_side = Side(style="thin", color="BFBFBF")
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    ws["A1"] = "File mẫu danh sách hóa đơn để nhập vào phần mềm "
    ws["A1"].font = font_title
    ws["A2"] = "Hướng dẫn:"
    ws["A2"].font = font_guide_bold
    ws["A3"] = "- Điền dữ liệu hóa đơn cần lập trên phần mềm vào các cột tương ứng trên file này"
    ws["A4"] = "- Các cột có dấu (*) là những cột bắt buộc"
    ws["A5"] = "- Nếu muốn nhập thêm thông tin khác, người dùng có thể tự thêm cột trên file này (VD: Mã khách hàng, Mã hàng, Tỷ lệ chiết khấu)"
    ws["A6"] = "- Các dòng dữ liệu phía dưới chỉ là ví dụ minh họa"
    for r in range(3, 7):
        ws.cell(row=r, column=1).font = font_guide

    ws.row_dimensions[8].height = 25
    for col_idx, (title, width) in enumerate(zip(MTT_HEADERS, MTT_COL_WIDTHS), start=1):
        col_letter = openpyxl.utils.get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = width

        c8 = ws.cell(row=8, column=col_idx, value=title)
        c8.font = font_header
        c8.fill = header_fill
        c8.border = thin_border
        c8.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # Tạo định dạng chuẩn cho dòng 9 để sao chép xuống các dòng dữ liệu
        c9 = ws.cell(row=9, column=col_idx, value=None)
        c9.font = font_data
        c9.border = thin_border
        if col_idx in (13, 14, 15):
            c9.alignment = Alignment(horizontal="right", vertical="center")
        elif col_idx in (1, 2, 12):
            c9.alignment = Alignment(horizontal="center", vertical="center")
        else:
            c9.alignment = Alignment(horizontal="left", vertical="center")

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ==============================================================================
# HÀM GỌI GEMINI TỰ ĐỘNG THỬ LẠI & CHUYỂN MODEL DỰ PHÒNG
# ==============================================================================
def generate_with_fallback(client, contents, status_placeholder=None):
    """Tự động thử lại khi gặp lỗi 503/429 và chuyển sang model dự phòng nếu vẫn bận."""
    last_err = None
    for idx, model_name in enumerate(FALLBACK_MODELS):
        for attempt in range(2):
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=contents,
                )
                return response, model_name
            except Exception as e:
                last_err = e
                err_str = str(e)
                is_overload = any(
                    k in err_str
                    for k in ("503", "UNAVAILABLE", "429", "RESOURCE_EXHAUSTED")
                )
                is_not_found = "404" in err_str or "NOT_FOUND" in err_str

                if is_not_found:
                    if status_placeholder and idx < len(FALLBACK_MODELS) - 1:
                        status_placeholder.info(
                            f"🔄 Model `{model_name}` không khả dụng, chuyển sang `{FALLBACK_MODELS[idx + 1]}`..."
                        )
                    break

                if is_overload:
                    if attempt == 0:
                        if status_placeholder:
                            status_placeholder.info(
                                f"⏳ `{model_name}` đang quá tải, tự thử lại sau 3 giây..."
                            )
                        time.sleep(3)
                        continue
                    if status_placeholder and idx < len(FALLBACK_MODELS) - 1:
                        status_placeholder.info(
                            f"🔄 `{model_name}` vẫn bận, tự chuyển sang `{FALLBACK_MODELS[idx + 1]}`..."
                        )
                    break

                raise e
    raise last_err


# ==============================================================================
# HÀM TIỆN ÍCH CHUẨN HÓA & GỢI Ý KHÁCH HÀNG
# ==============================================================================
def strip_accents(text: str) -> str:
    if not text:
        return ""
    text = str(text).replace("đ", "d").replace("Đ", "D")
    nfkd = unicodedata.normalize("NFKD", text)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def clean_str(val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = unicodedata.normalize("NFC", str(val)).strip()
    return "" if s.lower() == "nan" else s


def to_number(val, default=0.0):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return default
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s:
        return default
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        if len(parts) == 2 and len(parts[1]) <= 2:
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    elif s.count(".") > 1:
        s = s.replace(".", "")
    elif "." in s:
        parts = s.split(".")
        if len(parts) == 2 and len(parts[1]) == 3 and len(parts[0]) >= 1:
            s = s.replace(".", "")
    try:
        return float(s)
    except Exception:
        return default


def clean_excel_num(val):
    num = to_number(val, 0.0)
    if abs(num - round(num)) < 1e-9:
        return int(round(num))
    return round(num, 4)


def process_item_name(name: str) -> str:
    """Quy tắc tên hàng hộ kinh doanh trái cây:
    - Xóa hoàn toàn chữ 'Total' (không phân biệt hoa thường).
    - Xóa phần '/...' phía sau nếu là chữ tiếng Anh.
    - Chuẩn hóa khoảng trắng và viết hoa chữ cái đầu mỗi từ.
    """
    if not name:
        return ""
    name = str(name)
    name = re.sub(r"(?i)\btotal\b", "", name).strip()
    name = re.sub(r"\s*/\s*[a-zA-Z].*$", "", name).strip()
    name = re.sub(r"\s+", " ", name)
    words = name.split()
    return " ".join(w.lower().capitalize() for w in words)


def process_unit(name: str, unit: str, don_gia) -> str:
    """Quy tắc đơn vị tính ngành trái cây:
    - Chuối / Chuối ba lùn có đơn giá > 0 và < 8.000đ -> 'Trái'.
    - Nếu trống -> mặc định 'Kg'.
    - Có sẵn -> viết hoa chữ cái đầu.
    """
    unit_clean = clean_str(unit)
    name_lower = clean_str(name).lower()
    price = to_number(don_gia, 0.0)

    if "chuối" in name_lower and 0 < price < 8000:
        return "Trái"
    if not unit_clean:
        return "Kg"
    return unit_clean.lower().capitalize()


def clean_filename_for_match(filename: str) -> str:
    base = re.sub(r"\.[^.]+$", "", filename or "")
    base = re.sub(r"^[\d\s._\-/]+", "", base)
    base = re.sub(r"[\s._\-]+\d+$", "", base)
    base = strip_accents(base).lower()
    base = re.sub(r"[^a-z0-9\s]", " ", base)
    return re.sub(r"\s+", " ", base).strip()


def expand_tokens(cleaned: str):
    tokens = []
    for tok in cleaned.split():
        if tok in ABBREV_MAP:
            tokens.extend(ABBREV_MAP[tok].split())
        else:
            tokens.append(tok)
    return tokens


def suggest_customer_code(filename: str, df_kh: pd.DataFrame) -> str:
    """Trả về 'Mã khách hàng' gợi ý tốt nhất hoặc '' nếu không đủ ngưỡng khớp."""
    cleaned = clean_filename_for_match(filename)
    if not cleaned or df_kh.empty:
        return ""

    compact = cleaned.replace(" ", "")
    expanded_tokens = set(expand_tokens(cleaned))
    best_code = ""
    best_score = 0.0

    for _, row in df_kh.iterrows():
        ma_raw = clean_str(row.get("Mã khách hàng", ""))
        ma = strip_accents(ma_raw).lower()
        ma_compact = re.sub(r"[^a-z0-9]", "", ma)
        ten = strip_accents(clean_str(row.get("Tên khách hàng", ""))).lower()
        ten_clean = re.sub(r"[^a-z0-9\s]", " ", ten)
        ten_tokens = set(ten_clean.split())

        score = 0.0
        # Ưu tiên 1: Khớp Mã khách hàng
        if ma_compact:
            if compact == ma_compact:
                score = max(score, 1.0)
            elif len(compact) >= 3 and (compact in ma_compact or ma_compact in compact):
                ratio = min(len(compact), len(ma_compact)) / max(len(compact), len(ma_compact))
                score = max(score, 0.85 + 0.14 * ratio)

        # Ưu tiên 2: Mở rộng viết tắt và so khớp từ với Tên khách hàng
        if expanded_tokens and ten_tokens:
            matched = sum(1 for t in expanded_tokens if t in ten_tokens)
            overlap = matched / len(expanded_tokens)
            if overlap > 0:
                score = max(score, 0.65 * overlap + (0.25 if overlap == 1.0 else 0.0))

        # Ưu tiên 3: So khớp mờ
        if ma_compact:
            score = max(score, SequenceMatcher(None, compact, ma_compact).ratio() * 0.8)
        if ten_clean:
            score = max(score, SequenceMatcher(None, cleaned, ten_clean).ratio() * 0.75)

        if score > best_score:
            best_score = score
            best_code = ma_raw

    return best_code if best_score >= 0.5 else ""


# ==============================================================================
# TỰ ĐỘNG DÒ & ĐỌC FILE TRONG KHO GITHUB
# ==============================================================================
def find_existing_file(candidates, keyword=None):
    for p in candidates:
        if os.path.exists(p):
            return p
    if keyword:
        try:
            for fname in sorted(os.listdir(".")):
                low = strip_accents(fname).lower()
                if (low.endswith(".xlsx") or low.endswith(".xls")) and keyword in low:
                    return fname
        except Exception:
            pass
    return None


def load_customers_from_bytes(raw_bytes: bytes) -> pd.DataFrame:
    xls = pd.ExcelFile(io.BytesIO(raw_bytes))
    sheet = "Danh sách khách hàng" if "Danh sách khách hàng" in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=sheet, header=3, dtype=str)
    for col in CUSTOMER_COLS:
        if col not in df.columns:
            df[col] = ""
    df = df[CUSTOMER_COLS].copy()
    for col in CUSTOMER_COLS:
        df[col] = df[col].apply(clean_str)
    df = df[df["Tên khách hàng"] != ""]
    df = df[df["Ngừng sử dụng"].str.lower().isin(["", "false", "0", "không"])]
    return df.reset_index(drop=True)


def export_customers_to_xlsx_bytes(df_kh: pd.DataFrame) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Danh sách khách hàng"
    ws.cell(row=2, column=1, value="DANH SÁCH KHÁCH HÀNG")
    for c_idx, col_name in enumerate(CUSTOMER_COLS, start=1):
        ws.cell(row=4, column=c_idx, value=col_name)
    for r_idx, (_, row) in enumerate(df_kh.iterrows(), start=5):
        for c_idx, col_name in enumerate(CUSTOMER_COLS, start=1):
            val = str(r_idx - 4) if col_name == "STT" else clean_str(row.get(col_name, ""))
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.number_format = "@"
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def save_customers_to_github(xlsx_bytes: bytes, target_path: str = "Danh_sach_khach_hang.xlsx"):
    """Tự động commit danh sách khách hàng lên GitHub nếu có GITHUB_TOKEN & GITHUB_REPO."""
    token = st.secrets.get("GITHUB_TOKEN", "")
    repo = st.secrets.get("GITHUB_REPO", "")
    branch = st.secrets.get("GITHUB_BRANCH", "main")
    if not token or not repo:
        return False, "Chưa cấu hình GITHUB_TOKEN/GITHUB_REPO trong Secrets."

    url = f"https://api.github.com/repos/{repo}/contents/{target_path}"
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json",
    }
    sha = None
    get_res = requests.get(url, headers=headers, params={"ref": branch}, timeout=15)
    if get_res.status_code == 200:
        sha = get_res.json().get("sha")

    payload = {
        "message": f"Cap nhat danh sach khach hang ({datetime.now().strftime('%d/%m/%Y %H:%M')})",
        "content": base64.b64encode(xlsx_bytes).decode("utf-8"),
        "branch": branch,
    }
    if sha:
        payload["sha"] = sha

    put_res = requests.put(url, headers=headers, json=payload, timeout=20)
    if put_res.status_code in (200, 201):
        return True, "Đã tự động lưu vĩnh viễn danh sách khách hàng lên GitHub!"
    return False, f"Không thể lưu lên GitHub (mã {put_res.status_code})."


def copy_row_style(ws, src_row: int, dst_row: int, max_col: int = 15):
    for col in range(1, max_col + 1):
        src_cell = ws.cell(row=src_row, column=col)
        dst_cell = ws.cell(row=dst_row, column=col)
        if src_cell.has_style:
            dst_cell.font = copy(src_cell.font)
            dst_cell.fill = copy(src_cell.fill)
            dst_cell.border = copy(src_cell.border)
            dst_cell.alignment = copy(src_cell.alignment)
            dst_cell.number_format = src_cell.number_format


# ==============================================================================
# ĐỔ DỮ LIỆU VÀO MẪU MTT (GIÁ TRỊ SỐ THUẦN Ở CỘT O + ĐỦ CỘT C)
# ==============================================================================
def build_mtt_excel(
    template_bytes: bytes,
    items_df: pd.DataFrame,
    image_customer_map: dict,
    df_kh: pd.DataFrame,
    invoice_date_str: str,
    payment_method: str,
    group_by_customer: bool = True,
    repeat_buyer_info: bool = True,
):
    """Xuất file Excel chuẩn mẫu MTT:
    - Cột O ghi trực tiếp GIÁ TRỊ SỐ (không ghi công thức) để phần mềm hiển thị ngay.
    - Dọn sạch dataValidation/definedName lỗi của openpyxl để không mất cột C.
    - Trả về (bytes_excel, preview_df).
    """
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    sheet_name = "Hóa đơn bán hàng" if "Hóa đơn bán hàng" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet_name]

    # Dọn sạch dataValidations và definedNames lỗi (nguyên nhân gây lỗi đọc cột C khi import)
    ws.data_validations.dataValidation.clear()
    wb.defined_names.clear()

    # Đảm bảo dòng tiêu đề 8 khớp chính xác 15 cột chuẩn MTT
    for col_idx, header_text in enumerate(MTT_HEADERS, start=1):
        ws.cell(row=8, column=col_idx, value=header_text)

    # Xóa sạch TẤT CẢ 15 cột (kể cả công thức cũ ở cột O) từ dòng 9 trở đi
    if ws.max_row > 9:
        ws.delete_rows(10, ws.max_row - 9)
    for c in range(1, 16):
        cell = ws.cell(row=9, column=c)
        cell.value = None
        cell.hyperlink = None

    # Bảng tra cứu khách hàng theo Mã khách hàng
    kh_lookup = {}
    for _, r_kh in df_kh.iterrows():
        code = clean_str(r_kh.get("Mã khách hàng", ""))
        if code:
            kh_lookup[code] = r_kh.to_dict()

    working_df = items_df.copy()
    working_df["ten_hang_clean"] = working_df["Tên hàng"].apply(clean_str)
    working_df = working_df[working_df["ten_hang_clean"] != ""].reset_index(drop=True)
    working_df["_ma_kh"] = working_df["Tên ảnh"].apply(
        lambda img: clean_str(image_customer_map.get(clean_str(img), ""))
    )

    invoice_groups = []
    if group_by_customer:
        seen_keys = []
        grouped_rows = {}
        for _, row in working_df.iterrows():
            ma_kh = row["_ma_kh"]
            key = f"KH::{ma_kh}" if ma_kh else f"IMG::{clean_str(row['Tên ảnh'])}"
            if key not in grouped_rows:
                seen_keys.append(key)
                grouped_rows[key] = []
            grouped_rows[key].append(row)
        for key in seen_keys:
            rows_list = grouped_rows[key]
            invoice_groups.append((rows_list[0]["_ma_kh"], rows_list))
    else:
        seen_imgs = []
        grouped_rows = {}
        for _, row in working_df.iterrows():
            img = clean_str(row["Tên ảnh"])
            if img not in grouped_rows:
                seen_imgs.append(img)
                grouped_rows[img] = []
            grouped_rows[img].append(row)
        for img in seen_imgs:
            rows_list = grouped_rows[img]
            invoice_groups.append((rows_list[0]["_ma_kh"], rows_list))

    preview_rows = []
    current_row = 9
    for inv_idx, (ma_kh, rows_list) in enumerate(invoice_groups, start=1):
        kh_info = kh_lookup.get(ma_kh, {})
        ten_dv = clean_str(kh_info.get("Tên khách hàng", ""))
        mst = clean_str(kh_info.get("MST/CCCD chủ hộ", ""))
        dia_chi = clean_str(kh_info.get("Địa chỉ", ""))
        nguoi_mua = clean_str(kh_info.get("Người nhận HĐ", ""))
        email = clean_str(kh_info.get("Email", ""))
        sdt = clean_str(kh_info.get("Số điện thoại", ""))
        cccd = clean_str(kh_info.get("Căn cước công dân", ""))

        for item_pos, item in enumerate(rows_list):
            if current_row > 9:
                copy_row_style(ws, 9, current_row, max_col=15)

            fill_buyer = repeat_buyer_info or (item_pos == 0)

            # Cột A: Số thứ tự hóa đơn (*)
            c_stt = ws.cell(row=current_row, column=1, value=str(inv_idx))
            c_stt.number_format = "@"

            # Cột B..J: Thông tin chung hóa đơn & người mua
            b_val = invoice_date_str if fill_buyer else ""
            c_val = ten_dv if fill_buyer else ""
            d_val = mst if fill_buyer else ""
            e_val = dia_chi if fill_buyer else ""
            f_val = nguoi_mua if fill_buyer else ""
            g_val = email if fill_buyer else ""
            h_val = sdt if fill_buyer else ""
            i_val = cccd if fill_buyer else ""
            j_val = payment_method if fill_buyer else ""

            text_fields = [
                (2, b_val),
                (3, c_val),
                (4, d_val),
                (5, e_val),
                (6, f_val),
                (7, g_val),
                (8, h_val),
                (9, i_val),
                (10, j_val),
            ]
            for col_i, val_str in text_fields:
                cell = ws.cell(row=current_row, column=col_i, value=val_str if val_str else None)
                cell.hyperlink = None
                cell.number_format = "@"

            # Cột K: Tên hàng hóa/dịch vụ (*)
            ten_hang = clean_str(item["Tên hàng"])
            c_ten = ws.cell(row=current_row, column=11, value=ten_hang)
            c_ten.number_format = "@"

            # Cột L: ĐVT
            dvt = clean_str(item["ĐVT"])
            c_dvt = ws.cell(row=current_row, column=12, value=dvt)
            c_dvt.number_format = "@"

            # Cột M: Số lượng (kiểu số thuần)
            sl_val = clean_excel_num(item["Số lượng"])
            c_sl = ws.cell(row=current_row, column=13, value=sl_val)
            c_sl.number_format = "General"

            # Cột N: Đơn giá (kiểu số thuần)
            dg_val = clean_excel_num(item["Đơn giá"])
            c_dg = ws.cell(row=current_row, column=14, value=dg_val)
            c_dg.number_format = "#,##0"

            # Cột O: Thành tiền (GHI TRỰC TIẾP GIÁ TRỊ SỐ - KHÔNG DÙNG CÔNG THỨC)
            tt_val = int(round(to_number(sl_val, 0.0) * to_number(dg_val, 0.0)))
            c_tt = ws.cell(row=current_row, column=15, value=tt_val)
            c_tt.number_format = "#,##0"

            preview_rows.append(
                {
                    "STT HĐ": str(inv_idx),
                    "Ngày HĐ": b_val,
                    "Tên đơn vị mua hàng (Cột C)": c_val,
                    "MST": d_val,
                    "Địa chỉ": e_val,
                    "HTTT": j_val,
                    "Tên hàng hóa": ten_hang,
                    "ĐVT": dvt,
                    "Số lượng": sl_val,
                    "Đơn giá": dg_val,
                    "Thành tiền (Cột O)": tt_val,
                }
            )
            current_row += 1

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), pd.DataFrame(preview_rows)


# ==============================================================================
# KHỞI TẠO SESSION STATE (TỰ NẠP KHÁCH HÀNG TỪ GITHUB NẾU CÓ)
# ==============================================================================
if "customers_df" not in st.session_state:
    local_kh_path = find_existing_file(DEFAULT_CUSTOMER_PATHS, keyword="khach")
    if local_kh_path:
        try:
            with open(local_kh_path, "rb") as f:
                st.session_state["customers_df"] = load_customers_from_bytes(f.read())
            st.session_state["customers_source_note"] = f"Tự động nạp từ kho GitHub (`{local_kh_path}`)"
        except Exception as e:
            st.warning(f"Không đọc được file khách hàng sẵn có ({local_kh_path}): {e}")
            st.session_state["customers_df"] = pd.DataFrame(columns=CUSTOMER_COLS)
            st.session_state["customers_source_note"] = "Chưa có file khách hàng"
    else:
        st.session_state["customers_df"] = pd.DataFrame(columns=CUSTOMER_COLS)
        st.session_state["customers_source_note"] = "Chưa đặt file khách hàng lên GitHub"

if "last_uploaded_kh_id" not in st.session_state:
    st.session_state["last_uploaded_kh_id"] = None

if "customers_modified" not in st.session_state:
    st.session_state["customers_modified"] = False

if "scanned_df" not in st.session_state:
    st.session_state["scanned_df"] = None

if "image_customer_map" not in st.session_state:
    st.session_state["image_customer_map"] = {}

if "scan_id" not in st.session_state:
    st.session_state["scan_id"] = 0


# ==============================================================================
# BƯỚC 1: TẢI ẢNH & TÙY CHỌN NHANH
# ==============================================================================
st.subheader("1️⃣ Chọn ảnh đơn hàng & thiết lập hóa đơn")

uploaded_images = st.file_uploader(
    "📷 Chọn hoặc chụp nhiều ảnh hóa đơn/đơn hàng trái cây:",
    type=["jpg", "jpeg", "png", "webp"],
    accept_multiple_files=True,
)

col1, col2 = st.columns(2)
with col1:
    invoice_date = st.date_input("📅 Ngày hóa đơn", value=date.today(), format="DD/MM/YYYY")
with col2:
    payment_method = st.selectbox(
        "💳 Hình thức thanh toán (*)",
        options=["TM/CK", "Chuyển khoản", "Tiền mặt"],
        index=0,
    )

group_mode = st.radio(
    "🧾 Quy tắc gom hóa đơn:",
    options=[
        "Gom các ảnh cùng khách hàng thành 1 hóa đơn",
        "Mỗi ảnh là 1 hóa đơn riêng",
    ],
    index=0,
)
group_by_customer = group_mode.startswith("Gom")

repeat_buyer_info = st.checkbox(
    "📋 Điền thông tin đơn vị mua hàng (Cột B–J) ở tất cả các dòng của hóa đơn (khuyên dùng để tránh mất Cột C khi import)",
    value=True,
)

# Kiểm tra trạng thái file mẫu & danh sách khách hàng
local_tpl_path = find_existing_file(DEFAULT_TEMPLATE_PATHS, keyword="mau_hoa_don")
kh_count = len(st.session_state["customers_df"])

if kh_count > 0:
    st.caption(
        f"✅ Mẫu MTT: **Tích hợp sẵn** | "
        f"Danh sách khách hàng: **{kh_count} đơn vị** ({st.session_state['customers_source_note']})"
    )
else:
    st.warning(
        "⚠️ Chưa tìm thấy file `Danh_sach_khach_hang (2).xls` trong kho GitHub. "
        "Bạn có thể tải file lên GitHub 1 lần duy nhất (không cần đính kèm lại) hoặc tải tạm ở mục bên dưới."
    )

with st.expander("📂 Tùy chọn nâng cao: Tải đè file mẫu MTT hoặc file khách hàng"):
    if local_tpl_path:
        st.caption(f"Đang ưu tiên file mẫu trên GitHub: `{local_tpl_path}` (nếu không có sẽ dùng mẫu chuẩn tích hợp sẵn).")
    else:
        st.caption("✅ Đang dùng **file mẫu MTT chuẩn 15 cột tích hợp sẵn trong code** (không cần tải thêm).")
    uploaded_template = st.file_uploader(
        "Tải file mẫu hóa đơn MTT khác (chỉ dùng khi muốn đổi mẫu):",
        type=["xls", "xlsx"],
        key="tpl_uploader",
    )

    uploaded_kh_file = st.file_uploader(
        "Tải file Danh sách khách hàng khác (.xls / .xlsx):",
        type=["xls", "xlsx"],
        key="kh_uploader",
    )
    if uploaded_kh_file is not None:
        kh_file_id = f"{uploaded_kh_file.name}_{uploaded_kh_file.size}"
        if st.session_state["last_uploaded_kh_id"] != kh_file_id:
            try:
                st.session_state["customers_df"] = load_customers_from_bytes(uploaded_kh_file.getvalue())
                st.session_state["last_uploaded_kh_id"] = kh_file_id
                st.session_state["customers_source_note"] = f"Nạp tạm từ `{uploaded_kh_file.name}`"
                # Tự gợi ý lại khách hàng nếu đã quét ảnh trước đó
                if st.session_state["scanned_df"] is not None:
                    new_map = {}
                    for img_name in st.session_state["scanned_df"]["Tên ảnh"].unique():
                        new_map[img_name] = suggest_customer_code(img_name, st.session_state["customers_df"])
                    st.session_state["image_customer_map"] = new_map
                st.success(f"✅ Đã nạp {len(st.session_state['customers_df'])} khách hàng!")
            except Exception as e:
                st.error(f"Lỗi đọc file khách hàng: {e}")


# ==============================================================================
# BƯỚC 2: QUÉT AI THEO LÔ (TỰ ĐỘNG CHUYỂN MODEL DỰ PHÒNG KHI QUÁ TẢI)
# ==============================================================================
if uploaded_images:
    if st.button("🚀 Tiến hành quét AI và chuẩn hóa hàng loạt"):
        status_box = st.empty()
        try:
            client = genai.Client(api_key=API_KEY)
            all_data = []
            used_models = []
            batch_size = 5
            file_chunks = [
                uploaded_images[i : i + batch_size]
                for i in range(0, len(uploaded_images), batch_size)
            ]
            progress_bar = st.progress(0, text="🤖 Đang chuẩn bị quét ảnh...")

            for chunk_idx, chunk in enumerate(file_chunks):
                progress_bar.progress(
                    chunk_idx / len(file_chunks),
                    text=f"🤖 Đang xử lý lô {chunk_idx + 1}/{len(file_chunks)} ({len(chunk)} ảnh)...",
                )
                images = []
                file_names = []
                for f_img in chunk:
                    f_img.seek(0)
                    images.append(Image.open(f_img))
                    file_names.append(f_img.name)

                prompt = f"""
Bạn là trợ lý kế toán chuyên nghiệp cho hộ kinh doanh trái cây. Dưới đây là {len(chunk)} ảnh đơn hàng/hóa đơn.
Tên các file ảnh theo đúng thứ tự là: {file_names}.
Hãy đọc kỹ từng ảnh, tự động sửa lỗi chính tả tiếng Việt do ảnh mờ hoặc chữ viết tay, và trích xuất danh sách từng mặt hàng cho TỪNG ảnh.
Bỏ qua các dòng tổng cộng (Total / Tổng cộng / Thành tiền tổng).

Trả về duy nhất một mảng JSON hợp lệ (không có văn bản giải thích, không có markdown). Mỗi phần tử gồm các trường:
- ten_anh: Tên file ảnh chính xác chứa mặt hàng đó
- ten_hang: Tên mặt hàng tiếng Việt chuẩn xác
- don_vi_tinh: Đơn vị tính ghi trong ảnh (nếu không ghi thì để chuỗi rỗng "")
- so_luong: Số lượng dạng số (float hoặc int)
- don_gia: Đơn giá dạng số (float hoặc int)
- thanh_tien: Thành tiền dạng số (float hoặc int)

Ví dụ cấu trúc:
[{{"ten_anh": "{file_names[0]}", "ten_hang": "Táo đỏ", "don_vi_tinh": "Kg", "so_luong": 2.5, "don_gia": 45000, "thanh_tien": 112500}}]
"""

                response, used_model = generate_with_fallback(
                    client=client,
                    contents=images + [prompt],
                    status_placeholder=status_box,
                )
                if used_model not in used_models:
                    used_models.append(used_model)

                text_res = (response.text or "").strip()
                if text_res.startswith("```json"):
                    text_res = text_res[7:]
                elif text_res.startswith("```"):
                    text_res = text_res[3:]
                if text_res.endswith("```"):
                    text_res = text_res[:-3]
                text_res = text_res.strip()

                items = json.loads(text_res)
                if isinstance(items, list):
                    all_data.extend(items)

            progress_bar.progress(1.0, text="✅ Hoàn tất quét tất cả các lô ảnh!")
            status_box.empty()

            if not all_data:
                st.warning("Không trích xuất được dòng hàng hóa nào từ các ảnh đã tải lên.")
            else:
                df_raw = pd.DataFrame(all_data)
                for col in ["ten_anh", "ten_hang", "don_vi_tinh", "so_luong", "don_gia", "thanh_tien"]:
                    if col not in df_raw.columns:
                        df_raw[col] = ""

                df_raw["Tên ảnh"] = df_raw["ten_anh"].apply(clean_str)
                df_raw["Tên hàng"] = df_raw["ten_hang"].apply(process_item_name)
                df_raw["Số lượng"] = df_raw["so_luong"].apply(lambda x: to_number(x, 0.0))
                df_raw["Đơn giá"] = df_raw["don_gia"].apply(lambda x: to_number(x, 0.0))
                df_raw["ĐVT"] = df_raw.apply(
                    lambda r: process_unit(r["Tên hàng"], r["don_vi_tinh"], r["Đơn giá"]),
                    axis=1,
                )
                df_raw["Thành tiền"] = (df_raw["Số lượng"] * df_raw["Đơn giá"]).round(0)

                df_clean = df_raw[df_raw["Tên hàng"] != ""][
                    ["Tên ảnh", "Tên hàng", "ĐVT", "Số lượng", "Đơn giá", "Thành tiền"]
                ].reset_index(drop=True)

                st.session_state["scan_id"] += 1
                st.session_state["scanned_df"] = df_clean

                df_kh_now = st.session_state["customers_df"]
                new_map = {}
                for img_name in df_clean["Tên ảnh"].unique():
                    new_map[img_name] = suggest_customer_code(img_name, df_kh_now)
                st.session_state["image_customer_map"] = new_map

                model_note = f" (Model: {', '.join(used_models)})" if used_models else ""
                st.success(
                    f"🎉 Đã quét xong {len(df_clean)} dòng hàng từ "
                    f"{df_clean['Tên ảnh'].nunique()} ảnh!{model_note}"
                )

        except Exception as e:
            status_box.empty()
            err_msg = str(e)
            if "429" in err_msg or "RESOURCE_EXHAUSTED" in err_msg:
                st.error(
                    "⚠️ Cả 3 model đều chạm giới hạn lượt gọi miễn phí (Lỗi 429). "
                    "Vui lòng chờ 1-2 phút rồi thử lại hoặc đổi API Key khác."
                )
            elif "503" in err_msg or "UNAVAILABLE" in err_msg:
                st.error(
                    "⚠️ Hệ thống AI Google (cả model chính và dự phòng) đều đang quá tải tạm thời (Lỗi 503). "
                    "Bạn đợi khoảng 1-2 phút rồi bấm quét lại nhé!"
                )
            elif "404" in err_msg or "NOT_FOUND" in err_msg:
                st.error(
                    "⚠️ Không tìm thấy model phù hợp (Lỗi 404). "
                    "Mở mục '⚙️ Công cụ kỹ thuật' ở cuối trang để kiểm tra danh sách model khả dụng."
                )
            else:
                st.error(f"Đã có lỗi xảy ra: {err_msg}")


# ==============================================================================
# BƯỚC 3: GÁN KHÁCH HÀNG (ĐÃ TỰ GỢI Ý) + THÊM KHÁCH MỚI + RÀ SOÁT HÀNG HÓA
# ==============================================================================
if st.session_state["scanned_df"] is not None and not st.session_state["scanned_df"].empty:
    st.markdown("---")
    st.subheader("2️⃣ Gán đơn vị mua hàng theo từng ảnh")

    df_kh = st.session_state["customers_df"]
    kh_codes = [""] + df_kh["Mã khách hàng"].tolist()
    kh_labels = {"": "-- Chưa chọn (Cột C sẽ trống) --"}
    kh_info_map = {}
    for _, r_kh in df_kh.iterrows():
        code = clean_str(r_kh["Mã khách hàng"])
        name = clean_str(r_kh["Tên khách hàng"])
        kh_labels[code] = f"[{code}] {name}"
        kh_info_map[code] = r_kh.to_dict()

    unique_imgs = list(dict.fromkeys(st.session_state["scanned_df"]["Tên ảnh"].tolist()))
    missing_kh_imgs = []

    for img_idx, img_name in enumerate(unique_imgs):
        current_code = st.session_state["image_customer_map"].get(img_name, "")
        if current_code not in kh_codes:
            current_code = ""
        default_idx = kh_codes.index(current_code)

        selected_code = st.selectbox(
            f"🖼️ Ảnh `{img_name}`",
            options=kh_codes,
            index=default_idx,
            format_func=lambda c: kh_labels.get(c, c),
            key=f"sel_kh_{st.session_state['scan_id']}_{img_idx}_{img_name}",
        )
        st.session_state["image_customer_map"][img_name] = selected_code

        if selected_code and selected_code in kh_info_map:
            info = kh_info_map[selected_code]
            details = [f"**Tên ĐV (Cột C):** {info['Tên khách hàng']}"]
            if clean_str(info.get("MST/CCCD chủ hộ")):
                details.append(f"**MST:** {info['MST/CCCD chủ hộ']}")
            if clean_str(info.get("Địa chỉ")):
                details.append(f"**ĐC:** {info['Địa chỉ']}")
            st.caption(" | ".join(details))
        else:
            missing_kh_imgs.append(img_name)

    if missing_kh_imgs:
        st.warning(
            f"⚠️ Có **{len(missing_kh_imgs)} ảnh** chưa chọn đơn vị mua hàng (`{', '.join(missing_kh_imgs)}`). "
            "Nếu không chọn, Cột C (Tên đơn vị mua hàng) của các ảnh này sẽ để trống."
        )

    with st.expander("➕ Thêm khách hàng mới vào danh sách"):
        with st.form("form_add_customer", clear_on_submit=True):
            new_ten = st.text_input("Tên đơn vị mua hàng (*)")
            new_ma = st.text_input("Mã khách hàng (để trống sẽ tự tạo từ tên)")
            c_a, c_b = st.columns(2)
            with c_a:
                new_mst = st.text_input("Mã số thuế / CCCD chủ hộ")
                new_ng_mua = st.text_input("Người mua hàng / Nhận HĐ")
                new_sdt = st.text_input("Số điện thoại")
            with c_b:
                new_cccd = st.text_input("Căn cước công dân")
                new_email = st.text_input("Email")
                new_msdv = st.text_input("MSĐVCQHVNS (nếu có)")
            new_diachi = st.text_input("Địa chỉ")
            submitted = st.form_submit_button("💾 Lưu khách hàng mới")

            if submitted:
                if not new_ten.strip():
                    st.error("Vui lòng nhập Tên đơn vị mua hàng!")
                else:
                    auto_code = new_ma.strip().upper()
                    if not auto_code:
                        auto_code = re.sub(
                            r"[^A-Z0-9]", "", strip_accents(new_ten).upper()
                        )[:16]
                    existing_codes = set(st.session_state["customers_df"]["Mã khách hàng"].tolist())
                    base_code = auto_code or "KHMOI"
                    suffix = 1
                    while auto_code in existing_codes:
                        auto_code = f"{base_code}{suffix}"
                        suffix += 1

                    new_row = {
                        "STT": str(len(st.session_state["customers_df"]) + 1),
                        "Mã khách hàng": auto_code,
                        "Tên khách hàng": new_ten.strip(),
                        "MST/CCCD chủ hộ": new_mst.strip(),
                        "MSĐVCQHVNS": new_msdv.strip(),
                        "Căn cước công dân": new_cccd.strip(),
                        "Địa chỉ": new_diachi.strip(),
                        "Người nhận HĐ": new_ng_mua.strip(),
                        "Email": new_email.strip(),
                        "Số điện thoại": new_sdt.strip(),
                        "Ngừng sử dụng": "",
                    }
                    st.session_state["customers_df"] = pd.concat(
                        [st.session_state["customers_df"], pd.DataFrame([new_row])],
                        ignore_index=True,
                    )
                    st.session_state["customers_modified"] = True

                    xlsx_kh_bytes = export_customers_to_xlsx_bytes(st.session_state["customers_df"])
                    try:
                        with open("Danh_sach_khach_hang.xlsx", "wb") as f_out:
                            f_out.write(xlsx_kh_bytes)
                    except Exception:
                        pass

                    ok_gh, msg_gh = save_customers_to_github(xlsx_kh_bytes, "Danh_sach_khach_hang.xlsx")
                    if ok_gh:
                        st.success(f"✅ Đã thêm [{auto_code}] {new_ten.strip()}. {msg_gh}")
                    else:
                        st.success(f"✅ Đã thêm [{auto_code}] {new_ten.strip()} vào phiên làm việc hiện tại!")
                    st.rerun()

    if st.session_state["customers_modified"]:
        st.download_button(
            label="💾 Tải danh sách khách hàng mới nhất (.xlsx)",
            data=export_customers_to_xlsx_bytes(st.session_state["customers_df"]),
            file_name="Danh_sach_khach_hang.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.markdown("---")
    st.subheader("3️⃣ Rà soát & chỉnh sửa nhanh chi tiết hàng hóa")
    st.caption("Chạm trực tiếp vào ô để sửa tên hàng, ĐVT, số lượng, đơn giá trước khi xuất hóa đơn MTT.")

    edited_df = st.data_editor(
        st.session_state["scanned_df"],
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Tên ảnh": st.column_config.TextColumn("Tên ảnh"),
            "Tên hàng": st.column_config.TextColumn("Tên hàng hóa (*)"),
            "ĐVT": st.column_config.TextColumn("ĐVT"),
            "Số lượng": st.column_config.NumberColumn("Số lượng", format="%.2f"),
            "Đơn giá": st.column_config.NumberColumn("Đơn giá", format="%,.0f"),
            "Thành tiền": st.column_config.NumberColumn("Thành tiền", format="%,.0f", disabled=True),
        },
        key=f"editor_items_{st.session_state['scan_id']}",
    )

    calc_df = edited_df.copy()
    calc_df["Số lượng"] = calc_df["Số lượng"].apply(lambda x: to_number(x, 0.0))
    calc_df["Đơn giá"] = calc_df["Đơn giá"].apply(lambda x: to_number(x, 0.0))
    calc_df["Thành tiền"] = (calc_df["Số lượng"] * calc_df["Đơn giá"]).round(0)

    # ==========================================================================
    # BƯỚC 4: XEM TRƯỚC & XUẤT FILE EXCEL MẪU HÓA ĐƠN MTT
    # ==========================================================================
    st.markdown("---")
    st.subheader("4️⃣ Kiểm tra trước & xuất file Excel MTT")

    if uploaded_template is not None:
        template_bytes = uploaded_template.getvalue()
    elif local_tpl_path is not None:
        with open(local_tpl_path, "rb") as f_tpl:
            template_bytes = f_tpl.read()
    else:
        template_bytes = create_builtin_mtt_template()

    try:
        mtt_bytes, preview_df = build_mtt_excel(
            template_bytes=template_bytes,
            items_df=calc_df,
            image_customer_map=st.session_state["image_customer_map"],
            df_kh=st.session_state["customers_df"],
            invoice_date_str=invoice_date.strftime("%d/%m/%Y"),
            payment_method=payment_method,
            group_by_customer=group_by_customer,
            repeat_buyer_info=repeat_buyer_info,
        )

        so_hd = preview_df["STT HĐ"].nunique() if not preview_df.empty else 0
        tong_tien = preview_df["Thành tiền (Cột O)"].sum() if not preview_df.empty else 0
        st.info(
            f"📊 Dự kiến xuất **{so_hd} hóa đơn** | **{len(preview_df)} dòng hàng** | "
            f"Tổng tiền: **{tong_tien:,.0f} đ** | Thanh toán: **{payment_method}**"
        )

        with st.expander("👁️ Xem trước dữ liệu chuẩn MTT (kiểm tra Cột C & Cột O trước khi tải)", expanded=True):
            st.dataframe(preview_df, use_container_width=True, hide_index=True)

        out_filename = f"Hoa_don_MTT_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        st.download_button(
            label="📥 Tải file Excel mẫu hóa đơn MTT (.xlsx)",
            data=mtt_bytes,
            file_name=out_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"Lỗi khi tạo file Excel mẫu MTT: {e}")

elif not uploaded_images:
    st.info("💡 Chọn hoặc chụp một hay nhiều ảnh đơn hàng ở phía trên để bắt đầu.")


# ==============================================================================
# CÔNG CỤ KỸ THUẬT THU GỌN Ở CUỐI TRANG
# ==============================================================================
st.markdown("---")
with st.expander("⚙️ Công cụ kỹ thuật (Kiểm tra Model Gemini)"):
    if st.button("🔍 Dò tìm tên Model chuẩn"):
        try:
            client = genai.Client(api_key=API_KEY)
            models = client.models.list()
            st.success("Các Model tài khoản của bạn được phép dùng:")
            for m in models:
                st.write(f"- `{m.name}`")
        except Exception as e:
            st.error(f"Lỗi khi dò tìm: {e}")
