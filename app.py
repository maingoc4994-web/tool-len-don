import io
import json
import re
import pandas as pd
from PIL import Image
import streamlit as st
from google import genai

# Cấu hình giao diện ứng dụng
st.set_page_config(
    page_title="Tool lên đơn",
    page_icon="🛒",
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
    h1 { color: #1b5e20; text-align: center; font-size: 1.8rem !important; }
    .stDownloadButton>button {
        width: 100%;
        background-color: #0277bd;
        color: white;
        font-weight: bold;
        border-radius: 10px;
        padding: 10px;
        font-size: 16px;
    }
    .stTextInput>div>div>input { border-radius: 8px; }
    </style>
""",
    unsafe_allow_html=True,
)

st.markdown("<h1>🛒 Tool lên đơn (AI Thông Minh)</h1>", unsafe_allow_html=True)
st.write(
    "<p style='text-align: center; color: #555;'>Sử dụng AI Google Gemini để đọc ảnh cực chuẩn, tự động sửa lỗi chính tả và chuẩn hóa dữ liệu.</p>",
    unsafe_allow_html=True,
)

# Nhập API Key của Gemini
api_key = st.text_input(
    "🔑 Nhập Google Gemini API Key của bạn:",
    type="password",
    placeholder="Dán API Key vào đây...",
)

# Chọn file ảnh (Hỗ trợ chụp trực tiếp trên điện thoại/máy tính)
uploaded_files = st.file_uploader(
    "📷 Tải lên hoặc chụp ảnh hóa đơn/danh sách (chọn 1 hoặc nhiều file):",
    type=["jpg", "jpeg", "png", "webp"],
    accept_multiple_files=True,
)


def format_number(val):
  """Định dạng số: dấu chấm (.) ngăn cách hàng nghìn, dấu phẩy (,) thập phân."""
  try:
    val = float(val)
    s = f"{val:,.2f}"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    if s.endswith(",00"):
      s = s[:-3]
    return s
  except:
    return val


def process_item_name(name):
  """Quy tắc Tên hàng:

  - Bỏ hoàn toàn chữ 'Total' (không phân biệt hoa thường).
  - Viết hoa chữ cái đầu mỗi từ, sửa lỗi chính tả chuẩn hóa.
  """
  if not name:
    return ""
  name = re.sub(r"(?i)\btotal\b", "", name).strip()
  name = re.sub(r"\s+", " ", name)

  words = name.split()
  processed_words = [w.lower().capitalize() for w in words]
  return " ".join(processed_words)


def process_unit(row):
  """Quy tắc Đơn vị tính:

  - Chuối / Chuối ba lùn có đơn giá dưới 8.000đ -> 'Trái'.
  - Không có -> Mặc định 'Kg'.
  - Có sẵn trong ảnh -> Viết thường, viết hoa chữ cái đầu.
  """
  unit = str(row.get("don_vi_tinh", "")).strip()
  name = str(row.get("ten_hang", "")).lower()
  price_raw = row.get("don_gia", 0)

  try:
    if isinstance(price_raw, str):
      price = float(price_raw.replace(".", "").replace(",", "."))
    else:
      price = float(price_raw)
  except:
    price = 0

  if "chuối" in name and 0 < price < 8000:
    return "Trái"

  if not unit or unit.lower() == "nan" or unit == "":
    return "Kg"

  # Viết thường và viết hoa chữ cái đầu cho đơn vị tính
  return unit.lower().capitalize()


if uploaded_files and api_key:
  if st.button("🚀 Tiến hành quét AI và chuẩn hóa"):
    with st.spinner("🤖 AI đang đọc ảnh, tự động sửa lỗi chính tả..."):
      try:
        client = genai.Client(api_key=api_key)
        all_data = []

        for uploaded_file in uploaded_files:
          image = Image.open(uploaded_file)

          prompt = """
                    Bạn là một trợ lý kế toán chuyên nghiệp. Hãy đọc toàn bộ nội dung ảnh hóa đơn/danh sách hàng hóa này, tự động sửa mọi lỗi chính tả do hình ảnh mờ hoặc viết tay, và trích xuất danh sách các mặt hàng thành định dạng JSON chuẩn gồm các trường sau cho mỗi dòng:
                    - ten_hang: Tên hàng hóa (tự động sửa chính tả tiếng Việt cho chuẩn xác)
                    - don_vi_tinh: Đơn vị tính có trong ảnh (nếu không có thì để trống chuỗi "")
                    - so_luong: Số lượng (dạng số)
                    - don_gia: Đơn giá (dạng số)
                    - thanh_tien: Thành tiền (dạng số)
                    
                    Chỉ trả về cấu trúc JSON thuần túy, tuyệt đối không kèm Markdown hay văn bản giải thích nào khác.
                    Cấu trúc mẫu:
                    [{"ten_hang": "...", "don_vi_tinh": "...", "so_luong": 0, "don_gia": 0, "thanh_tien": 0}]
                    """

          response = client.models.generate_content(
              model="gemini-3-flash-preview", contents=[image, prompt]
          )

          text_res = response.text.strip()
          if text_res.startswith("```json"):
            text_res = text_res[7:-3].strip()
          elif text_res.startswith("```"):
            text_res = text_res[3:-3].strip()

          items = json.loads(text_res)
          all_data.extend(items)

        df = pd.DataFrame(all_data)

        df["ten_hang"] = df["ten_hang"].apply(process_item_name)
        df["don_vi_tinh"] = df.apply(process_unit, axis=1)

        df["so_luong"] = df["so_luong"].apply(format_number)
        df["don_gia"] = df["don_gia"].apply(format_number)
        df["thanh_tien"] = df["thanh_tien"].apply(format_number)

        df.columns = [
            "Tên hàng",
            "Đơn vị tính",
            "Số lượng",
            "Đơn giá",
            "Thành tiền",
        ]

        st.success("🎉 Quét ảnh và chuẩn hóa thành công!")
        st.dataframe(df, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
          df.to_excel(writer, index=False, sheet_name="ToolLenDon")
        excel_data = output.getvalue()

        st.download_button(
            label="📥 Tải xuống file Excel (.xlsx)",
            data=excel_data,
            file_name="tool_len_don_ai.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

      except Exception as e:
        st.error(f"Đã có lỗi xảy ra: {e}")

elif not api_key and uploaded_files:
  st.warning("⚠️ Vui lòng nhập Google Gemini API Key ở ô phía trên để tiếp tục.")