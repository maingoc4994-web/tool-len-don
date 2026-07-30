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
    </style>
""",
    unsafe_allow_html=True,
)

st.markdown("<h1>🛒 Tool lên đơn (AI Thông Minh)</h1>", unsafe_allow_html=True)
st.write(
    "<p style='text-align: center; color: #555;'>Sử dụng AI Google Gemini để đọc ảnh cực chuẩn, tự động phân loại theo từng ảnh và chuẩn hóa dữ liệu.</p>",
    unsafe_allow_html=True,
)

# Tích hợp sẵn API Key trực tiếp vào code
API_KEY = "AIzaSyDl1p0krbXrDmnzFIyd0oSAHVjowM3KtEg"

# Cho phép chọn nhiều ảnh một lúc
uploaded_files = st.file_uploader(
    "📷 Tải lên hoặc chọn nhiều ảnh hóa đơn/danh sách cùng lúc:",
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

  return unit.lower().capitalize()


if uploaded_files:
  if st.button("🚀 Tiến hành quét AI và chuẩn hóa"):
    with st.spinner("🤖 AI đang đọc toàn bộ ảnh, tự động phân nhóm và chuẩn hóa..."):
      try:
        client = genai.Client(api_key=API_KEY)
        all_data = []

        # Lần lượt đọc từng ảnh trong danh sách các ảnh đã chọn
        for uploaded_file in uploaded_files:
          image = Image.open(uploaded_file)
          image_name = uploaded_file.name  # Lấy tên file ảnh để phân biệt

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

          # Gắn thêm tên ảnh vào từng dòng dữ liệu để phân biệt
          for item in items:
            item["ten_anh"] = image_name
            all_data.append(item)

        df = pd.DataFrame(all_data)

        df["ten_hang"] = df["ten_hang"].apply(process_item_name)
        df["don_vi_tinh"] = df.apply(process_unit, axis=1)

        df["so_luong"] = df["so_luong"].apply(format_number)
        df["don_gia"] = df["don_gia"].apply(format_number)
        df["thanh_tien"] = df["thanh_tien"].apply(format_number)

        # Sắp xếp lại thứ tự cột, đưa Tên ảnh lên đầu tiên
        df = df[
            [
                "ten_anh",
                "ten_hang",
                "don_vi_tinh",
                "so_luong",
                "don_gia",
                "thanh_tien",
            ]
        ]
        df.columns = [
            "Tên ảnh",
            "Tên hàng",
            "Đơn vị tính",
            "Số lượng",
            "Đơn giá",
            "Thành tiền",
        ]

        st.success("🎉 Quét hàng loạt ảnh và chuẩn hóa thành công!")
        st.dataframe(df, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
          df.to_excel(writer, index=False, sheet_name="ToolLenDon")
        excel_data = output.getvalue()

        st.download_button(
            label="📥 Tải xuống file Excel tổng hợp (.xlsx)",
            data=excel_data,
            file_name="tool_len_don_tong_hop.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

      except Exception as e:
        st.error(f"Đã có lỗi xảy ra: {e}")

else:
  st.info("💡 Vui lòng chọn một hoặc nhiều ảnh hóa đơn để bắt đầu.")
