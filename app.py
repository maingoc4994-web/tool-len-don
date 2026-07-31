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
    "<p style='text-align: center; color: #555;'>Sử dụng AI Google Gemini đọc ảnh theo lô tối ưu, tự động phân loại và chuẩn hóa dữ liệu hàng loạt.</p>",
    unsafe_allow_html=True,
)

# Lấy API Key bảo mật từ Streamlit Secrets (An toàn tuyệt đối)
try:
  API_KEY = st.secrets["GEMINI_API_KEY"]
except Exception:
  st.error(
      "⚠️ Chưa cấu hình GEMINI_API_KEY trong mục Secrets của Streamlit Cloud!"
  )
  st.stop()

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
  - Xóa phần '/...' phía sau nếu là chữ tiếng Anh.
  - Viết hoa chữ cái đầu mỗi từ, sửa lỗi chính tả chuẩn hóa.
  """
  if not name:
    return ""

  # Bỏ chữ Total
  name = re.sub(r"(?i)\btotal\b", "", name).strip()

  # Xóa phần từ dấu '/' trở đi nếu phần sau chứa chữ tiếng Anh
  name = re.sub(r"\s*/\s*[a-zA-Z].*$", "", name).strip()

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
  if st.button("🚀 Tiến hành quét AI và chuẩn hóa hàng loạt"):
    with st.spinner(
        "🤖 AI đang xử lý theo lô ảnh để tối ưu hạn mức và tốc độ..."
    ):
      try:
        client = genai.Client(api_key=API_KEY)
        all_data = []

        # Chia danh sách ảnh thành các nhóm (lô) tối đa 5 ảnh mỗi lần gọi để tránh quá tải
        batch_size = 5
        file_chunks = [
            uploaded_files[i : i + batch_size]
            for i in range(0, len(uploaded_files), batch_size)
        ]

        for chunk in file_chunks:
          images = []
          file_names = []
          for uploaded_file in chunk:
            images.append(Image.open(uploaded_file))
            file_names.append(uploaded_file.name)

          prompt = f"""
                    Bạn là một trợ lý kế toán chuyên nghiệp. Dưới đây là {len(chunk)} ảnh hóa đơn/danh sách hàng hóa. Tên các file ảnh lần lượt theo đúng thứ tự là: {file_names}.
                    Hãy đọc nội dung từng ảnh, tự động sửa mọi lỗi chính tả do hình ảnh mờ hoặc viết tay, và trích xuất danh sách các mặt hàng cho TỪNG ảnh tương ứng.
                    Trả về định dạng JSON chuẩn là một mảng các đối tượng, trong đó mỗi đối tượng bắt buộc phải có trường "ten_anh" ứng với tên file ảnh chính xác chứa mặt hàng đó, cùng các trường sau:
                    - ten_anh: Tên file ảnh tương ứng của mặt hàng
                    - ten_hang: Tên hàng hóa (tự động sửa chính tả tiếng Việt cho chuẩn xác)
                    - don_vi_tinh: Đơn vị tính có trong ảnh (nếu không có thì để trống chuỗi "")
                    - so_luong: Số lượng (dạng số)
                    - don_gia: Đơn giá (dạng số)
                    - thanh_tien: Thành tiền (dạng số)
                    
                    Chỉ trả về cấu trúc JSON thuần túy, tuyệt đối không kèm Markdown hay văn bản giải thích nào khác.
                    Cấu trúc mẫu:
                    [{{"ten_anh": "{file_names[0]}", "ten_hang": "...", "don_vi_tinh": "...", "so_luong": 0, "don_gia": 0, "thanh_tien": 0}}]
                    """

          response = client.models.generate_content(
              model="gemini-3-flash-preview", contents=images + [prompt]
          )

          text_res = response.text.strip()
          if text_res.startswith("```json"):
            text_res = text_res[7:-3].strip()
          elif text_res.startswith("```"):
            text_res = text_res[3:-3].strip()

          items = json.loads(text_res)
          for item in items:
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

        st.success(
            "🎉 Quét hàng loạt ảnh theo lô và chuẩn hóa thành công!"
        )
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
