import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
import io
import math

def convert_to_roman(num):
    val = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
    syb = ["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"]
    roman_num = ''
    i = 0
    while num > 0:
        for _ in range(num // val[i]):
            roman_num += syb[i]
            num -= val[i]
        i += 1
    return roman_num

def safe_str(val):
    if pd.isna(val):
        return ""
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return str(val).strip()

def process_files(excel_file, word_template):
    # Đọc dữ liệu Excel/CSV (bỏ qua 2 dòng tiêu đề đầu tiên)
    if excel_file.name.endswith('.csv'):
        df = pd.read_csv(excel_file, skiprows=2, header=None)
    else:
        df = pd.read_excel(excel_file, skiprows=2, header=None)
        
    # Loại bỏ các dòng không có tên
    df = df.dropna(subset=[1])

    # Mở file Word mẫu
    doc = Document(word_template)
    table = doc.tables[0] # Lấy bảng đầu tiên trong file Word
    
    # Xóa các dòng dữ liệu cũ trong file mẫu (chỉ giữ lại 1 dòng tiêu đề đầu tiên)
    for row in table.rows[1:]:
        table._element.remove(row._element)

    # Nhóm dữ liệu theo Tỉnh (Cột index 13) và Xã (Cột index 12)
    grouped_tinh = df.groupby(13)
    
    tinh_counter = 1
    for tinh, group_tinh in grouped_tinh:
        tinh_str = safe_str(tinh)
        if not tinh_str: continue
        
        # --- Thêm dòng Tỉnh ---
        row_tinh = table.add_row()
        # Gộp tất cả các ô trong dòng thành 1 ô
        row_tinh.cells[0].merge(row_tinh.cells[-1])
        cell_tinh = row_tinh.cells[0]
        cell_tinh.text = f"{convert_to_roman(tinh_counter)}. Tỉnh {tinh_str}"
        # In đậm dòng Tỉnh
        if cell_tinh.paragraphs:
            cell_tinh.paragraphs[0].runs[0].bold = True
            cell_tinh.paragraphs[0].runs[0].font.name = 'Times New Roman'
        tinh_counter += 1
        
        grouped_xa = group_tinh.groupby(12)
        xa_counter = 1
        for xa, group_xa in grouped_xa:
            xa_str = safe_str(xa)
            if not xa_str: continue
            
            # --- Thêm dòng Xã ---
            row_xa = table.add_row()
            row_xa.cells[0].merge(row_xa.cells[-1])
            cell_xa = row_xa.cells[0]
            cell_xa.text = f"{xa_counter}. Xã {xa_str}"
            if cell_xa.paragraphs:
                cell_xa.paragraphs[0].runs[0].bold = True
                cell_xa.paragraphs[0].runs[0].font.name = 'Times New Roman'
            xa_counter += 1
            
            # --- Thêm chi tiết từng người ---
            for _, row_data in group_xa.iterrows():
                row_cells = table.add_row().cells
                
                # Trích xuất dữ liệu
                ho_ten = safe_str(row_data[1])
                dv = safe_str(row_data[2])
                
                # Chuẩn hóa Cấp bậc, Chức vụ (VD: Binh nhì -> B2, Chiến sĩ -> CS)
                cb = safe_str(row_data[3]).replace("Binh nhì", "B2") 
                cv = safe_str(row_data[4]).replace("Chiến sĩ", "CS")
                
                nn = safe_str(row_data[5])
                dt = safe_str(row_data[6])
                vh = safe_str(row_data[7])
                cccd = safe_str(row_data[8])
                
                # Cắt số thập phân (.0) nếu có ở Ngày, Tháng, Năm
                ngay = safe_str(row_data[9])
                thang = safe_str(row_data[10])
                nam = safe_str(row_data[11])
                ngay_sinh = f"{ngay}/{thang}/{nam}" if ngay and nam else ""
                
                bo = safe_str(row_data[16])
                me = safe_str(row_data[17])
                bo_me = f"{bo}\n{me}".strip() # Xuống dòng giữa tên bố và mẹ
                
                sdt = safe_str(row_data[18])
                que_quan = f"{xa_str}-{tinh_str}"

                # Ghi vào 21 cột tương ứng của file Word mẫu
                row_cells[0].text = "" # Cột TT để trống
                row_cells[1].text = ho_ten
                row_cells[2].text = ngay_sinh
                row_cells[3].text = cb
                row_cells[4].text = cv
                row_cells[5].text = dv
                row_cells[6].text = nn
                row_cells[7].text = "BN" # Thành phần
                row_cells[8].text = vh
                row_cells[9].text = "" # Sức khỏe
                row_cells[10].text = dt
                row_cells[11].text = "" # Tôn giáo
                row_cells[12].text = "" # Ngày vào đoàn
                row_cells[13].text = "" # Ngày vào Đảng
                row_cells[14].text = que_quan
                row_cells[15].text = "" # Trú quán
                row_cells[16].text = bo_me # Họ tên bố mẹ
                row_cells[17].text = "" # Họ tên vợ con
                row_cells[18].text = sdt # SĐT
                row_cells[19].text = cccd
                row_cells[20].text = "" # Ghi chú
                
                # Set font Times New Roman và size 11 cho đồng bộ với bảng
                for cell in row_cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.name = 'Times New Roman'
                            run.font.size = Pt(11)

    # Lưu kết quả
    byte_io = io.BytesIO()
    doc.save(byte_io)
    byte_io.seek(0)
    return byte_io

# --- GIAO DIỆN STREAMLIT ---
st.set_page_config(page_title="Công cụ điền dữ liệu Trích Ngang", layout="wide")
st.title("Tạo danh sách Trích Ngang từ file Word mẫu")

st.markdown("""
**Hướng dẫn sử dụng:**
1. Tải lên **File Word mẫu** của bạn (file `Tổng hợp trích ngang.docx`).
2. Tải lên **File Excel/CSV** chứa danh sách quân nhân.
3. Bấm nút tạo và phần mềm sẽ điền tự động vào bảng của file Word mẫu.
""")

col1, col2 = st.columns(2)
with col1:
    word_file = st.file_uploader("1. Tải lên File Word Mẫu (.docx)", type=['docx'])
with col2:
    excel_file = st.file_uploader("2. Tải lên File Excel dữ liệu (.xlsx, .csv)", type=['xlsx', 'csv'])

if word_file and excel_file:
    if st.button("🚀 Chạy và Tạo File Word Cuối Cùng", use_container_width=True):
        with st.spinner("Đang xử lý dữ liệu và chèn vào biểu mẫu..."):
            try:
                result_doc = process_files(excel_file, word_file)
                st.success("✅ Tạo file thành công!")
                st.download_button(
                    label="📥 TẢI XUỐNG FILE KẾT QUẢ",
                    data=result_doc,
                    file_name="Ket_qua_Trich_ngang_Hoan_chinh.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Đã xảy ra lỗi: {e}")
