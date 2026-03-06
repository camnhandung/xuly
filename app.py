import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
import io

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

def safe_get(row, index, default=""):
    try:
        val = row[index]
        if pd.isna(val): return default
        if isinstance(val, float) and val.is_integer(): return str(int(val))
        return str(val).strip()
    except:
        return default

def process_files(excel_file, word_template):
    # Đọc dữ liệu, bỏ qua 2 dòng đầu (tiêu đề)
    if excel_file.name.endswith('.csv'):
        df = pd.read_csv(excel_file, skiprows=2, header=None)
    else:
        df = pd.read_excel(excel_file, skiprows=2, header=None)
        
    # Lọc bỏ dòng trống
    df = df.dropna(subset=[1])

    doc = Document(word_template)
    table = doc.tables[0]
    
    # Xóa dữ liệu cũ trong bảng mẫu (giữ dòng header)
    for row in table.rows[1:]:
        table._element.remove(row._element)

    # Nhóm theo Tỉnh (Cột 13) và Xã (Cột 12)
    # Lưu ý: Nếu file của bạn ít cột hơn, hãy kiểm tra lại số 13 và 12
    grouped_tinh = df.groupby(13)
    
    tinh_idx = 1
    for tinh, group_tinh in grouped_tinh:
        tinh_name = str(tinh).strip()
        if not tinh_name or tinh_name == "nan": continue
        
        # Dòng Tỉnh
        row_tinh = table.add_row()
        row_tinh.cells[0].merge(row_tinh.cells[-1])
        row_tinh.cells[0].text = f"{convert_to_roman(tinh_idx)}. Tỉnh {tinh_name}"
        row_tinh.cells[0].paragraphs[0].runs[0].bold = True
        tinh_idx += 1
        
        grouped_xa = group_tinh.groupby(12)
        xa_idx = 1
        for xa, group_xa in grouped_xa:
            xa_name = str(xa).strip()
            if not xa_name or xa_name == "nan": continue
            
            # Dòng Xã
            row_xa = table.add_row()
            row_xa.cells[0].merge(row_xa.cells[-1])
            row_xa.cells[0].text = f"{xa_idx}. Xã {xa_name}"
            row_xa.cells[0].paragraphs[0].runs[0].bold = True
            xa_idx += 1
            
            # Dòng dữ liệu cá nhân
            for _, row in group_xa.iterrows():
                new_row = table.add_row().cells
                
                # Lấy dữ liệu an toàn bằng hàm safe_get để tránh lỗi Index
                ho_ten = safe_get(row, 1)
                dv = safe_get(row, 2)
                cb = safe_get(row, 3).replace("Binh nhì", "B2")
                cv = safe_get(row, 4).replace("Chiến sĩ", "CS")
                nn = safe_get(row, 5)
                dt = safe_get(row, 6)
                vh = safe_get(row, 7)
                cccd = safe_get(row, 8)
                ngay_sinh = f"{safe_get(row,9)}/{safe_get(row,10)}/{safe_get(row,11)}"
                bo_me = f"{safe_get(row,16)}\n{safe_get(row,17)}".strip()
                sdt = safe_get(row, 18)
                que = f"{xa_name}-{tinh_name}"

                # Điền vào bảng (đảm bảo file Word có đủ 21 cột như mẫu)
                mapping = {
                    1: ho_ten, 2: ngay_sinh, 3: cb, 4: cv, 5: dv,
                    6: nn, 7: "BN", 8: vh, 10: dt, 14: que,
                    16: bo_me, 18: sdt, 19: cccd
                }
                
                for col_i, text in mapping.items():
                    if col_i < len(new_row):
                        new_row[col_i].text = text
                        # Định dạng font
                        for p in new_row[col_i].paragraphs:
                            for run in p.runs:
                                run.font.name = 'Times New Roman'
                                run.font.size = Pt(10)

    byte_io = io.BytesIO()
    doc.save(byte_io)
    byte_io.seek(0)
    return byte_io

st.title("Fix lỗi: Trích ngang Tỉnh/Xã")
w_file = st.file_uploader("Tải Word mẫu", type=['docx'])
e_file = st.file_uploader("Tải Excel dữ liệu", type=['xlsx', 'csv'])

if w_file and e_file:
    if st.button("Bắt đầu xử lý"):
        try:
            out = process_files(e_file, w_file)
            st.download_button("Tải file kết quả", out, "Ket_qua.docx")
        except Exception as e:
            st.error(f"Lỗi chi tiết: {e}")
