import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml import parse_xml
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
        if pd.isna(val) or str(val).lower() == "nan": return default
        if isinstance(val, float) and val.is_integer(): return str(int(val))
        return str(val).strip()
    except:
        return default

def copy_row_format(source_row, target_row):
    """Sao chép định dạng từ dòng nguồn sang dòng đích ở mức độ XML"""
    for i, cell in enumerate(source_row.cells):
        if i < len(target_row.cells):
            # Copy định dạng ô (độ rộng, viền...)
            target_row.cells[i]._tc.set(parse_xml(source_row.cells[i]._tc.xml)[0].attrib, None)
            # Copy định dạng paragraph (căn lề)
            target_row.cells[i].paragraphs[0].alignment = source_row.cells[i].paragraphs[0].alignment

def process_files(excel_file, word_template):
    # 1. Đọc dữ liệu Excel
    if excel_file.name.endswith('.csv'):
        df = pd.read_csv(excel_file, skiprows=2, header=None)
    else:
        df = pd.read_excel(excel_file, skiprows=2, header=None)
    
    # Lọc bỏ dòng không có tên (cột index 1)
    df = df.dropna(subset=[1])

    # 2. Mở file Word mẫu
    doc = Document(word_template)
    table = doc.tables[0]
    
    # Lấy dòng mẫu (dòng số 3 - index 2 hoặc 3 tùy file, thường là dòng có dữ liệu Lường Xuân Lộc)
    # Chúng ta lấy dòng cuối cùng hiện có của bảng làm mẫu định dạng
    sample_row = table.rows[-1]
    
    # Xóa các dòng cũ nhưng giữ lại header (thường là 2 dòng đầu)
    for i in range(len(table.rows)-1, 1, -1):
        table._element.remove(table.rows[i]._element)

    # 3. Nhóm dữ liệu theo Tỉnh (Cột 13) và Xã (Cột 12)
    grouped_tinh = df.groupby(13)
    
    t_idx = 1
    for tinh, group_tinh in grouped_tinh:
        tinh_name = str(tinh).strip()
        if not tinh_name or tinh_name == "nan": continue
        
        # --- Tạo dòng Tỉnh ---
        row_tinh = table.add_row()
        row_tinh.cells[0].merge(row_tinh.cells[-1])
        row_tinh.cells[0].text = f"{convert_to_roman(t_idx)}. Tỉnh {tinh_name}"
        run = row_tinh.cells[0].paragraphs[0].runs[0]
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        t_idx += 1
        
        grouped_xa = group_tinh.groupby(12)
        x_idx = 1
        record_idx = 1 # Số thứ tự trong xã
        
        for xa, group_xa in grouped_xa:
            xa_name = str(xa).strip()
            if not xa_name or xa_name == "nan": continue
            
            # --- Tạo dòng Xã ---
            row_xa = table.add_row()
            row_xa.cells[0].merge(row_xa.cells[-1])
            row_xa.cells[0].text = f"{x_idx}. Xã {xa_name}"
            run_x = row_xa.cells[0].paragraphs[0].runs[0]
            run_x.bold = True
            run_x.font.name = 'Times New Roman'
            run_x.font.size = Pt(11)
            x_idx += 1
            
            # --- Điền dữ liệu chi tiết ---
            for _, row_data in group_xa.iterrows():
                new_row = table.add_row()
                # Sao chép định dạng từ dòng mẫu để không bị lệch hàng
                copy_row_format(sample_row, new_row)
                
                new_cells = new_row.cells
                
                # Ánh xạ dữ liệu khớp với file Excel bạn gửi
                ho_ten = safe_get(row_data, 1)
                don_vi = safe_get(row_data, 2)
                cap_bac = safe_get(row_data, 3).replace("Binh nhì", "B2")
                chuc_vu = safe_get(row_data, 4).replace("Chiến sĩ", "CS")
                nhap_ngu = safe_get(row_data, 5)
                dan_toc = safe_get(row_data, 6)
                van_hoa = safe_get(row_data, 7)
                cccd = safe_get(row_data, 8)
                
                ngay = safe_get(row_data, 9)
                thang = safe_get(row_data, 10)
                nam = safe_get(row_data, 11)
                ngay_sinh = f"{ngay}/{thang}/{nam}" if ngay else ""
                
                bo = safe_get(row_data, 16)
                me = safe_get(row_data, 17)
                sdt = safe_get(row_data, 18)
                
                # Điền vào các cột (Index 0-20)
                new_cells[0].text = str(record_idx) # TT
                new_cells[1].text = ho_ten
                new_cells[2].text = ngay_sinh
                new_cells[3].text = cap_bac
                new_cells[4].text = chuc_vu
                new_cells[5].text = don_vi
                new_cells[6].text = nhap_ngu
                new_cells[7].text = "BN"
                new_cells[8].text = van_hoa
                new_cells[10].text = dan_toc
                new_cells[14].text = f"{xa_name}-{tinh_name}"
                new_cells[16].text = f"{bo}\n{me}" # Bố mẹ xuống dòng
                new_cells[18].text = sdt
                new_cells[19].text = cccd
                
                record_idx += 1

                # Áp dụng font Times New Roman cho tất cả ô vừa điền
                for cell in new_cells:
                    for p in cell.paragraphs:
                        for run in p.runs:
                            run.font.name = 'Times New Roman'
                            run.font.size = Pt(10)

    # Xuất file
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# --- Giao diện Streamlit ---
st.set_page_config(page_title="Xuất trích ngang chuẩn", layout="centered")
st.header("Trích ngang dữ liệu Tỉnh/Xã")
st.info("Hãy đảm bảo file Word mẫu có dòng 'Lường Xuân Lộc' ở cuối bảng để lấy định dạng.")

w_file = st.file_uploader("Tải Word mẫu (.docx)", type=['docx'])
e_file = st.file_uploader("Tải Excel/CSV dữ liệu", type=['xlsx', 'csv'])

if w_file and e_file:
    if st.button("🚀 Tạo file Word khớp mẫu"):
        try:
            res = process_files(e_file, w_file)
            st.success("Đã xử lý xong!")
            st.download_button("📥 Tải file kết quả", res, "Trich_Ngang_Chuan_Khop.docx")
        except Exception as e:
            st.error(f"Lỗi: {e}")
