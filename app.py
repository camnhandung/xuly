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
        if pd.isna(val) or str(val).lower() == "nan": return default
        if isinstance(val, float) and val.is_integer(): return str(int(val))
        return str(val).strip()
    except:
        return default

def apply_format_from_sample(sample_row, target_row):
    """Sao chép định dạng đoạn văn (căn lề) từ dòng mẫu sang dòng mới"""
    for i, cell in enumerate(target_row.cells):
        if i < len(sample_row.cells):
            # Sao chép kiểu căn lề (giữa/trái/phải)
            target_row.cells[i].paragraphs[0].alignment = sample_row.cells[i].paragraphs[0].alignment

def process_files(excel_file, word_template):
    # 1. Đọc dữ liệu Excel (Bỏ qua 2 dòng đầu)
    if excel_file.name.endswith('.csv'):
        df = pd.read_csv(excel_file, skiprows=2, header=None)
    else:
        df = pd.read_excel(excel_file, skiprows=2, header=None)
    
    df = df.dropna(subset=[1]) # Lọc dòng có tên

    # 2. Mở file Word mẫu
    doc = Document(word_template)
    table = doc.tables[0]
    
    # Lấy dòng cuối cùng làm mẫu định dạng (dòng Lường Xuân Lộc)
    sample_row = table.rows[-1]
    
    # Xóa các dòng cũ, giữ lại 2 dòng tiêu đề đầu
    for i in range(len(table.rows)-1, 1, -1):
        table._element.remove(table.rows[i]._element)

    # 3. Nhóm dữ liệu
    # Theo Tỉnh (Cột 13), Xã (Cột 12) dựa trên file Excel của bạn
    grouped_tinh = df.groupby(13)
    
    t_idx = 1
    for tinh, group_tinh in grouped_tinh:
        tinh_name = str(tinh).strip()
        if not tinh_name or tinh_name == "nan": continue
        
        # Dòng Tỉnh
        row_tinh = table.add_row()
        row_tinh.cells[0].merge(row_tinh.cells[-1])
        row_tinh.cells[0].text = f"{convert_to_roman(t_idx)}. Tỉnh {tinh_name}"
        para_t = row_tinh.cells[0].paragraphs[0]
        para_t.runs[0].bold = True
        para_t.runs[0].font.name = 'Times New Roman'
        para_t.runs[0].font.size = Pt(12)
        t_idx += 1
        
        grouped_xa = group_tinh.groupby(12)
        x_idx = 1
        stt_trong_tinh = 1 # Đánh số TT 1, 2, 3...
        
        for xa, group_xa in grouped_xa:
            xa_name = str(xa).strip()
            if not xa_name or xa_name == "nan": continue
            
            # Dòng Xã
            row_xa = table.add_row()
            row_xa.cells[0].merge(row_xa.cells[-1])
            row_xa.cells[0].text = f"{x_idx}. Xã {xa_name}"
            para_x = row_xa.cells[0].paragraphs[0]
            para_x.runs[0].bold = True
            para_x.runs[0].font.name = 'Times New Roman'
            para_x.runs[0].font.size = Pt(11)
            x_idx += 1
            
            # Dòng dữ liệu chi tiết
            for _, row_data in group_xa.iterrows():
                new_row = table.add_row()
                apply_format_from_sample(sample_row, new_row)
                
                cells = new_row.cells
                
                # Trích xuất dữ liệu từ các cột tương ứng trong Excel
                ho_ten = safe_get(row_data, 1)
                don_vi = safe_get(row_data, 2)
                cap_bac = safe_get(row_data, 3).replace("Binh nhì", "B2")
                chuc_vu = safe_get(row_data, 4).replace("Chiến sĩ", "CS")
                nhap_ngu = safe_get(row_data, 5)
                dan_toc = safe_get(row_data, 6)
                van_hoa = safe_get(row_data, 7)
                cccd = safe_get(row_data, 8)
                
                ngay_sinh = f"{safe_get(row_data, 9)}/{safe_get(row_data, 10)}/{safe_get(row_data, 11)}"
                
                bo = safe_get(row_data, 16)
                me = safe_get(row_data, 17)
                sdt = safe_get(row_data, 18)
                
                # Điền vào Word (Khớp theo các cột trong bảng mẫu)
                cells[0].text = str(stt_trong_tinh)
                cells[1].text = ho_ten
                cells[2].text = ngay_sinh
                cells[3].text = cap_bac
                cells[4].text = chuc_vu
                cells[5].text = don_vi
                cells[6].text = nhap_ngu
                cells[7].text = "BN"
                cells[8].text = van_hoa
                cells[10].text = dan_toc
                cells[14].text = f"{xa_name}-{tinh_name}"
                cells[16].text = f"{bo}\n{me}" # Xuống dòng tên bố mẹ
                cells[18].text = sdt
                cells[19].text = cccd
                
                stt_trong_tinh += 1

                # Định dạng lại toàn bộ chữ trong dòng về Times New Roman 10pt
                for cell in cells:
                    for p in cell.paragraphs:
                        for run in p.runs:
                            run.font.name = 'Times New Roman'
                            run.font.size = Pt(10)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# --- Giao diện Streamlit ---
st.title("Tạo Trích Ngang Chuẩn Mẫu")
w_file = st.file_uploader("Tải Word mẫu", type=['docx'])
e_file = st.file_uploader("Tải Excel dữ liệu", type=['xlsx', 'csv'])

if w_file and e_file:
    if st.button("🚀 Xuất file Word"):
        try:
            res = process_files(e_file, w_file)
            st.download_button("📥 Tải kết quả", res, "Ket_qua_trich_ngang.docx")
        except Exception as e:
            st.error(f"Lỗi: {e}")
