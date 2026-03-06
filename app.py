import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO

def create_final_word(df, template_path):
    # Đọc mẫu để lấy style bảng
    doc_temp = Document(template_path)
    if not doc_temp.tables:
        st.error("File Word mẫu không có bảng!")
        return None
    
    template_table = doc_temp.tables[0]
    num_cols = len(template_table.columns)
    
    new_doc = Document()

    # Nhóm dữ liệu theo Tỉnh (Cột 13) và Xã (Cột 12) theo vị trí index
    # (Index 12 là cột M, Index 13 là cột N trong file Excel của bạn)
    provinces = df.iloc[:, 13].unique()

    for prov in provinces:
        if pd.isna(prov) or str(prov).strip() == "" or str(prov).lower() == "tỉnh": continue
        new_doc.add_heading(f"I. Tỉnh {prov}", level=1)
        
        prov_df = df[df.iloc[:, 13] == prov]
        communes = prov_df.iloc[:, 12].unique()

        for comm in communes:
            if pd.isna(comm) or str(comm).strip() == "" or str(comm).lower() == "xã": continue
            new_doc.add_heading(f"1. Xã {comm}", level=2)
            
            # Tạo bảng mới và copy header
            table = new_doc.add_table(rows=1, cols=num_cols)
            table.style = template_table.style
            for i in range(num_cols):
                table.rows[0].cells[i].text = template_table.rows[0].cells[i].text

            # Lọc dữ liệu theo xã
            comm_df = prov_df[prov_df.iloc[:, 12] == comm]
            
            for _, row in comm_df.iterrows():
                row_cells = table.add_row().cells
                
                def fill(idx, val):
                    if idx < num_cols:
                        # Xử lý xóa .0 cho số và chuyển về string
                        v = str(val).replace('.0', '') if pd.notna(val) else ""
                        row_cells[idx].text = v if v != 'nan' else ""

                # Điền dựa theo thứ tự cột trong file Excel của bạn
                fill(0, row.iloc[0])   # TT
                fill(1, row.iloc[1])   # Họ tên
                
                # Cột 2: Ngày sinh (Ghép cột 9, 10, 11)
                d = str(row.iloc[9]).replace('.0', '')
                m = str(row.iloc[10]).replace('.0', '')
                y = str(row.iloc[11]).replace('.0', '')
                fill(2, f"{d}/{m}/{y}" if d != 'nan' and d != '' else "")

                fill(3, row.iloc[3])   # Cấp bậc (CB)
                fill(4, row.iloc[4])   # Chức vụ (CV)
                fill(5, row.iloc[2])   # Đơn vị (ĐV)
                fill(6, row.iloc[5])   # Nhập ngũ (N.N)
                fill(8, row.iloc[7])   # Văn hóa
                fill(10, row.iloc[6])  # Dân tộc
                
                # Cột 14: Quê quán (Tự động gộp Xã và Tỉnh)
                xa = str(row.iloc[12])
                tinh = str(row.iloc[13])
                fill(14, f"{xa}, {tinh}")

                # Cột 16: Bố mẹ (Ghép cột 16, 17)
                bo = str(row.iloc[16])
                me = str(row.iloc[17])
                fill(16, f"{bo}, {me}" if bo != 'nan' or me != 'nan' else "")

                fill(18, row.iloc[18]) # SDT
                fill(19, row.iloc[8])  # CCCD

    output = BytesIO()
    new_doc.save(output)
    output.seek(0)
    return output

# --- Streamlit UI ---
st.set_page_config(page_title="Hệ thống Trích Ngang", layout="wide")
st.title("Tạo Trích Ngang Quân Nhân Tự Động")

col1, col2 = st.columns(2)
with col1:
    ex_file = st.file_uploader("Bước 1: Tải Excel (.xlsx)", type=["xlsx"])
with col2:
    wd_file = st.file_uploader("Bước 2: Tải Word mẫu (.docx)", type=["docx"])

if ex_file and wd_file:
    # Đọc Excel: Lấy dữ liệu từ dòng 3 trở đi để bỏ qua merged headers
    # Dữ liệu thật bắt đầu từ row 2 (0-indexed)
    df = pd.read_excel(ex_file, header=None)
    
    # Lọc bỏ các dòng tiêu đề rác, chỉ giữ dòng có số TT là số
    df_clean = df[pd.to_numeric(df.iloc[:, 0], errors='coerce').notnull()]

    st.success(f"Đã tìm thấy {len(df_clean)} quân nhân.")
    st.dataframe(df_clean.head())

    if st.button("🚀 Xuất file Word"):
        with st.spinner("Đang xử lý dữ liệu..."):
            res = create_final_word(df_clean, wd_file)
            if res:
                st.download_button("📥 Tải xuống kết quả", res, "Trich_Ngang_Chuan.docx")
