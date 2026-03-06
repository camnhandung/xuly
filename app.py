import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO

def create_word_file(df, template_path):
    doc_template = Document(template_path)
    if not doc_template.tables:
        st.error("File Word mẫu không có bảng!")
        return None
        
    # Lấy bảng mẫu để biết số cột thực tế
    template_table = doc_template.tables[0]
    num_cols = len(template_table.columns)
    
    new_doc = Document()
    
    # Làm sạch dữ liệu: Lấy những dòng có tên
    df = df[df['Họ Và tên'].notna()]

    # Lấy danh sách Tỉnh/Xã
    provinces = df['Tỉnh'].unique()

    for province in provinces:
        if pd.isna(province): continue
        new_doc.add_heading(f"I. Tỉnh {province}", level=1)
        
        province_df = df[df['Tỉnh'] == province]
        communes = province_df['Xã'].unique()

        for commune in communes:
            if pd.isna(commune): continue
            new_doc.add_heading(f"1. Xã {commune}", level=2)
            
            # Tạo bảng mới
            new_table = new_doc.add_table(rows=1, cols=num_cols)
            new_table.style = template_table.style
            
            # Copy tiêu đề từ bảng mẫu
            for i in range(num_cols):
                new_table.rows[0].cells[i].text = template_table.rows[0].cells[i].text

            # Điền dữ liệu từng người
            commune_df = province_df[province_df['Xã'] == commune]
            for _, row in commune_df.iterrows():
                row_cells = new_table.add_row().cells
                
                # Hàm an toàn để điền dữ liệu vào ô nếu cột đó tồn tại trong bảng Word
                def fill(index, text):
                    if index < num_cols:
                        row_cells[index].text = str(text) if pd.notna(text) else ""

                fill(0, row.get('TT', ''))
                fill(1, row.get('Họ Và tên', ''))
                
                # Ngày sinh
                d = str(row.get('Ngày', '')).split('.')[0]
                m = str(row.get('Tháng', '')).split('.')[0]
                y = str(row.get('năm', '')).split('.')[0]
                fill(2, f"{d}/{m}/{y}" if d != 'nan' and d != '' else "")
                
                fill(3, row.get('CB', ''))
                fill(4, row.get('CV', ''))
                fill(5, row.get('ĐV', ''))
                fill(6, row.get('N.N', ''))
                fill(8, row.get('Văn Hóa', ''))
                fill(10, row.get('Dân tộc', ''))

                # QUÊ QUÁN: GỘP XÃ VÀ TỈNH
                xa = str(row.get('Xã', ''))
                tinh = str(row.get('Tỉnh', ''))
                fill(14, f"{xa}, {tinh}") 
                
                # BỐ MẸ
                bo = str(row.get('Bố', ''))
                me = str(row.get('Mẹ', ''))
                fill(16, f"{bo}, {me}" if bo != 'nan' or me != 'nan' else "")
                
                fill(18, str(row.get('SDT gia đình', '')).split('.')[0])
                fill(19, row.get('Số CCCD', ''))

    bio = BytesIO()
    new_doc.save(bio)
    bio.seek(0)
    return bio

st.title("Tool Trích Ngang Quân Nhân (Đã sửa lỗi)")

ex_file = st.file_uploader("Tải lên file Excel", type=["xlsx"])
wd_file = st.file_uploader("Tải lên file Word mẫu", type=["docx"])

if ex_file and wd_file:
    # Đọc Excel thông minh
    df_temp = pd.read_excel(ex_file, header=None)
    header_idx = 0
    for i, r in df_temp.iterrows():
        if "Họ Và tên" in r.values:
            header_idx = i
            break
    
    df = pd.read_excel(ex_file, header=header_idx)
    
    # Chuẩn hóa tên cột để tránh lỗi KeyError
    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]

    st.write("Dữ liệu xem trước:")
    st.dataframe(df.head())

    if st.button("Xuất file Word"):
        result = create_word_file(df, wd_file)
        if result:
            st.download_button("📥 Tải file kết quả", result, "Ket_qua_final.docx")
