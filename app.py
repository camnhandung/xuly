import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO

def create_word_from_template(df_data, template_file):
    doc_template = Document(template_file)
    new_doc = Document()
    
    if len(doc_template.tables) == 0:
        st.error("File Word mẫu không có bảng!")
        return None
    template_table = doc_template.tables[0]

    # Nhóm theo Tỉnh và Xã để tạo tiêu đề danh sách tự động
    # Sử dụng tên cột theo vị trí (index) nếu tên cột bị sai lệch do merged cells
    provinces = df_data['Tỉnh'].unique()

    for prov in provinces:
        if pd.isna(prov) or str(prov).strip() == "" or str(prov).lower() == "tỉnh": continue
        new_doc.add_heading(f"I. Tỉnh {prov}", level=1)
        
        prov_df = df_data[df_data['Tỉnh'] == prov]
        communes = prov_df['Xã'].unique()

        for comm in communes:
            if pd.isna(comm) or str(comm).strip() == "" or str(comm).lower() == "xã": continue
            new_doc.add_heading(f"1. Xã {comm}", level=2)
            
            # Tạo bảng mới và copy tiêu đề
            new_table = new_doc.add_table(rows=1, cols=len(template_table.columns))
            new_table.style = template_table.style
            for i, cell in enumerate(template_table.rows[0].cells):
                new_table.rows[0].cells[i].text = cell.text

            comm_df = prov_df[prov_df['Xã'] == comm]
            for _, row in comm_df.iterrows():
                row_cells = new_table.add_row().cells
                
                # Điền dữ liệu - Sử dụng get() để tránh lỗi nếu thiếu cột
                row_cells[0].text = str(row.get('TT', '')) if pd.notna(row.get('TT')) else ""
                row_cells[1].text = str(row.get('Họ Và tên', ''))
                
                # Ghép Ngày/Tháng/Năm sinh
                d = str(row.get('Ngày', '')).split('.')[0]
                m = str(row.get('Tháng', '')).split('.')[0]
                y = str(row.get('năm', '')).split('.')[0]
                row_cells[2].text = f"{d}/{m}/{y}" if d != 'nan' else ""
                
                # TỰ ĐỘNG GỘP QUÊ QUÁN: Xã + Tỉnh
                xa = str(row.get('Xã', ''))
                tinh = str(row.get('Tỉnh', ''))
                row_cells[14].text = f"{xa}, {tinh}"
                
                # Ghép Bố Mẹ
                bo = str(row.get('Bố', '')) if pd.notna(row.get('Bố')) else ""
                me = str(row.get('Mẹ', '')) if pd.notna(row.get('Mẹ')) else ""
                row_cells[16].text = f"{bo}, {me}"
                
                # Các cột khác điền nếu có dữ liệu
                row_cells[3].text = str(row.get('CB', ''))
                row_cells[4].text = str(row.get('CV', ''))
                row_cells[5].text = str(row.get('ĐV', ''))
                row_cells[18].text = str(row.get('SDT gia đình', '')).split('.')[0] if pd.notna(row.get('SDT gia đình')) else ""

    target_stream = BytesIO()
    new_doc.save(target_stream)
    target_stream.seek(0)
    return target_stream

st.title("Tool Trích Ngang Quân Nhân")

ex_file = st.file_uploader("Tải lên file Excel", type=["xlsx"])
wd_file = st.file_uploader("Tải lên file Word mẫu", type=["docx"])

if ex_file and wd_file:
    # Đọc file Excel và tự động xử lý header nếu có ô gộp (merged cells)
    # Thử đọc dòng 0, nếu thấy 'Họ Và tên' bị nan thì đọc từ dòng 1
    df = pd.read_excel(ex_file)
    if 'Họ Và tên' not in df.columns:
        df = pd.read_excel(ex_file, header=1)
    
    # Xóa bỏ các dòng tiêu đề phụ lặp lại (nếu có)
    df = df[df['Họ Và tên'].notna()]
    df = df[df['Họ Và tên'] != 'Họ Và tên']

    st.success("Đã nạp dữ liệu!")
    st.dataframe(df.head())

    if st.button("Tạo và Tải file Word"):
        result = create_word_from_template(df, wd_file)
        st.download_button("📥 Tải file kết quả", result, "Ket_qua.docx")
