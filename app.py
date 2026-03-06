import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from io import BytesIO
import copy

def create_word_from_template(df_data, template_file):
    # Đọc file template
    doc_template = Document(template_file)
    
    # Tạo một file Word mới để chứa kết quả
    new_doc = Document()
    
    # Lấy bảng mẫu đầu tiên trong file template để dùng làm khuôn
    # (Giả sử bảng mẫu của bạn là bảng đầu tiên trong file word)
    template_table = doc_template.tables[0]
    
    # Lọc bỏ dòng trống trong Excel
    df_data = df_data.dropna(subset=['Họ Và tên'])

    # Nhóm theo Tỉnh
    provinces = df_data['Tỉnh'].unique()

    for prov in provinces:
        if pd.isna(prov): continue
        new_doc.add_heading(f"I. Tỉnh {prov}", level=1)
        
        prov_df = df_data[df_data['Tỉnh'] == prov]
        communes = prov_df['Xã'].unique()

        for comm in communes:
            if pd.isna(comm): continue
            new_doc.add_heading(f"1. Xã {comm}", level=2)
            
            # Copy bảng từ template vào file mới
            new_table = new_doc.add_table(rows=1, cols=len(template_table.columns))
            new_table.style = template_table.style
            
            # Sao chép tiêu đề từ bảng mẫu
            for i, cell in enumerate(template_table.rows[0].cells):
                new_table.rows[0].cells[i].text = cell.text

            # Lấy dữ liệu của xã hiện tại
            comm_df = prov_df[prov_df['Xã'] == comm]
            
            for _, row in comm_df.iterrows():
                row_cells = new_table.add_row().cells
                
                # Điền các cột (Dựa theo vị trí cột trong file Excel của bạn)
                row_cells[0].text = str(row['TT']) if pd.notna(row['TT']) else ""
                row_cells[1].text = str(row['Họ Và tên'])
                
                # Ghép Ngày/Tháng/Năm
                d = str(int(row['Ngày'])) if pd.notna(row['Ngày']) else ""
                m = str(int(row['Tháng'])) if pd.notna(row['Tháng']) else ""
                y = str(int(row['năm'])) if pd.notna(row['năm']) else ""
                row_cells[2].text = f"{d}/{m}/{y}"
                
                row_cells[3].text = str(row['CB']) if pd.notna(row['CB']) else ""
                row_cells[4].text = str(row['CV']) if pd.notna(row['CV']) else ""
                row_cells[5].text = str(row['ĐV']) if pd.notna(row['ĐV']) else ""
                row_cells[6].text = str(row['N.N']) if pd.notna(row['N.N']) else ""
                row_cells[8].text = str(row['Văn Hóa']) if pd.notna(row['Văn Hóa']) else ""
                row_cells[10].text = str(row['Dân tộc']) if pd.notna(row['Dân tộc']) else ""
                
                # Quê quán (Xã - Tỉnh)
                row_cells[14].text = f"{row['Xã']} - {row['Tỉnh']}"
                
                # Họ tên bố mẹ
                bo = str(row['Bố']) if pd.notna(row['Bố']) else ""
                me = str(row['Mẹ']) if pd.notna(row['Mẹ']) else ""
                row_cells[16].text = f"{bo}, {me}"
                
                row_cells[18].text = str(row['SDT gia đình']) if pd.notna(row['SDT gia đình']) else ""
                row_cells[19].text = str(row['Số CCCD']) if pd.notna(row['Số CCCD']) else ""
                row_cells[20].text = str(row['Ghi chú']) if pd.notna(row['Ghi chú']) else ""

    # Lưu file vào bộ nhớ
    target_stream = BytesIO()
    new_doc.save(target_stream)
    target_stream.seek(0)
    return target_stream

# GIAO DIỆN STREAMLIT
st.set_page_config(page_title="Tool Trích Ngang", layout="wide")
st.title("Tạo File Word Trích Ngang Theo Mẫu")

col1, col2 = st.columns(2)

with col1:
    excel_file = st.file_uploader("1. Tải lên file Excel dữ liệu (.xlsx)", type=["xlsx"])

with col2:
    word_template = st.file_uploader("2. Tải lên file Word mẫu (.docx)", type=["docx"])

if excel_file and word_template:
    # Đọc dữ liệu Excel, bắt đầu từ dòng chứa tiêu đề thật (thường là dòng 0 hoặc 1)
    df = pd.read_excel(excel_file, header=0)
    
    # Hiển thị bản xem trước
    st.write("Dữ liệu tìm thấy:")
    st.dataframe(df.head(5))

    if st.button("Bắt đầu xử lý và Tải về"):
        with st.spinner("Đang tạo file..."):
            try:
                final_docx = create_word_from_template(df, word_template)
                st.download_button(
                    label="📥 Tải xuống file Word kết quả",
                    data=final_docx,
                    file_name="Ket_qua_trich_ngang.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                st.success("Thành công!")
            except Exception as e:
                st.error(f"Có lỗi xảy ra: {e}")
