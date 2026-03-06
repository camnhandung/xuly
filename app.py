import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches
from docx.enum.section import WD_ORIENT
from io import BytesIO

def create_word_from_excel(df):
    doc = Document()
    
    # Chuyển khổ giấy sang ngang (Landscape) để bảng không bị tràn
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    new_width, new_height = section.page_height, section.page_width
    section.page_width = new_width
    section.page_height = new_height

    # Định nghĩa tiêu đề cột cho file Word [cite: 1]
    headers = [
        "TT", "Họ và tên", "Ngày sinh", "Cb", "CV", "Đơn vị", 
        "Nhập ngũ", "Thành phần", "Văn hóa", "Sức khỏe", "DT", "TG", 
        "Đoàn", "Đảng", "Quê quán", "Trú Quán", 
        "Họ tên bố, mẹ", "Họ tên vợ, con", "SĐT Gia đình", "Số CCCD", "Ghi chú"
    ]

    # Lọc bỏ các dòng không có tên để tránh tạo bảng trống [cite: 5]
    df = df.dropna(subset=['Họ Và tên'])

    # Nhóm theo Tỉnh và Xã 
    provinces = df['Tỉnh'].unique()

    for prov in provinces:
        if pd.isna(prov): continue
        doc.add_heading(f"I. Tỉnh {prov}", level=1)
        
        prov_df = df[df['Tỉnh'] == prov]
        communes = prov_df['Xã'].unique()

        for comm in communes:
            if pd.isna(comm): continue
            doc.add_heading(f"1. Xã {comm}", level=2)
            
            # Tạo bảng và đổ dữ liệu
            table = doc.add_table(rows=1, cols=len(headers))
            table.style = 'Table Grid'
            
            hdr_cells = table.rows[0].cells
            for i, h in enumerate(headers):
                hdr_cells[i].text = h

            comm_df = prov_df[prov_df['Xã'] == comm]
            for _, row in comm_df.iterrows():
                row_cells = table.add_row().cells
                
                # Điền dữ liệu dựa trên các cột từ file Excel 
                row_cells[0].text = str(row['TT'])
                row_cells[1].text = str(row['Họ Và tên'])
                
                # Ghép Ngày/Tháng/Năm sinh [cite: 2, 3]
                day = str(int(row['Ngày'])) if pd.notna(row['Ngày']) else ""
                month = str(int(row['Tháng'])) if pd.notna(row['Tháng']) else ""
                year = str(int(row['năm'])) if pd.notna(row['năm']) else ""
                row_cells[2].text = f"{day}/{month}/{year}"
                
                row_cells[3].text = str(row['CB'])
                row_cells[4].text = str(row['CV'])
                row_cells[5].text = str(row['ĐV'])
                row_cells[6].text = str(row['N.N'])
                row_cells[8].text = str(row['Văn Hóa'])
                row_cells[10].text = str(row['Dân tộc'])
                row_cells[14].text = f"{row['Xã']}, {row['Tỉnh']}"
                
                # Thông tin Bố, Mẹ 
                bo = str(row['Bố']) if pd.notna(row['Bố']) else ""
                me = str(row['Mẹ']) if pd.notna(row['Mẹ']) else ""
                row_cells[16].text = f"{bo}, {me}"
                
                row_cells[18].text = str(row['SDT gia đình']) if pd.notna(row['SDT gia đình']) else ""
                row_cells[19].text = str(row['Số CCCD']) if pd.notna(row['Số CCCD']) else ""
                row_cells[20].text = str(row['Ghi chú']) if pd.notna(row['Ghi chú']) else ""

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# Giao diện Streamlit
st.title("Xuất Trích Ngang từ File Excel (.xlsx)")

# Chấp nhận file .xlsx
uploaded_file = st.file_uploader("Tải file Excel quân nhân", type=["xlsx"])

if uploaded_file:
    # Đọc file Excel (mặc định trang đầu tiên)
    df = pd.read_excel(uploaded_file)
    
    st.success("Đã nhận file Excel!")
    st.dataframe(df.head())

    if st.button("Tạo file Word"):
        with st.spinner("Đang xử lý..."):
            word_output = create_word_from_excel(df)
            st.download_button(
                label="📥 Tải xuống file Word mẫu",
                data=word_output,
                file_name="Trich_ngang_tu_Excel.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
