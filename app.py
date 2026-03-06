import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches
from io import BytesIO

def create_word_file(df):
    doc = Document()
    
    # Tiêu đề cột dựa theo file mẫu của bạn
    headers = [
        "TT", "Họ và tên", "Ngày tháng năm sinh", "Cb", "CV", "Đơn vị", 
        "Nhập ngũ", "Thành phần", "Văn hóa", "Sức khỏe", "DT", "TG", 
        "Ngày vào Đoàn", "Ngày vào Đảng", "Quê quán", "Trú Quán", 
        "Họ tên bố, mẹ", "Họ tên vợ, con", "SỐ ĐIỆN THOẠI GIA ĐÌNH", "Số cccd", "Ghi chú"
    ]

    # Nhóm dữ liệu theo Tỉnh và Xã (dựa trên cột Quê quán)
    # Giả sử định dạng quê quán là "Xã, Tỉnh" hoặc chỉ có tên
    unique_provinces = df['Tỉnh'].unique()

    for province in unique_provinces:
        doc.add_heading(f"I. Tỉnh {province}", level=1)
        
        province_df = df[df['Tỉnh'] == province]
        unique_communes = province_df['Xã'].unique()

        for commune in unique_communes:
            doc.add_heading(f"1. Xã {commune}", level=2)
            
            # Tạo bảng
            table = doc.add_table(rows=1, cols=len(headers))
            table.style = 'Table Grid'
            hdr_cells = table.rows[0].cells
            for i, header in enumerate(headers):
                hdr_cells[i].text = header

            commune_df = province_df[province_df['Xã'] == commune]
            
            for idx, row in commune_df.iterrows():
                row_cells = table.add_row().cells
                # Map dữ liệu từ Excel (CSV) sang Word
                row_cells[0].text = str(row['TT'])
                row_cells[1].text = str(row['Họ Và tên'])
                # Ghép ngày/tháng/năm sinh
                dob = f"{row['Ngày']}/{row['Tháng']}/{row['năm']}"
                row_cells[2].text = dob
                row_cells[3].text = str(row['CB'])
                row_cells[4].text = str(row['CV'])
                row_cells[5].text = str(row['ĐV'])
                row_cells[6].text = str(row['N.N'])
                row_cells[8].text = str(row['Văn Hóa'])
                row_cells[10].text = str(row['Dân tộc'])
                row_cells[14].text = f"{row['Xã']}-{row['Tỉnh']}"
                row_cells[16].text = f"{row['Bố']}, {row['Mẹ']}"
                row_cells[18].text = str(row['SDT gia đình'])
                row_cells[19].text = str(row['Số CCCD'])
                row_cells[20].text = str(row['Ghi chú'])

    # Lưu vào bộ nhớ đệm
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# Giao diện Streamlit
st.title("Chuyển đổi dữ liệu Quân nhân (Excel -> Word)")
st.write("Tải lên file CSV/Excel để tạo file Word trích ngang theo mẫu.")

uploaded_file = st.file_uploader("Chọn file dữ liệu (CSV)", type=["csv"])

if uploaded_file is not None:
    # Đọc dữ liệu, bỏ qua các dòng trống
    df = pd.read_csv(uploaded_file)
    df = df.dropna(subset=['Họ Và tên']) # Chỉ lấy dòng có tên
    
    st.success("Đã tải dữ liệu thành công!")
    st.dataframe(df.head())

    if st.button("Tạo file Word"):
        word_file = create_word_file(df)
        st.download_button(
            label="Tải xuống file Word",
            data=word_file,
            file_name="Tong_hop_trich_ngang.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
