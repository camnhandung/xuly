import streamlit as st
import pandas as pd
from docx import Document

# Upload file Excel và Word
excel_file = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx"])
word_file = st.file_uploader("Upload file Word (.docx)", type=["docx"])

if excel_file and word_file:
    # Đọc dữ liệu từ Excel
    df = pd.read_excel(excel_file)

    # Đọc file Word
    doc = Document(word_file)

    # Nhóm dữ liệu theo Tỉnh -> Xã
    grouped = df.groupby(["Tỉnh", "Xã"])

    for (tinh, xa), group in grouped:
        # Kiểm tra xem tỉnh đã có trong Word chưa
        found_tinh = any(tinh in p.text for p in doc.paragraphs)
        if not found_tinh:
            doc.add_paragraph(f"II. Tỉnh {tinh}")

        # Thêm xã
        doc.add_paragraph(f"{xa}")

        # Tạo bảng mới cho xã
        table = doc.add_table(rows=1, cols=len(df.columns))
        hdr_cells = table.rows[0].cells
        for i, col in enumerate(df.columns):
            hdr_cells[i].text = col

        # Thêm dữ liệu từng dòng
        for _, row in group.iterrows():
            row_cells = table.add_row().cells
            for i, value in enumerate(row):
                row_cells[i].text = str(value)

    # Lưu file kết quả
    output_path = "ket_qua.docx"
    doc.save(output_path)
    st.success("Đã tạo file Word hoàn chỉnh!")
    with open(output_path, "rb") as f:
        st.download_button("Tải file Word kết quả", f, file_name="ket_qua.docx")
