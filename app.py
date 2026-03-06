import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO

def create_word_from_template(df_data, template_file):
    # 1. Đọc file mẫu
    doc_template = Document(template_file)
    new_doc = Document()
    
    # Lấy bảng mẫu (giả sử bảng 1)
    if len(doc_template.tables) == 0:
        st.error("File Word mẫu không có bảng nào!")
        return None
    template_table = doc_template.tables[0]
    
    # 2. Làm sạch dữ liệu Excel
    # Lọc bỏ các dòng không có tên
    df_data = df_data.dropna(subset=['Họ Và tên'])
    
    # Nhóm theo Tỉnh để tạo danh sách tự động
    provinces = df_data['Tỉnh'].unique()

    for prov in provinces:
        if pd.isna(prov): continue
        new_doc.add_heading(f"I. Tỉnh {prov}", level=1)
        
        prov_df = df_data[df_data['Tỉnh'] == prov]
        communes = prov_df['Xã'].unique()

        for comm in communes:
            if pd.isna(comm): continue
            new_doc.add_heading(f"1. Xã {comm}", level=2)
            
            # Tạo bảng mới dựa trên số cột của bảng mẫu
            new_table = new_doc.add_table(rows=1, cols=len(template_table.columns))
            new_table.style = template_table.style
            
            # Ghi tiêu đề cho bảng
            hdr_cells = new_table.rows[0].cells
            headers = [
                "TT", "Họ và tên", "Ngày sinh", "Cb", "CV", "Đơn vị", 
                "Nhập ngũ", "Thành phần", "Văn hóa", "Sức khỏe", "DT", "TG", 
                "Đoàn", "Đảng", "Quê quán", "Trú Quán", 
                "Họ tên bố, mẹ", "Họ tên vợ, con", "SĐT", "CCCD", "Ghi chú"
            ]
            for i, h in enumerate(headers):
                if i < len(hdr_cells):
                    hdr_cells[i].text = h

            # 3. Điền dữ liệu quân nhân
            comm_df = prov_df[prov_df['Xã'] == comm]
            for _, row in comm_df.iterrows():
                row_cells = new_table.add_row().cells
                
                # Cột 1: TT
                row_cells[0].text = str(row['TT']) if pd.notna(row['TT']) else ""
                # Cột 2: Họ tên
                row_cells[1].text = str(row['Họ Và tên'])
                # Cột 3: Ngày sinh (Ghép Ngày/Tháng/Năm)
                d = str(row['Ngày']).split('.')[0] if pd.notna(row['Ngày']) else ""
                m = str(row['Tháng']).split('.')[0] if pd.notna(row['Tháng']) else ""
                y = str(row['năm']).split('.')[0] if pd.notna(row['năm']) else ""
                row_cells[2].text = f"{d}/{m}/{y}"
                
                # Cột 15: Quê quán (TỰ ĐỘNG GỘP XÃ VÀ TỈNH)
                xa = str(row['Xã']) if pd.notna(row['Xã']) else ""
                tinh = str(row['Tỉnh']) if pd.notna(row['Tỉnh']) else ""
                row_cells[14].text = f"{xa}, {tinh}"
                
                # Cột 17: Bố mẹ
                bo = str(row['Bố']) if pd.notna(row['Bố']) else ""
                me = str(row['Mẹ']) if pd.notna(row['Mẹ']) else ""
                row_cells[16].text = f"{bo}, {me}"

                # Điền các thông tin khác nếu có
                row_cells[3].text = str(row['CB']) if 'CB' in row and pd.notna(row['CB']) else ""
                row_cells[4].text = str(row['CV']) if 'CV' in row and pd.notna(row['CV']) else ""
                row_cells[5].text = str(row['ĐV']) if 'ĐV' in row and pd.notna(row['ĐV']) else ""

    # Lưu và trả về file
    target_stream = BytesIO()
    new_doc.save(target_stream)
    target_stream.seek(0)
    return target_stream

# GIAO DIỆN STREAMLIT
st.title("Xử lý Trích Ngang - Tự động gộp Quê quán")

file_excel = st.file_uploader("Tải lên file Excel", type=["xlsx"])
file_word = st.file_uploader("Tải lên file Word mẫu", type=["docx"])

if file_excel and file_word:
    # Xử lý file Excel: Bỏ qua dòng 1 (dòng con của merged cells) để lấy đúng tiêu đề
    df = pd.read_excel(file_excel, header=0)
    
    # Nếu dòng đầu tiên chứa "Ngày, Tháng, năm" do merged cells, ta cần làm sạch
    if "Ngày" not in df.columns:
        df = pd.read_excel(file_excel, header=1)

    st.write("Xem trước dữ liệu (Quê quán sẽ được gộp khi xuất file):")
    st.dataframe(df.head())

    if st.button("Xuất file Word"):
        result = create_word_from_template(df, file_word)
        if result:
            st.download_button(
                label="📥 Tải xuống kết quả",
                data=result,
                file_name="Trich_Ngang_Hoan_Thien.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
