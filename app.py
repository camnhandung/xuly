import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO

def create_word_file(df, template_path):
    # Đọc file mẫu để lấy định dạng bảng
    doc = Document(template_path)
    if not doc.tables:
        st.error("File Word mẫu không có bảng nào!")
        return None
        
    template_table = doc.tables[0]
    new_doc = Document()
    
    # Lọc bỏ các dòng không có tên
    df = df.dropna(subset=['Họ Và tên'])

    # Lấy danh sách Tỉnh duy nhất để phân loại
    # Lưu ý: tên cột 'Tỉnh' và 'Xã' cần được đảm bảo tồn tại sau khi xử lý header
    provinces = df['Tỉnh'].unique()

    for province in provinces:
        if pd.isna(province): continue
        new_doc.add_heading(f"I. Tỉnh {province}", level=1)
        
        province_df = df[df['Tỉnh'] == province]
        communes = province_df['Xã'].unique()

        for commune in communes:
            if pd.isna(commune): continue
            new_doc.add_heading(f"1. Xã {commune}", level=2)
            
            # Tạo bảng mới copy từ template
            new_table = new_doc.add_table(rows=1, cols=len(template_table.columns))
            new_table.style = template_table.style
            
            # Copy dòng tiêu đề
            for i, cell in enumerate(template_table.rows[0].cells):
                new_table.rows[0].cells[i].text = cell.text

            commune_df = province_df[province_df['Xã'] == commune]
            for _, row in commune_df.iterrows():
                row_cells = new_table.add_row().cells
                
                # Điền dữ liệu - Index dựa trên file mẫu bạn gửi
                row_cells[0].text = str(row.get('TT', '')) if pd.notna(row.get('TT')) else ""
                row_cells[1].text = str(row.get('Họ Và tên', ''))
                
                # Ghép Ngày/Tháng/Năm
                d = str(row.get('Ngày', '')).split('.')[0]
                m = str(row.get('Tháng', '')).split('.')[0]
                y = str(row.get('năm', '')).split('.')[0]
                row_cells[2].text = f"{d}/{m}/{y}" if d != 'nan' else ""
                
                row_cells[3].text = str(row.get('CB', '')) if pd.notna(row.get('CB')) else ""
                row_cells[4].text = str(row.get('CV', '')) if pd.notna(row.get('CV')) else ""
                row_cells[5].text = str(row.get('ĐV', '')) if pd.notna(row.get('ĐV')) else ""
                
                # QUÊ QUÁN: Tự động gộp Xã và Tỉnh
                xa = str(row.get('Xã', ''))
                tinh = str(row.get('Tỉnh', ''))
                row_cells[14].text = f"{xa}, {tinh}"
                
                # Bố mẹ
                bo = str(row.get('Bố', '')) if pd.notna(row.get('Bố')) else ""
                me = str(row.get('Mẹ', '')) if pd.notna(row.get('Mẹ')) else ""
                row_cells[16].text = f"{bo}, {me}"
                
                # SDT và CCCD
                row_cells[18].text = str(row.get('SDT gia đình', '')).split('.')[0] if pd.notna(row.get('SDT gia đình')) else ""
                row_cells[19].text = str(row.get('Số CCCD', '')) if pd.notna(row.get('Số CCCD')) else ""

    bio = BytesIO()
    new_doc.save(bio)
    bio.seek(0)
    return bio

# Giao diện Streamlit
st.title("Xuất Trích Ngang Quân Nhân")

ex_file = st.file_uploader("Tải lên file Excel (.xlsx)", type=["xlsx"])
wd_file = st.file_uploader("Tải lên file Word Mẫu (.docx)", type=["docx"])

if ex_file and wd_file:
    # Xử lý đọc file Excel có Merged Cells (tiêu đề 2 dòng)
    df_raw = pd.read_excel(ex_file, header=None)
    
    # Tìm dòng chứa "Họ Và tên" để làm header
    header_row_idx = 0
    for i, row in df_raw.iterrows():
        if "Họ Và tên" in row.values:
            header_row_idx = i
            break
            
    # Đọc lại với header đúng và lấy thêm thông tin từ dòng phụ nếu cần
    df = pd.read_excel(ex_file, header=header_row_idx)
    
    # Sửa lỗi tên cột nếu dòng 2 chứa Xã/Tỉnh thay vì dòng 1
    if "Xã" not in df.columns:
        # Thử lấy từ dòng tiếp theo nếu bị trống do merged
        df.columns = [
            'TT', 'Họ Và tên', 'ĐV', 'CB', 'CV', 'N.N', 'Dân tộc', 'Văn Hóa', 'Số CCCD', 
            'Ngày', 'Tháng', 'năm', 'Xã', 'Tỉnh', 'Khu vực phòng thủ', 'Ghi chú', 'Bố', 'Mẹ', 'SDT gia đình', 'gửi'
        ]

    st.write("Dữ liệu đã sẵn sàng:")
    st.dataframe(df.head())

    if st.button("Bắt đầu xuất file Word"):
        result = create_word_file(df, wd_file)
        if result:
            st.download_button("📥 Tải xuống kết quả", result, "Ket_qua_trich_ngang.docx")
