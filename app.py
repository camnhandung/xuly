import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO

def create_word_output(df_data, template_file):
    # Đọc mẫu để lấy định dạng
    doc_template = Document(template_file)
    if not doc_template.tables:
        st.error("File Word mẫu không có bảng!")
        return None
    
    template_table = doc_template.tables[0]
    # Xác định số cột thực tế của bảng mẫu
    num_cols = len(template_table.columns)
    
    new_doc = Document()

    # Nhóm theo Tỉnh và Xã (Dữ liệu đã được gộp và làm sạch ở bước ngoài)
    provinces = df_data['Tỉnh'].unique()

    for prov in provinces:
        if pd.isna(prov) or str(prov).strip() == "": continue
        new_doc.add_heading(f"I. Tỉnh {prov}", level=1)
        
        prov_df = df_data[df_data['Tỉnh'] == prov]
        communes = prov_df['Xã'].unique()

        for comm in communes:
            if pd.isna(comm) or str(comm).strip() == "": continue
            new_doc.add_heading(f"1. Xã {comm}", level=2)
            
            # Tạo bảng mới copy từ template
            new_table = new_doc.add_table(rows=1, cols=num_cols)
            new_table.style = template_table.style
            
            # Copy header từ template
            for i in range(num_cols):
                new_table.rows[0].cells[i].text = template_table.rows[0].cells[i].text

            # Lấy dữ liệu quân nhân trong xã
            comm_df = prov_df[prov_df['Xã'] == comm]
            for _, row in comm_df.iterrows():
                row_cells = new_table.add_row().cells
                
                # Hàm điền dữ liệu an toàn để tránh IndexError
                def safe_fill(idx, value):
                    if idx < num_cols:
                        row_cells[idx].text = str(value) if pd.notna(value) else ""

                # Điền dữ liệu theo đúng vị trí bảng mẫu của bạn
                safe_fill(0, row.get('TT', ''))
                safe_fill(1, row.get('Họ Và tên', ''))
                
                # Ngày sinh (Ghép từ 3 cột)
                d = str(row.get('Ngày', '')).replace('.0', '')
                m = str(row.get('Tháng', '')).replace('.0', '')
                y = str(row.get('năm', '')).replace('.0', '')
                safe_fill(2, f"{d}/{m}/{y}" if d != 'nan' and d != '' else "")
                
                safe_fill(3, row.get('CB', ''))
                safe_fill(4, row.get('CV', ''))
                safe_fill(5, row.get('ĐV', ''))
                safe_fill(6, row.get('N.N', ''))
                safe_fill(8, row.get('Văn Hóa', ''))
                safe_fill(10, row.get('Dân tộc', ''))

                # --- QUÊ QUÁN: GỘP XÃ VÀ TỈNH ---
                xa_tinh = f"{row.get('Xã', '')}, {row.get('Tỉnh', '')}"
                safe_fill(14, xa_tinh) # Cột Quê quán thường là cột 15 (index 14)
                
                # --- HỌ TÊN BỐ MẸ ---
                bo_me = f"{row.get('Bố', '')}, {row.get('Mẹ', '')}"
                safe_fill(16, bo_me if len(bo_me) > 2 else "")

                safe_fill(18, str(row.get('SDT gia đình', '')).replace('.0', ''))
                safe_fill(19, row.get('Số CCCD', ''))

    bio = BytesIO()
    new_doc.save(bio)
    bio.seek(0)
    return bio

# --- GIAO DIỆN STREAMLIT ---
st.set_page_config(page_title="Xử lý Trích Ngang", layout="wide")
st.title("Phần mềm tạo Trích Ngang tự động")

col1, col2 = st.columns(2)
with col1:
    ex_file = st.file_uploader("Tải Excel (.xlsx)", type=["xlsx"])
with col2:
    wd_file = st.file_uploader("Tải Word mẫu (.docx)", type=["docx"])

if ex_file and wd_file:
    # Bước 1: Đọc Excel và xử lý tiêu đề gộp
    df_raw = pd.read_excel(ex_file, header=None)
    
    # Tìm dòng chứa "Họ Và tên"
    header_idx = 0
    for i, r in df_raw.iterrows():
        if "Họ Và tên" in r.values:
            header_idx = i
            break
            
    df = pd.read_excel(ex_file, header=header_idx)
    
    # Làm sạch tên cột (xóa khoảng trắng thừa)
    df.columns = [str(c).strip() for c in df.columns]
    
    # Bỏ các dòng rác (không có tên)
    df = df.dropna(subset=['Họ Và tên'])
    # Chuyển các cột số (Ngày, Tháng, Năm) sang dạng chuỗi để gộp không bị lỗi .0
    for col in ['Ngày', 'Tháng', 'năm', 'TT']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace('.0', '', regex=False)

    st.success("Dữ liệu nạp thành công!")
    st.dataframe(df.head())

    if st.button("🚀 Bắt đầu tạo file Word"):
        with st.spinner("Đang xử lý..."):
            result_docx = create_word_output(df, wd_file)
            if result_docx:
                st.download_button(
                    label="📥 Tải file Word kết quả",
                    data=result_docx,
                    file_name="Ket_qua_Trich_Ngang.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
