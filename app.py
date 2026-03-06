import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO

# Cấu hình giao diện Streamlit
st.set_page_config(page_title="Tự động hóa Trích ngang", layout="wide")

def format_roman(n):
    """Chuyển số sang số La Mã cho tiêu đề Tỉnh"""
    roman = {1: 'I', 2: 'II', 3: 'III', 4: 'IV', 5: 'V', 6: 'VI', 7: 'VII', 8: 'VIII', 9: 'IX', 10: 'X'}
    return roman.get(n, str(n))

def create_word_file(df):
    doc = Document()
    
    # Thiết lập font chữ mặc định (Times New Roman)
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(11)

    # Tiền xử lý dữ liệu
    df['NgaySinh'] = df.apply(lambda r: f"{int(r['Ngày']):02d}/{int(r['Tháng']):02d}/{int(r['năm'])}", axis=1)
    df['Cb_Short'] = df['CB'].replace("Binh nhì", "B2")
    df['CV_Short'] = df['CV'].replace("Chiến sĩ", "CS")
    
    # Gom nhóm theo Tỉnh
    tinh_groups = df.groupby('Tỉnh')
    
    for t_idx, (tinh_name, tinh_data) in enumerate(tinh_groups, 1):
        # I. Tỉnh Phú Thọ
        p_tinh = doc.add_paragraph(f"{format_roman(t_idx)}. Tỉnh {tinh_name}")
        p_tinh.runs[0].bold = True
        p_tinh.runs[0].font.size = Pt(13)
        
        # Gom nhóm theo Xã trong Tỉnh
        xa_groups = tinh_data.groupby('Xã')
        for x_idx, (xa_name, xa_data) in enumerate(xa_groups, 1):
            # 1. Xã Bao La
            p_xa = doc.add_paragraph(f"  {x_idx}. Xã {xa_name}")
            p_xa.runs[0].bold = True
            p_xa.runs[0].font.italic = True
            
            # Tạo bảng (12 cột dựa theo mẫu của bạn)
            headers = ["TT", "Họ và tên", "Ngày sinh", "Cb", "CV", "Đơn vị", "Nhập ngũ", "DT", "Văn hóa", "Quê quán", "Bố, Mẹ", "SĐT"]
            table = doc.add_table(rows=1, cols=len(headers))
            table.style = 'Table Grid'
            
            # Header bảng
            hdr_cells = table.rows[0].cells
            for i, h in enumerate(headers):
                hdr_cells[i].text = h
                paragraph = hdr_cells[i].paragraphs[0]
                run = paragraph.runs[0]
                run.font.bold = True
                run.font.size = Pt(10)
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # Điền dữ liệu quân nhân vào bảng
            for _, row in xa_data.iterrows():
                cells = table.add_row().cells
                cells[0].text = str(row['TT'])
                cells[1].text = str(row['Họ Và tên'])
                cells[2].text = str(row['NgaySinh'])
                cells[3].text = str(row['Cb_Short'])
                cells[4].text = str(row['CV_Short'])
                cells[5].text = str(row['ĐV'])
                cells[6].text = str(row['N.N'])
                cells[7].text = str(row['Dân tộc'])
                cells[8].text = str(row['Văn Hóa'])
                cells[9].text = f"{row['Xã']}-{row['Tỉnh']}"
                cells[10].text = f"{row['Bố']}, {row['Mẹ']}"
                cells[11].text = str(row['SDT gia đình'])
                
                # Chỉnh cỡ chữ nhỏ trong bảng cho gọn
                for cell in cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.size = Pt(9)

            doc.add_paragraph() # Ngắt dòng sau mỗi bảng xã

    # Lưu file vào bộ nhớ đệm
    target_stream = BytesIO()
    doc.save(target_stream)
    target_stream.seek(0)
    return target_stream

# Giao diện ứng dụng
st.title("🎖️ Công cụ Trích ngang Quân nhân")
st.write("Tải file Excel/CSV lên để tự động tạo file Word phân loại theo Tỉnh/Xã.")

uploaded_file = st.file_uploader("Chọn file dữ liệu (CSV/Excel)", type=["csv", "xlsx"])

if uploaded_file:
    try:
        # Đọc dữ liệu (Bỏ qua dòng metadata thứ 2 trong file của bạn)
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, skiprows=[1])
        else:
            df = pd.read_excel(uploaded_file, skiprows=[1])
            
        st.success("Tải dữ liệu thành công!")
        st.dataframe(df.head(10)) # Hiển thị 10 dòng đầu xem trước

        if st.button("🚀 Xuất file Word"):
            with st.spinner("Đang tạo bảng và phân loại..."):
                result_file = create_word_file(df)
                st.download_button(
                    label="📥 Tải xuống File Word kết quả",
                    data=result_file,
                    file_name="Tong_hop_trich_ngang_moi.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
    except Exception as e:
        st.error(f"Đã xảy ra lỗi: {e}")
