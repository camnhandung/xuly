import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from io import BytesIO

# Cấu hình trang
st.set_page_config(page_title="Tự động hóa Trích ngang", layout="wide")

def set_cell_border(cell):
    """Hàm phụ trợ để kẻ bảng cho Word"""
    from docx.oxml import OxmlElement
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for border in ['top', 'left', 'bottom', 'right']:
        edge = OxmlElement(f'w:{border}')
        edge.set(qn('w:val'), 'single')
        edge.set(qn('w:sz'), '4')
        edge.set(qn('w:color'), '000000')
        tcPr.append(edge)

def create_word_file(df):
    doc = Document()
    
    # Thiết lập font chữ mặc định là Times New Roman
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(11)

    # 1. Tiền xử lý dữ liệu
    # Kết hợp Ngày/Tháng/Năm sinh
    df['NgaySinh'] = df.apply(lambda r: f"{int(r['Ngày']):02d}/{int(r['Tháng']):02d}/{int(r['năm'])}", axis=1)
    # Ánh xạ Cấp bậc & Chức vụ (Binh nhì -> B2, Chiến sĩ -> CS)
    df['Cb_Short'] = df['CB'].replace("Binh nhì", "B2")
    df['CV_Short'] = df['CV'].replace("Chiến sĩ", "CS")
    
    # 2. Gom nhóm theo Tỉnh, sau đó theo Xã
    tinh_groups = df.groupby('Tỉnh')
    
    tinh_roman = ["I", "II", "III", "IV", "V", "VI"] # Có thể mở rộng thêm
    
    for t_idx, (tinh_name, tinh_data) in enumerate(tinh_groups):
        # Tiêu đề Tỉnh (Ví dụ: I. Tỉnh Điện Biên)
        t_label = tinh_roman[t_idx] if t_idx < len(tinh_roman) else str(t_idx + 1)
        p_tinh = doc.add_paragraph(f"{t_label}. Tỉnh {tinh_name}")
        p_tinh.runs[0].bold = True
        p_tinh.runs[0].font.size = Pt(13)
        
        xa_groups = tinh_data.groupby('Xã')
        for x_idx, (xa_name, xa_data) in enumerate(xa_groups):
            # Tiêu đề Xã (Ví dụ: 1. Xã Chà Tở)
            p_xa = doc.add_paragraph(f"{x_idx + 1}. Xã {xa_name}")
            p_xa.runs[0].bold = True
            p_xa.runs[0].font.italic = True
            
            # Tạo bảng theo các cột trong file Word mẫu
            headers = [
                "TT", "Họ và tên", "Ngày tháng năm sinh", "Cb", "CV", "Đơn vị", 
                "Nhập ngũ", "DT", "Văn hóa", "Quê quán", "Họ tên bố, mẹ", "SĐT"
            ]
            
            table = doc.add_table(rows=1, cols=len(headers))
            table.style = 'Table Grid'
            
            # Header bảng
            hdr_cells = table.rows[0].cells
            for i, h in enumerate(headers):
                hdr_cells[i].text = h
                # Định dạng chữ đậm cho Header
                paragraph = hdr_cells[i].paragraphs[0]
                run = paragraph.runs[0]
                run.font.bold = True
                run.font.size = Pt(10)
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # Điền dữ liệu từng người
            for row_idx, row in xa_data.iterrows():
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
                
                # Chỉnh cỡ chữ nhỏ cho nội dung bảng
                for cell in cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.size = Pt(9)

            doc.add_paragraph() # Khoảng trống giữa các xã

    # Lưu vào buffer để tải về
    target_stream = BytesIO()
    doc.save(target_stream)
    target_stream.seek(0)
    return target_stream

# GIAO DIỆN STREAMLIT
st.title("🚀 Công cụ tự động điền danh sách Trích ngang")
st.info("Hướng dẫn: Tải file Excel/CSV lên, hệ thống sẽ tự tạo danh sách theo Tỉnh và Xã.")

uploaded_file = st.file_uploader("Chọn file CSV/Excel", type=["csv", "xlsx"])

if uploaded_file is not None:
    try:
        # Đọc dữ liệu (Bỏ qua dòng metadata thứ 2 của bạn)
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, skiprows=[1])
        else:
            df = pd.read_excel(uploaded_file, skiprows=[1])
            
        st.subheader("Xem trước dữ liệu gốc")
        st.write(df.head())

        if st.button("Tạo file Word ngay"):
            with st.spinner('Đang tạo bảng và phân loại theo Tỉnh/Xã...'):
                doc_file = create_word_file(df)
                st.success("Đã xử lý xong!")
                st.download_button(
                    label="📥 Tải file Word kết quả",
                    data=doc_file,
                    file_name="Tong_hop_trich_ngang_KQX.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
    except Exception as e:
        st.error(f"Lỗi khi đọc file: {e}. Vui lòng kiểm tra lại định dạng cột.")
