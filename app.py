import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

def convert_to_roman(num):
    val = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
    syb = ["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"]
    roman_num = ''
    i = 0
    while num > 0:
        for _ in range(num // val[i]):
            roman_num += syb[i]
            num -= val[i]
        i += 1
    return roman_num

def generate_word_doc(df):
    # Tạo document và thiết lập khổ giấy ngang (Landscape) để vừa nhiều cột
    doc = Document()
    section = doc.sections[-1]
    new_width, new_height = section.page_height, section.page_width
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = new_width
    section.page_height = new_height
    
    # Tiêu đề
    title = doc.add_paragraph("TỔNG HỢP TRÍCH NGANG")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.bold = True
        run.font.size = Pt(14)
    
    # Các cột trong bảng Word mẫu
    headers = [
        "TT", "Họ và tên", "Ngày tháng năm sinh", "Cb", "CV", "Đơn vị", 
        "Nhập ngũ", "Thành phần", "Văn hóa", "Sức khỏe", "DT", "TG", 
        "Ngày vào Đoàn", "Ngày vào Đảng", "Quê quán", "Trú Quán", 
        "Họ tên bố, mẹ", "Họ tên vợ, con", "SỐ ĐIỆN THOẠI", "Số cccd", "Ghi chú"
    ]
    
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    
    # Điền header
    hdr_cells = table.rows[0].cells
    for i, header_text in enumerate(headers):
        hdr_cells[i].text = header_text
        hdr_cells[i].paragraphs[0].runs[0].bold = True
        hdr_cells[i].paragraphs[0].runs[0].font.size = Pt(9)

    # Loại bỏ dữ liệu trống và thay thế NaN
    df = df.dropna(subset=[1]) # Drop nếu không có Tên
    df = df.fillna("")

    # Nhóm dữ liệu theo Tỉnh và Xã
    # Tỉnh ở cột index 13, Xã ở cột index 12
    grouped_tinh = df.groupby(13)
    
    tinh_counter = 1
    for tinh, group_tinh in grouped_tinh:
        if str(tinh).strip() == "": continue
        
        # Thêm dòng Tỉnh
        row_tinh = table.add_row().cells
        row_tinh[0].merge(row_tinh[-1])
        row_tinh[0].text = f"{convert_to_roman(tinh_counter)}. Tỉnh {tinh}"
        row_tinh[0].paragraphs[0].runs[0].bold = True
        tinh_counter += 1
        
        grouped_xa = group_tinh.groupby(12)
        xa_counter = 1
        for xa, group_xa in grouped_xa:
            if str(xa).strip() == "": continue
            
            # Thêm dòng Xã
            row_xa = table.add_row().cells
            row_xa[0].merge(row_xa[-1])
            row_xa[0].text = f"{xa_counter}. Xã {xa}"
            row_xa[0].paragraphs[0].runs[0].bold = True
            xa_counter += 1
            
            # Thêm chi tiết từng người
            for index, row_data in group_xa.iterrows():
                row_cells = table.add_row().cells
                
                # Trích xuất dữ liệu dựa theo index cột Excel
                ho_ten = str(row_data[1])
                dv = str(row_data[2])
                cb = str(row_data[3]).replace("Binh nhì", "B2") # Có thể tùy chỉnh mapping
                cv = str(row_data[4]).replace("Chiến sĩ", "CS")
                nn = str(row_data[5])
                dt = str(row_data[6])
                vh = str(row_data[7])
                cccd = str(row_data[8])
                
                # Xử lý ngày sinh
                ngay, thang, nam = str(row_data[9]), str(row_data[10]), str(row_data[11])
                ngay_sinh = f"{ngay.split('.')[0]}/{thang.split('.')[0]}/{nam.split('.')[0]}" if ngay else ""
                
                bo, me = str(row_data[16]), str(row_data[17])
                bo_me = f"{bo}, {me}".strip(", ")
                
                sdt = str(row_data[18]).split('.')[0] # Bỏ số .0 nếu có ở cuối
                que_quan = f"{xa}-{tinh}"

                # Điền vào Word
                row_cells[0].text = "" # Cột TT để trống hoặc tự đánh số
                row_cells[1].text = ho_ten
                row_cells[2].text = ngay_sinh
                row_cells[3].text = cb
                row_cells[4].text = cv
                row_cells[5].text = dv
                row_cells[6].text = nn
                row_cells[7].text = "BN" # Thành phần (bạn có thể đổi theo ý)
                row_cells[8].text = vh
                row_cells[9].text = "" # Sức khỏe
                row_cells[10].text = dt
                row_cells[11].text = "" # TG
                row_cells[12].text = "" # Ngày vào đoàn
                row_cells[13].text = "" # Ngày vào Đảng
                row_cells[14].text = que_quan
                row_cells[15].text = "" # Trú quán
                row_cells[16].text = bo_me
                row_cells[17].text = "" # Vợ con
                row_cells[18].text = sdt
                row_cells[19].text = cccd
                row_cells[20].text = "" # Ghi chú
                
                # Chỉnh font size nhỏ cho vừa bảng
                for cell in row_cells:
                    if cell.text:
                        for paragraph in cell.paragraphs:
                            for run in paragraph.runs:
                                run.font.size = Pt(9)

    # Lưu file ra byte stream để tải xuống
    byte_io = io.BytesIO()
    doc.save(byte_io)
    byte_io.seek(0)
    return byte_io

# Giao diện Streamlit
st.set_page_config(page_title="Tạo File Trích Ngang", layout="wide")
st.title("Phần mềm tạo danh sách Trích Ngang tự động")
st.write("Tải lên file Excel dữ liệu danh sách để tự động nhóm theo Tỉnh/Xã và xuất ra file Word.")

uploaded_file = st.file_uploader("Chọn file Excel / CSV của bạn", type=['xlsx', 'csv'])

if uploaded_file is not None:
    try:
        # Bỏ qua 2 dòng đầu tiên (dòng tiêu đề chính và tiêu đề phụ) để lấy trực tiếp dữ liệu
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, skiprows=2, header=None)
        else:
            df = pd.read_excel(uploaded_file, skiprows=2, header=None)
            
        st.success("Đã đọc file thành công! Preview dữ liệu thô:")
        st.dataframe(df.head(3))
        
        if st.button("Tạo File Word"):
            with st.spinner("Đang xử lý và tạo file..."):
                word_file = generate_word_doc(df)
                
                st.download_button(
                    label="📥 Tải xuống File Word (.docx)",
                    data=word_file,
                    file_name="Tong_hop_trich_ngang_ket_qua.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
    except Exception as e:
        st.error(f"Có lỗi xảy ra khi xử lý file: {e}")
