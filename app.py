import streamlit as st
import pandas as pd
from docx import Document
import io

# Cấu hình trang
st.set_page_config(page_title="Tạo Danh Sách Trích Ngang", page_icon="📝")

st.title("📝 Ứng dụng Điền dữ liệu từ Excel vào Word")
st.markdown("Tải lên file Excel chứa dữ liệu và file Word mẫu để tự động tạo danh sách.")

# Tạo giao diện upload file
col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("📥 Tải lên file Excel (.xlsx)", type=["xlsx"])
with col2:
    word_file = st.file_uploader("📥 Tải lên file Word mẫu (.docx)", type=["docx"])

if excel_file is not None and word_file is not None:
    if st.button("🚀 Bắt đầu tạo file Word", use_container_width=True):
        try:
            with st.spinner('Đang xử lý dữ liệu...'):
                # 1. Đọc dữ liệu từ file Excel tải lên
                df = pd.read_excel(excel_file, skiprows=7, header=[0, 1])
                
                # Làm phẳng các cột MultiIndex của Pandas
                df.columns = [f"{col[0]}_{col[1]}" if not 'Unnamed' in str(col[1]) else col[0] for col in df.columns]

                # Lọc bỏ các dòng trống (dựa vào cột Họ và tên)
                if 'Họ Và tên' in df.columns:
                    df = df.dropna(subset=['Họ Và tên'])
                else:
                    st.error("Không tìm thấy cột 'Họ Và tên' ở dòng 8 trong file Excel. Vui lòng kiểm tra lại định dạng.")
                    st.stop()

                # 2. Đọc file Word mẫu tải lên
                doc = Document(word_file)
                table = doc.tables[0]
                
                # 3. Lặp và điền dữ liệu
                for index, row in df.iterrows():
                    ho_ten = str(row['Họ Và tên']) if pd.notna(row['Họ Và tên']) else ""
                    
                    # Xử lý ngày tháng năm sinh
                    ngay = str(int(row['Ngày tháng năm sinh_Ngày'])) if pd.notna(row['Ngày tháng năm sinh_Ngày']) else ""
                    thang = str(int(row['Ngày tháng năm sinh_Tháng'])) if pd.notna(row['Ngày tháng năm sinh_Tháng']) else ""
                    nam = str(int(row['Ngày tháng năm sinh_năm'])) if pd.notna(row['Ngày tháng năm sinh_năm']) else ""
                    ngay_sinh = f"{ngay}/{thang}/{nam}" if nam else ""
                    
                    cap_bac = str(row['CB']) if pd.notna(row['CB']) else ""
                    chuc_vu = str(row['CV']) if pd.notna(row['CV']) else ""
                    don_vi = str(row['ĐV']) if pd.notna(row['ĐV']) else ""
                    nhap_ngu = str(row['N.N']) if pd.notna(row['N.N']) else ""
                    dan_toc = str(row['Dân tộc']) if pd.notna(row['Dân tộc']) else ""
                    van_hoa = str(row['Văn Hóa']) if pd.notna(row['Văn Hóa']) else ""
                    so_cccd = str(row['Số CCCD']).replace(".0", "") if pd.notna(row['Số CCCD']) else ""
                    
                    # Xử lý quê quán
                    xa = str(row['Địa phương xuất ngũ_Xã']) if pd.notna(row['Địa phương xuất ngũ_Xã']) else ""
                    tinh = str(row['Địa phương xuất ngũ_Tỉnh']) if pd.notna(row['Địa phương xuất ngũ_Tỉnh']) else ""
                    que_quan = f"{xa} - {tinh}" if xa and tinh else (xa or tinh)
                    
                    # Xử lý thông tin gia đình
                    bo = str(row['Bố']) if pd.notna(row['Bố']) else ""
                    me = str(row['Mẹ']) if pd.notna(row['Mẹ']) else ""
                    ho_ten_bo_me = f"{bo}\n{me}".strip()
                    sdt = str(row['SDT gia đình']).replace(".0", "") if pd.notna(row['SDT gia đình']) else ""

                    # Thêm dòng mới vào Word
                    row_cells = table.add_row().cells
                    row_cells[0].text = str(index + 1)
                    row_cells[1].text = ho_ten
                    row_cells[2].text = ngay_sinh
                    row_cells[3].text = cap_bac
                    row_cells[4].text = chuc_vu
                    row_cells[5].text = don_vi
                    row_cells[6].text = nhap_ngu
                    row_cells[7].text = "BN" 
                    row_cells[8].text = van_hoa
                    row_cells[9].text = "" 
                    row_cells[10].text = dan_toc
                    row_cells[11].text = "" 
                    row_cells[12].text = "" 
                    row_cells[13].text = "" 
                    row_cells[14].text = que_quan
                    row_cells[15].text = "" 
                    row_cells[16].text = ho_ten_bo_me
                    row_cells[17].text = "" 
                    row_cells[18].text = sdt
                    row_cells[19].text = so_cccd
                    row_cells[20].text = "" 

                # 4. Lưu kết quả vào bộ nhớ đệm (BytesIO) thay vì lưu vào ổ cứng
                output = io.BytesIO()
                doc.save(output)
                output.seek(0) # Quay lại đầu file để chuẩn bị download

            st.success("✅ Đã tạo file thành công! Bạn có thể tải về ngay bên dưới.")
            
            # Hiển thị nút Download
            st.download_button(
                label="⬇️ Tải file Kết quả (.docx)",
                data=output,
                file_name="Tong_hop_trich_ngang_Hoan_thanh.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        except Exception as e:
            st.error(f"❌ Có lỗi xảy ra trong quá trình xử lý: {e}")