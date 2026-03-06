import streamlit as st
import pandas as pd
from docx import Document
import io

st.set_page_config(page_title="Fix Lỗi Trích Ngang", page_icon="🛠️")

st.title("🛠️ Sửa lỗi 'tuple index out of range'")
st.info("Lưu ý: Đảm bảo file Excel có tiêu đề ở dòng 8 & 9 và file Word có bảng đủ 21 cột.")

excel_file = st.file_uploader("📥 Tải lên file Excel (.xlsx)", type=["xlsx"])
word_file = st.file_uploader("📥 Tải lên file Word mẫu (.docx)", type=["docx"])

if excel_file and word_file:
    if st.button("🚀 Chạy lại chương trình"):
        try:
            # 1. Đọc Excel với xử lý lỗi Header
            df = pd.read_excel(excel_file, skiprows=7) 
            # Lấy dòng tiêu đề phụ (dòng 8 cũ) để gộp tên cột nếu cần
            # Nhưng để đơn giản và tránh lỗi tuple, ta sẽ map thủ công theo vị trí cột
            
            doc = Document(word_file)
            if not doc.tables:
                st.error("File Word không có bảng nào!")
                st.stop()
            
            table = doc.tables[0]
            num_cols_in_word = len(table.columns)

            for index, row in df.iterrows():
                # Kiểm tra nếu dòng trống (cột Họ tên ở vị trí index 1)
                if pd.isna(row.iloc[1]): continue

                # Trích xuất dữ liệu dựa trên vị trí cột (iloc) để tránh lỗi tên cột
                ho_ten = str(row.iloc[1])
                dv = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""
                cap_bac = str(row.iloc[3]) if pd.notna(row.iloc[3]) else ""
                chuc_vu = str(row.iloc[4]) if pd.notna(row.iloc[4]) else ""
                nhap_ngu = str(row.iloc[5]) if pd.notna(row.iloc[5]) else ""
                dan_toc = str(row.iloc[6]) if pd.notna(row.iloc[6]) else ""
                van_hoa = str(row.iloc[7]) if pd.notna(row.iloc[7]) else ""
                cccd = str(row.iloc[8]).replace(".0", "") if pd.notna(row.iloc[8]) else ""
                
                # Ngày tháng năm (Cột 9, 10, 11)
                ngay = str(row.iloc[9]).split('.')[0] if pd.notna(row.iloc[9]) else ""
                thang = str(row.iloc[10]).split('.')[0] if pd.notna(row.iloc[10]) else ""
                nam = str(row.iloc[11]).split('.')[0] if pd.notna(row.iloc[11]) else ""
                ngay_sinh = f"{ngay}/{thang}/{nam}"

                # Quê quán (Cột 12, 13)
                que = f"{row.iloc[12]} - {row.iloc[13]}"
                
                # Gia đình (Cột 16, 17, 18)
                bo = str(row.iloc[16]) if pd.notna(row.iloc[16]) else ""
                me = str(row.iloc[17]) if pd.notna(row.iloc[17]) else ""
                sdt = str(row.iloc[18]).replace(".0", "") if pd.notna(row.iloc[18]) else ""

                # --- ĐIỀN VÀO WORD ---
                new_row = table.add_row().cells
                
                # Hàm an toàn để điền dữ liệu (tránh out of range)
                def safe_fill(idx, text):
                    if idx < num_cols_in_word:
                        new_row[idx].text = str(text)

                safe_fill(0, str(index + 1))
                safe_fill(1, ho_ten)
                safe_fill(2, ngay_sinh)
                safe_fill(3, cap_bac)
                safe_fill(4, chuc_vu)
                safe_fill(5, dv)
                safe_fill(6, nhap_ngu)
                safe_fill(8, van_hoa)
                safe_fill(10, dan_toc)
                safe_fill(14, que)
                safe_fill(16, f"{bo}\n{me}")
                safe_fill(18, sdt)
                safe_fill(19, cccd)

            output = io.BytesIO()
            doc.save(output)
            output.seek(0)

            st.success("✅ Thành công!")
            st.download_button("⬇️ Tải file", output, "Ket_qua.docx")

        except Exception as e:
            st.error(f"❌ Lỗi chi tiết: {e}")
            st.write("Gợi ý: Kiểm tra xem file Word của bạn có đủ số cột như file mẫu không.")
