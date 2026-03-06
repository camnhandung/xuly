import streamlit as st
import pandas as pd
from io import BytesIO

# Cấu hình giao diện
st.set_page_config(page_title="Trình gộp dữ liệu thông minh", layout="wide")

st.title("📂 Tự động tìm và gộp dữ liệu Bố Mẹ")
st.info("Ứng dụng sẽ tự quét các cột: Bố, Mẹ, Số điện thoại để gộp thành định dạng yêu cầu.")

uploaded_file = st.file_uploader("Tải file Excel của bạn lên đây", type=["xlsx", "xls"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        cols = df.columns.tolist()
        
        # --- THUẬT TOÁN TỰ TÌM CỘT ---
        col_bo = next((c for c in cols if "bố" in str(c).lower() or "cha" in str(c).lower()), None)
        col_me = next((c for c in cols if "mẹ" in str(c).lower()), None)
        col_sdt = next((c for c in cols if any(k in str(c).lower() for k in ["sđt", "sdt", "điện thoại", "phone"])), None)

        # Hiển thị kết quả tìm kiếm cột để người dùng xác nhận
        st.subheader("🔍 Kết quả quét các cột tương ứng:")
        c1, c2, c3 = st.columns(3)
        with c1: selected_bo = st.selectbox("Cột Bố:", cols, index=cols.index(col_bo) if col_bo else 0)
        with c2: selected_me = st.selectbox("Cột Mẹ:", cols, index=cols.index(col_me) if col_me else 0)
        with c3: selected_sdt = st.selectbox("Cột SĐT:", cols, index=cols.index(col_sdt) if col_sdt else 0)

        if st.button("🚀 Tiến hành gộp và tạo file mới"):
            # Hàm xử lý logic gộp dữ liệu
            def process_row(row):
                val_bo = str(row[selected_bo]).strip() if pd.notna(row[selected_bo]) else ""
                val_me = str(row[selected_me]).strip() if pd.notna(row[selected_me]) else ""
                val_sdt = str(row[selected_sdt]).strip() if pd.notna(row[selected_sdt]) else ""
                
                # Xử lý số điện thoại (tránh lỗi mất số 0 hoặc định dạng khoa học)
                if val_sdt.endswith('.0'): val_sdt = val_sdt[:-2]
                if val_sdt and not val_sdt.startswith('0'): val_sdt = '0' + val_sdt
                
                # Gộp theo định dạng ảnh 1 (dùng \n để xuống dòng trong 1 ô)
                return f"{val_bo}\n{val_me}\n{val_sdt}"

            # Tạo DataFrame kết quả
            new_column_name = "Họ tên bố, mẹ\nHọ tên vợ, con\nSỐ ĐIỆN THOẠI\nGIA ĐÌNH"
            result_df = pd.DataFrame({new_column_name: df.apply(process_row, axis=1)})

            # Hiển thị bản xem trước
            st.success("Đã tạo danh sách thành công!")
            st.dataframe(result_df, height=400)

            # Xuất file Excel có format Wrap Text (xuống dòng)
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                result_df.to_excel(writer, index=False, sheet_name='Danh_Sach_Gop')
                
                workbook  = writer.book
                worksheet = writer.sheets['Danh_Sach_Gop']
                
                # Format này rất quan trọng để hiển thị giống Ảnh 1
                cell_format = workbook.add_format({
                    'text_wrap': True, 
                    'valign': 'vcenter', 
                    'align': 'left',
                    'border': 1
                })
                
                # Độ rộng cột (khoảng 40)
                worksheet.set_column('A:A', 40, cell_format)
                
            st.download_button(
                label="📥 Tải file kết quả về máy",
                data=output.getvalue(),
                file_name="Danh_sach_gia_dinh_gop.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Lỗi: {e}. Vui lòng kiểm tra lại cấu trúc file.")
