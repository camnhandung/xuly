import streamlit as st
import pandas as pd
from io import BytesIO

# Cấu hình trang
st.set_page_config(page_title="Gộp Dữ Liệu Gia Đình", layout="centered")

st.title("🛠 Công cụ gộp thông tin Bố Mẹ & SĐT")
st.write("Tải lên file Excel có định dạng như ảnh 2 để chuyển đổi sang định dạng ảnh 1.")

# 1. Tải file lên
uploaded_file = st.file_uploader("Chọn file Excel (.xlsx hoặc .xls)", type=["xlsx", "xls"])

if uploaded_file:
    try:
        # Đọc dữ liệu
        df = pd.read_excel(uploaded_file)
        
        st.subheader("Dữ liệu gốc (Xem trước)")
        st.dataframe(df.head())

        # Kiểm tra sự tồn tại của các cột (tùy chỉnh tên cột nếu file của bạn khác)
        # Ở đây tôi mặc định lấy theo ảnh: 'Bố', 'Mẹ', 'SDT gia đình'
        columns = df.columns.tolist()
        
        st.info(f"Các cột tìm thấy: {', '.join(columns)}")

        # Cho phép người dùng chọn cột nếu tên không khớp hoàn toàn
        col_bo = st.selectbox("Chọn cột Họ tên Bố:", columns, index=0 if 'Bố' in columns else 0)
        col_me = st.selectbox("Chọn cột Họ tên Mẹ:", columns, index=1 if 'Mẹ' in columns else 0)
        col_sdt = st.selectbox("Chọn cột Số điện thoại:", columns, index=2 if 'SDT' in columns[2] or 'SĐT' in columns[2] else 0)

        if st.button("Bắt đầu xử lý và gộp dữ liệu"):
            # 2. Xử lý gộp dữ liệu
            # Hàm gộp: Tên Bố \n Tên Mẹ \n Số điện thoại
            def combine_info(row):
                bo = str(row[col_bo]).strip() if pd.notna(row[col_bo]) else ""
                me = str(row[col_me]).strip() if pd.notna(row[col_me]) else ""
                sdt = str(row[col_sdt]).strip() if pd.notna(row[col_sdt]) else ""
                
                # Định dạng số điện thoại nếu bị biến thành số thực (ví dụ 3.75e+08)
                if sdt.endswith('.0'): sdt = sdt[:-2]
                if len(sdt) > 0 and not sdt.startswith('0'): sdt = '0' + sdt

                return f"{bo}\n{me}\n{sdt}"

            # Tạo DataFrame mới chỉ có 1 cột kết quả
            output_df = pd.DataFrame()
            output_df['Họ tên bố, mẹ\nHọ tên vợ, con\nSỐ ĐIỆN THOẠI\nGIA ĐÌNH'] = df.apply(combine_info, axis=1)

            st.success("Đã xử lý xong!")
            st.subheader("Dữ liệu sau khi gộp")
            st.dataframe(output_df)

            # 3. Xuất file Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                output_df.to_excel(writer, index=False, sheet_name='Sheet1')
                
                # Định dạng tự động xuống dòng trong Excel (Wrap Text)
                workbook  = writer.book
                worksheet = writer.sheets['Sheet1']
                wrap_format = workbook.add_format({'text_wrap': True, 'align': 'center', 'valign': 'vcenter'})
                
                # Áp dụng định dạng cho cột A và đặt độ rộng
                worksheet.set_column('A:A', 35, wrap_format)

            processed_data = output.getvalue()

            st.download_button(
                label="📥 Tải xuống file Excel đã gộp",
                data=processed_data,
                file_name="Gia_Dinh_Da_Gop.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Đã có lỗi xảy ra: {e}")

---
### Hướng dẫn sử dụng:

1.  **Cài đặt thư viện cần thiết:**
    Chạy lệnh sau trong terminal của bạn:
    ```bash
    pip install streamlit pandas openpyxl xlsxwriter
    ```
2.  **Chạy ứng dụng:**
    Lưu đoạn mã trên vào file `app.py` và chạy:
    ```bash
    streamlit run app.py
    ```
3.  **Cách thức hoạt động:**
    * Ứng dụng sẽ đọc file Excel của bạn (Ảnh 2).
    * Nó sẽ tạo ra một cột mới, gộp tên Bố, tên Mẹ và SĐT lại, ngăn cách bằng dấu xuống dòng (`\n`).
    * Khi tải xuống, file Excel sẽ được tự động bật tính năng **Wrap Text** (Tự động xuống dòng) để hiển thị giống hệt như Ảnh 1.

**Lưu ý nhỏ:** Nếu số điện thoại trong file Excel của bạn bị mất số `0` ở đầu (do Excel hiểu là định dạng số), đoạn mã trên đã có hàm tự động bù lại số `0` cho bạn.

Bạn có muốn tôi chỉnh sửa thêm về giao diện hay thêm cột nào khác vào phần gộp không?
