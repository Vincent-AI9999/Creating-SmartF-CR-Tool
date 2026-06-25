# Mobifone Change Request (CR) Automation Tool

Công cụ tự động hóa quy trình tạo Change Request (CR) đồng bộ tham số mạng vô tuyến Nokia và Ericsson cho Mobifone.

## Tính năng chính

1. **Vendor & RAN Selections**: Hỗ trợ chọn **Nokia / Ericsson** và các công nghệ **3G / 4G / 5G**.
2. **Gợi ý KPI kém**: Tự động gợi ý danh sách Cell lỗi từ file báo cáo KPI SmartF mới nhất (`Cell_Analysis_*.xlsx` trong thư mục `Source_cell`).
3. **Đồng bộ tham số**: Tự động truy vấn giá trị tham số hiện tại (Old) từ file dump mạng lưới tương ứng và điền giá trị tối ưu hóa (New) dựa theo mẫu.
4. **Giữ nguyên định dạng Excel**: Sao chép nguyên vẹn font chữ, viền, màu sắc ô từ CR mẫu sang CR kết quả.
5. **Cập nhật thông tin phụ**: Tự động điền lại các thông tin ở sheet `Trạm ảnh hưởng` và `LNCELL` dựa trên danh sách cell tác động.

## Cài đặt & Chuẩn bị

### 1. Cài đặt các thư viện cần thiết
Mở terminal và chạy lệnh:
```bash
pip install streamlit pandas openpyxl numpy xlsxwriter
```

### 2. Cấu trúc thư mục dữ liệu
Ứng dụng sử dụng cấu trúc thư mục sau để đọc dữ liệu (đã được cấu hình cứng trong code):
- **Thư mục chứa CR mẫu**: `F:\OneDrive - Mobifone\F_WORKING\Python_Coding\Create CR\CR mẫu`
- **Thư mục chứa KPI cell**: `F:\OneDrive - Mobifone\F_WORKING\Python_Coding\MTCL 2026\Source_cell`
- **Thư mục chứa dump trạm**: `F:\OneDrive - Mobifone\F_WORKING\Python_Coding\Database trạm`

## Hướng dẫn sử dụng

### Cách 1: Chạy trực tiếp từ file Batch (Khuyên dùng)
- Nhấp đúp chuột vào file **`Run_CR_Generator.bat`**.
- Trình duyệt sẽ tự động mở giao diện web tại địa chỉ `http://localhost:8501`.

### Cách 2: Chạy bằng dòng lệnh
Mở terminal tại thư mục dự án và chạy:
```bash
streamlit run create_cr_app.py
```
hoặc
```bash
python -m streamlit run create_cr_app.py
```
