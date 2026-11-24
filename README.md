# Hướng Dẫn Sử Dụng Bộ Công Cụ Excel To PDF Converter Professional

Bộ công cụ chuyển đổi Excel sang PDF chuyên nghiệp dành cho hệ điều hành Windows, tích hợp **Microsoft Excel COM Automation** nguyên bản giúp chất lượng in ấn đạt độ hoàn hảo cao nhất.

---

## 1. Bản Đóng Gói Chạy Trực Tiếp (.EXE)

Không cần cài đặt Python. Bạn có thể sao chép file thực thi sau sang bất kỳ máy tính Windows nào (yêu cầu máy có cài Microsoft Excel):

* **Đường dẫn file**: `D:\Vinh.ph\AI\Antigravity\Excel_to_PDF\dist\ExcelToPDF_Professional.exe`

### Các điểm cải tiến vượt trội:
1. **Chất lượng file PDF xuất ra (Chỉnh chu, chuyên nghiệp)**:
   - **Tự động vừa vặn 1 trang ngang (Fit 1 Page Wide)**: Đảm bảo bảng tính Excel dù có nhiều cột cũng không bao giờ bị cắt vụn hay tràn 1-2 cột sang trang sau.
   - **Căn giữa trang theo chiều ngang (Center Horizontally)**: Nội dung bảng biểu luôn nằm cân đối ở giữa trang giấy A4, tạo cảm giác tài liệu báo cáo trang trọng.
   - **Tùy chọn Khổ giấy A4 / Letter**: Chuẩn hóa theo tiêu chuẩn văn phòng Việt Nam.
   - **Tùy chọn Hướng giấy linh hoạt**: Dọc (Portrait), Ngang (Landscape) hoặc Giữ nguyên theo file gốc.
   - **Tùy chọn linh hoạt**: Bạn hoàn toàn có thể bỏ tích *"Bật bộ tối ưu trang in chuyên nghiệp"* nếu file Excel của bạn đã được canh chỉnh Page Break thủ công cẩn thận từ trước.
2. **Độ ổn định tối đa (Dựa trên nền tảng `excel_to_pdf_windows`)**:
   - Tự động sao chép sang file tạm để xử lý, tuyệt đối **không khóa và không làm thay đổi nội dung file Excel gốc**.
   - Hỗ trợ vượt lỗi **Protected View / Mark-of-the-Web** (thường gặp khi tải file từ Zalo, Skype, Outlook hoặc trình duyệt).
   - Bọc toàn bộ các thuộc tính Excel COM trong khối an toàn, không bao giờ bị lỗi *"Property Visible can not be set"*.
   - Đảm bảo thoát và giải phóng hoàn toàn tiến trình `EXCEL.EXE` ngầm sau khi xuất xong.
3. **Giao diện người dùng hiện đại**:
   - Thiết kế theo phong cách Emerald Office (xanh ngọc lục bảo thanh lịch).
   - Hỗ trợ 3 chế độ xuất: Toàn bộ workbook, Từng sheet riêng biệt, hoặc Tích chọn các sheet cụ thể.
   - Tùy chọn nén tất cả các file PDF thành một tệp `.ZIP` gọn gàng.
   - Thanh tiến trình mượt mà kèm hộp nhật ký (log) chi tiết từng bước.

---

## 2. Cách Chạy Bằng Mã Nguồn Python

Nếu bạn muốn tùy biến mã nguồn hoặc chạy trực tiếp từ Python:

### Chạy Bản Desktop GUI (Tkinter):
```cmd
python excel_to_pdf_converter.py
```
*(hoặc `py excel_to_pdf_converter.py`)*

### Chạy Bản Web GUI (Streamlit) & Chia Sẻ Cho Các Máy Tính Khác:

#### 1. Dành cho máy tính trong cùng mạng Wi-Fi / Mạng nội bộ (LAN):
Mở trình duyệt trên máy tính khác và nhập địa chỉ IP mạng nội bộ của máy bạn:
```
http://192.168.88.213:8501
```

#### 2. Dành cho máy tính bất kỳ (Qua Internet / Ngoài công ty / Di động):
Bạn có thể tạo một đường dẫn HTTPS công khai an toàn (không cần cấu hình modem hay mở port) bằng cách nhấp đúp vào:
- **`Chia_Se_Link_Online.bat`** (sử dụng công nghệ Cloudflare Tunnel bảo mật cao).
- Đường link có dạng: `https://[ten-ngau-nhien].trycloudflare.com` gửi cho bất kỳ ai trên thế giới đều dùng được ngay lập tức!


### Tính Năng Mới Trên Bản Web Streamlit:
- **Xem trước trực tiếp file PDF trên web**: Tích hợp cả Trình xem tương tác (Interactive PDF Viewer - cuộn, phóng to, in) và Thư viện ảnh từng trang (Full-page gallery) chất lượng cao.
- **Tùy chọn chọn từng Sheet**: Tự động phát hiện danh sách sheet trong từng file và cho phép tick chọn xuất các sheet mong muốn.
- **Tối ưu in ấn chuyên nghiệp**: Fit 1 trang ngang (chống đứt cột), Căn giữa trang, Khổ giấy A4/Letter, Hướng dọc/ngang, Căn lề chuẩn/hẹp.
- **Tải về linh hoạt**: Tải lẻ từng file PDF hoặc đóng gói toàn bộ vào file `.ZIP` chỉ với 1 cú nhấp chuột.


---

## 3. Lệnh Đóng Gói Lại Bằng PyInstaller (Nếu Cần Cập Nhật)

```cmd
python -m PyInstaller excel_to_pdf_pro.spec
```
File thực thi độc lập sẽ tự động được sinh ra trong thư mục `dist\ExcelToPDF_Professional.exe`.
