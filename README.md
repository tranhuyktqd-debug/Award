# 🎓 Hệ Thống Xử Lý Mã CERT ASMO

Ứng dụng quản lý và xử lý giải thưởng ASMO với giao diện đồ họa hiện đại.

## ✨ Tính năng

### 📋 Tab 1: Xử lý Mã Cert
- So sánh và tách file Awards_Template_Full.xlsx với file trao giải
- Xếp hạng thí sinh dựa trên điểm số
- Tự động tạo mã CERT cho các giải thưởng
- Tạo báo cáo thống kê chi tiết

### 📦 Tab 2: Chia danh sách
- Chia danh sách học sinh theo STT túi
- Chọn sheet cụ thể trong file Excel
- Tùy chỉnh các cột xuất ra
- Tự động format Excel cho in ấn (A4 Landscape)
- Thêm viền và tự động điều chỉnh độ rộng cột
- Xuất file với tên tự động theo sheet

### 🔍 Tab 3: Tra cứu
- Tìm kiếm học sinh theo SBD, Họ tên, Ngày sinh
- Hỗ trợ tra cứu từ nhiều sheet cùng lúc
- Hiển thị thông tin chi tiết: điểm số, chứng chỉ, ảnh, QR code
- Giao diện trực quan với màu sắc theo huy chương

## 🚀 Cài đặt

### Yêu cầu
- Python 3.7+
- pip

### Cài đặt thư viện

```bash
pip install pandas openpyxl qrcode[pil] pillow
```

## 📖 Hướng dẫn sử dụng

### Chạy ứng dụng

```bash
python awards_processing_app.py
```

Hoặc sử dụng file batch:

```bash
start_awards_app.bat
```

### Xử lý Mã Cert (Tab 1)
1. Chọn file đầy đủ (Awards_Template_Full.xlsx)
2. Chọn file trao giải (Awards_TRAO GIAI.xlsx)
3. Chọn thư mục lưu kết quả
4. Click "▶ BẮT ĐẦU XỬ LÝ"

### Chia danh sách (Tab 2)
1. Chọn file nguồn (Awards_Comparison_WITH_CERT.xlsx)
2. Chọn sheet cần chia
3. Tùy chỉnh các cột cần xuất
4. Click "📦 CHIA DANH SÁCH"
5. Lưu file với tên tự động

### Tra cứu (Tab 3)
1. Chọn file dữ liệu
2. Chọn các sheet cần tra cứu
3. Click "📥 TẢI DỮ LIỆU"
4. Tìm kiếm bằng SBD/Họ tên/Ngày sinh
5. Xem thông tin chi tiết

## 📁 Cấu trúc dự án

```
TEST_TRA_CUU_TRAO_GIAI/
├── awards_processing_app.py    # Ứng dụng chính
├── email_config.py              # Cấu hình email
├── email_server.py              # Server email
├── send_student_awards.py       # Gửi email hàng loạt
├── web_server.py                # Web server tra cứu
├── index.html                   # Giao diện web tra cứu
├── photos/                      # Ảnh thí sinh
├── QR/                          # Mã QR điểm danh
├── QR_SEAMO/                    # Mã QR SEAMO
├── templates/                   # Templates email
└── outputs/                     # Kết quả xuất ra
```

## 🔧 Tạo file .exe

### Sử dụng PyInstaller

```bash
pip install pyinstaller
pyinstaller --onedir --windowed --name="ASMO_Awards_Processing" awards_processing_app.py
```

File .exe sẽ nằm trong thư mục `dist/ASMO_Awards_Processing/`

### Sử dụng Auto-Py-to-Exe (Có giao diện)

```bash
pip install auto-py-to-exe
auto-py-to-exe
```

## 🛠️ Các script tiện ích

- `check_excel_structure.py` - Kiểm tra cấu trúc file Excel
- `check_qr_excel.py` - Kiểm tra QR code trong Excel
- `check_sbd_format.py` - Kiểm tra format SBD
- `create_qr_for_all_students.py` - Tạo QR cho tất cả học sinh
- `merge_qr_email.py` - Gộp QR và email

## 📝 Changelog

### Version 2.0 (Latest)
- ✅ Thêm Tab "Chia danh sách" với tùy chỉnh linh hoạt
- ✅ Thêm Tab "Tra cứu" với tìm kiếm đa tiêu chí
- ✅ Hỗ trợ chọn nhiều sheet cùng lúc
- ✅ Tự động format Excel cho in ấn
- ✅ Giao diện tab với màu sắc nổi bật
- ✅ Tối ưu layout và UX
- ✅ Xử lý lỗi PermissionError khi file đang mở
- ✅ Hiển thị ảnh và QR code trong tra cứu

### Version 1.0
- 🎯 Xử lý mã CERT cơ bản
- 📊 Tạo báo cáo thống kê
- 📧 Gửi email hàng loạt

## 📧 Liên hệ

- Email: support@asmo.vn
- Website: [ASMO Vietnam](https://asmo.vn)

## 📄 License

© 2026 ASMO Vietnam. All rights reserved.

---

**Developed with ❤️ for ASMO Vietnam**
