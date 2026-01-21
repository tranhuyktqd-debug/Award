# 📋 CHANGELOG - Cập Nhật index.html và script.js

**Ngày:** 2026-01-19
**Tác vụ:** Cập nhật logic tra cứu học sinh cho khớp với cấu trúc file Awards_Comparison_WITH_CERT.xlsx

---

## 🔍 VẤN ĐỀ ĐÃ PHÁT HIỆN

File Excel `Awards_Comparison_WITH_CERT.xlsx` có cấu trúc cột KHÁC với những gì code đang expect:

### CẤU TRÚC CŨ (trong code):
- `D.O.B` → Ngày sinh
- `TOÁN` → Kết quả Toán
- `KHOA HỌC` → Kết quả Khoa học
- `TIẾNG ANH` → Kết quả Tiếng Anh
- `CERT CODE FULL` → Mã cert đầy đủ
- `CERT CODE` → Mã cert rút gọn
- `SĐT` → Số điện thoại
- `EMAIL` → Email

### CẤU TRÚC MỚI (trong Excel):
- `Ngày sinh` → D.O.B
- `KQ VQG TOÁN` → Kết quả Toán
- `KQ VQG KHOA HỌC` → Kết quả Khoa học
- `KQ VQG TIẾNG ANH` → Kết quả Tiếng Anh
- `MÃ CERT ĐẦY ĐỦ` → Mã cert đầy đủ
- `MÃ CERT` → Mã cert rút gọn
- `Số điện thoại liên hệ` → SĐT
- `Email liên hệ` → Email
- **KHÔNG CÓ** cột `KHU VỰC` (Area)

---

## ✅ CÁC THAY ĐỔI ĐÃ THỰC HIỆN

### 1. **File: `script.js`**

#### a) Cập nhật hàm đọc dữ liệu từ Excel (dòng 94-114):
```javascript
// CŨ:
dob: row['D.O.B'] || row['D.O.B2'] || row['DOB'] || row['Ngày sinh'] || '',
toan: row['TOÁN'] || row['Toán'] || '',
kh: row['KHOA HỌC'] || row['Khoa học'] || row['KH'] || '',
ta: row['TIẾNG ANH'] || row['Tiếng Anh'] || row['TA'] || '',

// MỚI:
dob: row['Ngày sinh'] || row['D.O.B'] || row['D.O.B2'] || row['DOB'] || '',
toan: row['KQ VQG TOÁN'] || row['TOÁN'] || row['Toán'] || '',
kh: row['KQ VQG KHOA HỌC'] || row['KHOA HỌC'] || row['Khoa học'] || row['KH'] || '',
ta: row['KQ VQG TIẾNG ANH'] || row['TIẾNG ANH'] || row['Tiếng Anh'] || row['TA'] || '',
certCode: row['MÃ CERT ĐẦY ĐỦ'] || row['CERT CODE FULL'] || ...,
certCode2: row['MÃ CERT'] || row['CERT CODE'] || ...,
sdt: row['Số điện thoại liên hệ'] || row['SĐT'] || '',
email: row['Email liên hệ'] || row['EMAIL'] || '',
```

**Lý do:** Ưu tiên tên cột mới từ file Excel trước, giữ lại fallback cho tương thích ngược.

#### b) Cập nhật hàm `getMedalClass()` (dòng 279-289):
```javascript
// Thêm hỗ trợ các định dạng mới:
- 'VÀNG' | 'VANG' | 'GOLD' → gold
- 'BẠC' | 'BAC' | 'SILVER' → silver
- 'ĐỒNG' | 'DONG' | 'BRONZE' → bronze
- 'KHUYẾN KHÍCH' | 'KHUYEN KHICH' | 'KK' → encouragement (MỚI)
- 'CHỨNG NHẬN' | 'CHUNG NHAN' | 'CN' → certificate (MỚI)
- Bỏ qua 'nan' và 'NaN'
```

**Lý do:** File Excel dùng định dạng đầy đủ "HUY CHƯƠNG VÀNG", cần nhận diện chính xác.

#### c) Cập nhật hàm hiển thị bảng kết quả (dòng 238-256):
```javascript
// CŨ: 10 cột (bao gồm Area)
// MỚI: 9 cột (bỏ Area)
- Hiển thị certCode2 (rút gọn) thay vì certCode (đầy đủ)
- Xử lý giá trị null/undefined với || ''
- Cập nhật colspan từ 10 → 9
```

#### d) Cập nhật hàm export Excel (dòng 372-385):
```javascript
// Cập nhật tên cột khi export:
'Ngày sinh': student.dob,
'KQ VQG TOÁN': student.toan,
'KQ VQG KHOA HỌC': student.kh,
'KQ VQG TIẾNG ANH': student.ta,
'MÃ CERT ĐẦY ĐỦ': student.certCode,
'MÃ CERT': student.certCode2,
'Số điện thoại liên hệ': student.sdt,
'Email liên hệ': student.email,
```

---

### 2. **File: `index.html`**

#### a) Cập nhật header bảng (dòng 131-144):
```html
<!-- CŨ: 10 cột -->
<th>Area</th>
<th>Toán</th>
<th>Khoa học</th>
<th>Tiếng Anh</th>

<!-- MỚI: 9 cột -->
<th>KQ Toán</th>
<th>KQ Khoa học</th>
<th>KQ Tiếng Anh</th>
```

**Thay đổi:**
- Bỏ cột "Area"
- Đổi tên "Toán" → "KQ Toán" (rõ nghĩa hơn)
- Cập nhật colspan từ 10 → 9

---

### 3. **File: `styles.css`**

#### a) Thêm màu cho badge mới (dòng 409-418):
```css
.score-badge.encouragement {
    background-color: #90EE90;  /* Xanh lá nhạt */
    color: #000;
    border: 2px solid #228B22;
}

.score-badge.certificate {
    background-color: #E0E0E0;  /* Xám nhạt */
    color: #000;
    border: 2px solid #808080;
}
```

#### b) Cập nhật width các cột (dòng 552-595):
```css
/* BỎ cột Area (nth-child 6) */

/* CẬP NHẬT width: */
- School: 15% → 20% (rộng hơn do bỏ Area)
- KQ Toán: child(7) → child(6), 8% → 12%
- KQ Khoa học: child(8) → child(7), 8% → 12%
- KQ Tiếng Anh: child(9) → child(8), 8% → 12%
- Cert Code: child(10) → child(9), 12% → 15%
  + Thêm font: 'Courier New', monospace
```

#### c) Cập nhật min-width bảng:
```css
/* CŨ */
min-width: 1200px;

/* MỚI */
min-width: 1000px;
```

**Lý do:** Giảm số cột từ 10 → 9, không cần bảng quá rộng.

---

## 🎨 MÀU SẮC MEDAL BADGES

| Loại giải | Class | Màu nền | Màu chữ | Border |
|-----------|-------|---------|---------|--------|
| VÀNG | `gold` | #FFD700 | #000 | #666 |
| BẠC | `silver` | #C0C0C0 | #000 | #666 |
| ĐỒNG | `bronze` | #CD7F32 | white | #666 |
| KHUYẾN KHÍCH | `encouragement` | #90EE90 | #000 | #228B22 |
| CHỨNG NHẬN | `certificate` | #E0E0E0 | #000 | #808080 |

---

## 📊 DEMO & TESTING

### File test đã tạo:
1. **`test_index.html`** - Hiển thị demo các medal badge với màu sắc mới
2. **`check_excel_columns.py`** - Script kiểm tra cấu trúc cột Excel

### Cách test:
```bash
# 1. Mở test_index.html để xem demo màu sắc
start test_index.html

# 2. Kiểm tra cột Excel
python check_excel_columns.py

# 3. Chạy web server và test với file thật
python web_server.py
# Mở http://localhost:8000/index.html
# Upload file Awards_Comparison_WITH_CERT.xlsx
```

---

## 🔄 TƯƠNG THÍCH NGƯỢC

Code vẫn giữ **fallback** cho các tên cột cũ:
- `D.O.B` (sau `Ngày sinh`)
- `TOÁN` (sau `KQ VQG TOÁN`)
- `KHOA HỌC` (sau `KQ VQG KHOA HỌC`)
- `TIẾNG ANH` (sau `KQ VQG TIẾNG ANH`)
- `CERT CODE FULL` (sau `MÃ CERT ĐẦY ĐỦ`)
- `CERT CODE` (sau `MÃ CERT`)
- `SĐT` (sau `Số điện thoại liên hệ`)
- `EMAIL` (sau `Email liên hệ`)

→ Vẫn hoạt động với file Excel cũ nếu có!

---

## 📝 GHI CHÚ

1. **Thứ tự ưu tiên cột:** Tên mới → Tên cũ → Empty string
2. **Medal recognition:** Hỗ trợ cả tiếng Việt có dấu và không dấu
3. **Cert Code hiển thị:** Ưu tiên MÃ CERT (rút gọn) thay vì MÃ CERT ĐẦY ĐỦ
4. **Area column:** Đã bỏ khỏi hiển thị (không có trong file Excel mới)

---

## ✅ HOÀN THÀNH

- [x] Cập nhật script.js - Đọc cột Excel mới
- [x] Cập nhật script.js - Hàm getMedalClass()
- [x] Cập nhật script.js - Hiển thị bảng kết quả
- [x] Cập nhật script.js - Export Excel
- [x] Cập nhật index.html - Header bảng
- [x] Cập nhật index.html - Colspan
- [x] Cập nhật styles.css - Màu badge mới
- [x] Cập nhật styles.css - Width các cột
- [x] Tạo test file để demo
- [x] Viết changelog chi tiết

**Status:** ✅ DONE - Sẵn sàng sử dụng!
