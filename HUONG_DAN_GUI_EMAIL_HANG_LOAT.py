"""
HƯỚNG DẪN GỬI EMAIL HÀNG LOẠT
==============================

Do vấn đề tương thích Python 3.14, khuyến nghị chạy trực tiếp:

CÁCH 1: GỬI TẤT CẢ EMAIL
-------------------------
python send_student_awards.py


CÁCH 2: GỬI CHO 1 HỌC SINH
---------------------------
1. Mở file send_student_awards.py
2. Tìm dòng: for idx, row in df.iterrows():
3. Thêm điều kiện lọc SBD nếu cần
4. Hoặc dùng script test với SBD cụ thể


THEO DÕI TIẾN ĐỘ
----------------
- Xem console/terminal đang chạy
- File log lỗi: send_awards_failed.txt
- Ước tính: ~1-2 giây/email
- Tổng thời gian cho 2399 học sinh: ~40-80 phút


LƯU Ý QUAN TRỌNG
----------------
✅ Đảm bảo file DS_KQ_WITH_QR.xlsx có:
   - Cột SBD (dạng text, giữ số 0)
   - Cột EMAIL  
   - Cột QR DATA

✅ Email config đúng trong email_config.py:
   - EMAIL_SENDER = 'tranhuyktqd@gmail.com'
   - EMAIL_PASSWORD = 'cwrscsxbjspabica'

✅ Logo file: logo ASMO.jpg

✅ Test trước khi gửi hàng loạt:
   python test_send_real_email.py
"""

if __name__ == "__main__":
    print(__doc__)
