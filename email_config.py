# === CẤU HÌNH EMAIL GỬI ===

# Thông tin email gửi
EMAIL_SENDER = 'tranhuyktqd@gmail.com'
EMAIL_PASSWORD = 'cwrscsxbjspabica'  # App Password đã bỏ dấu cách

# Cấu hình SMTP Server cho Gmail
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
USE_TLS = True

# Hoặc nếu dùng Gmail cá nhân:
# SMTP_SERVER = 'smtp.gmail.com'
# SMTP_PORT = 587
# EMAIL_PASSWORD = 'your-app-password-here'  # App Password, không phải mật khẩu thường

# Hoặc dùng Outlook:
# SMTP_SERVER = 'smtp-mail.outlook.com'
# SMTP_PORT = 587

# === CẤU HÌNH GỬI EMAIL ===
MAX_RETRIES = 3          # Số lần thử lại khi gửi thất bại
RETRY_DELAY = 2          # Thời gian chờ giữa các lần thử (giây)
EMAIL_DELAY = 1          # Thời gian chờ giữa các email (giây)

# === CẤU HÌNH THAY ĐỔI IP ===
EMAILS_PER_IP = 50       # Số email gửi trước khi đề nghị đổi IP
AUTO_PAUSE_FOR_IP_CHANGE = True  # Tự động tạm dừng để đổi IP (True/False)
IP_CHANGE_WAIT_TIME = 30 # Thời gian chờ sau khi đổi IP (giây)

# === CẤU HÌNH QR CODE ===
QR_SIZE = 300            # Kích thước QR code (pixels)
QR_BORDER = 2            # Border size

# === HƯỚNG DẪN ===
"""
1. Nếu dùng Gmail cá nhân:
   - Bật 2-Step Verification: https://myaccount.google.com/security
   - Tạo App Password: https://myaccount.google.com/apppasswords
   - Dán App Password vào EMAIL_PASSWORD

2. Nếu dùng SMTP Relay (Google Workspace):
   - Không cần password
   - Để EMAIL_PASSWORD = 'cwrs csxb jspa bica'

3. Test SMTP:
   python test_smtp_connection.py
"""
