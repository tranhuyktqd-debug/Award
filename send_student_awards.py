# -*- coding: utf-8 -*-
import os
import sys
import ssl
import smtplib
import pandas as pd
import re
import time
import qrcode
import io
import socket
import requests
from email.message import EmailMessage
from email.mime.image import MIMEImage
from datetime import datetime
from typing import List, Tuple, Optional

# Set UTF-8 encoding for console output
if sys.platform == 'win32':
    try:
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='ignore')
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='ignore')
    except:
        pass

# Import cấu hình từ file riêng
try:
    from email_config import *
    EMAIL = EMAIL_SENDER
    # Import các cấu hình thay đổi IP nếu có
    if 'EMAILS_PER_IP' not in dir():
        EMAILS_PER_IP = 50
    if 'AUTO_PAUSE_FOR_IP_CHANGE' not in dir():
        AUTO_PAUSE_FOR_IP_CHANGE = True
    if 'IP_CHANGE_WAIT_TIME' not in dir():
        IP_CHANGE_WAIT_TIME = 30
except ImportError:
    # Fallback nếu không có file config
    EMAIL = 'seamo@rice-ins.com'
    SMTP_SERVER = 'smtp-relay.gmail.com'
    SMTP_PORT = 587
    EMAIL_PASSWORD = ''
    USE_TLS = True
    EMAILS_PER_IP = 50
    AUTO_PAUSE_FOR_IP_CHANGE = True
    IP_CHANGE_WAIT_TIME = 30

# === CẤU HÌNH ĐƯỜNG DẪAN ===
BASE_DIR = os.path.dirname(__file__)
EXCEL_FILE = os.path.join(BASE_DIR, 'DS_KQ_WITH_QR.xlsx')  # File dữ liệu học sinh với QR
PHOTOS_DIR = os.path.join(BASE_DIR, 'photos')  # Thư mục ảnh học sinh
LOG_FILE = os.path.join(BASE_DIR, 'send_awards_log.txt')
FAILED_LOG_FILE = os.path.join(BASE_DIR, 'send_awards_failed.txt')

# === CẤU HÌNH GỬI EMAIL ===
# Các giá trị này sẽ được override bởi email_config.py nếu có
if 'MAX_RETRIES' not in dir():
    MAX_RETRIES = 3
if 'RETRY_DELAY' not in dir():
    RETRY_DELAY = 2
if 'EMAIL_DELAY' not in dir():
    EMAIL_DELAY = 1

# === CẤU HÌNH QR CODE ===
QR_SIZE = 300  # Kích thước QR code (pixels)
QR_BORDER = 2  # Border size

# === UTILITY FUNCTIONS ===
def is_valid_email(email: str) -> bool:
    """Kiểm tra email có hợp lệ không"""
    if not email or not isinstance(email, str):
        return False
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email.strip()) is not None

def generate_qr_code(data: str) -> bytes:
    """Tạo QR code từ dữ liệu text và trả về bytes"""
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10,
        border=QR_BORDER,
    )
    qr.add_data(data)
    qr.make(fit=True)
    
    img = qr.make_image(fill_color="black", back_color="white")
    
    # Convert to bytes
    img_bytes = io.BytesIO()
    img.save(img_bytes, format='PNG')
    img_bytes.seek(0)
    return img_bytes.read()

def get_medal_emoji(score: str) -> str:
    """Lấy emoji huy chương dựa trên kết quả"""
    if not score:
        return ""
    score_upper = str(score).upper()
    if 'VÀNG' in score_upper or 'GOLD' in score_upper:
        return "🥇"
    elif 'BẠC' in score_upper or 'SILVER' in score_upper:
        return "🥈"
    elif 'ĐỒNG' in score_upper or 'BRONZE' in score_upper:
        return "🥉"
    return ""

def send_email_with_retry(msg: EmailMessage, max_retries: int = MAX_RETRIES) -> Tuple[bool, str]:
    """Gửi email với cơ chế retry"""
    for attempt in range(max_retries):
        try:
            context = ssl.create_default_context()
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
                smtp.ehlo('rice-ins.com')
                if USE_TLS:
                    smtp.starttls(context=context)
                if EMAIL_PASSWORD:
                    smtp.login(EMAIL, EMAIL_PASSWORD)
                smtp.send_message(msg)
            return True, "Thành công"
        except Exception as e:
            error_msg = str(e)
            if attempt < max_retries - 1:
                delay = RETRY_DELAY * (2 ** attempt)
                print(f"[🔄] Lỗi lần {attempt + 1}/{max_retries}: {error_msg}. Thử lại sau {delay}s...")
                time.sleep(delay)
            else:
                return False, f"Thất bại sau {max_retries} lần thử: {error_msg}"
    return False, "Không xác định"

def get_current_ip() -> str:
    """Lấy địa chỉ IP công khai hiện tại"""
    try:
        # Thử nhiều service để đảm bảo
        services = [
            'https://api.ipify.org?format=text',
            'https://icanhazip.com',
            'https://ifconfig.me/ip'
        ]
        for service in services:
            try:
                response = requests.get(service, timeout=5)
                if response.status_code == 200:
                    return response.text.strip()
            except:
                continue
        return "Unknown"
    except Exception as e:
        return f"Error: {e}"

def check_keyboard_input():
    """Kiểm tra xem có phím nào được nhấn không (cross-platform)"""
    try:
        if sys.platform == 'win32':
            import msvcrt
            if msvcrt.kbhit():
                key = msvcrt.getch()
                return key == b'\r'  # Enter key
        else:
            # Unix/Linux/Mac
            import select
            if select.select([sys.stdin], [], [], 0)[0]:
                sys.stdin.readline()
                return True
        return False
    except:
        return False

def wait_for_ip_change(old_ip: str, timeout: int = 300) -> str:
    """Chờ đợi IP thay đổi (tối đa timeout giây)"""
    print(f"\n{'='*60}")
    print(f"🔄 IP HIỆN TẠI: {old_ip}")
    print(f"{'='*60}")
    print(f"\n⏸️  TẠM DỪNG ĐỂ THAY ĐỔI KẾT NỐI INTERNET")
    print(f"\n📋 HƯỚNG DẪN:")
    print(f"   1. Ngắt kết nối Internet hiện tại (WiFi/4G/Ethernet)")
    print(f"   2. Chuyển sang kết nối khác (VPN/4G/WiFi khác)")
    print(f"   3. Đợi hệ thống tự động phát hiện IP mới")
    print(f"   4. Hoặc nhấn ENTER để bỏ qua và tiếp tục\n")
    
    start_time = time.time()
    check_interval = 3  # Kiểm tra mỗi 3 giây
    
    while time.time() - start_time < timeout:
        # Cho phép người dùng bỏ qua
        if check_keyboard_input():
            print("\n⏭️  Bỏ qua kiểm tra IP, tiếp tục gửi email...")
            return get_current_ip()
        
        # Kiểm tra IP mới
        new_ip = get_current_ip()
        elapsed = int(time.time() - start_time)
        
        if new_ip != old_ip and new_ip != "Unknown":
            print(f"\n✅ PHÁT HIỆN IP MỚI: {new_ip}")
            print(f"⏱️  Chờ {IP_CHANGE_WAIT_TIME}s để ổn định kết nối...")
            time.sleep(IP_CHANGE_WAIT_TIME)
            print(f"✅ Sẵn sàng tiếp tục gửi email!\n")
            return new_ip
        
        # Hiển thị tiến trình chờ
        print(f"\r⏳ Đang chờ IP mới... ({elapsed}s/{timeout}s) - IP: {new_ip} | Nhấn ENTER để bỏ qua", end='', flush=True)
        time.sleep(check_interval)
    
    print(f"\n⚠️  Timeout! Tiếp tục với IP hiện tại: {get_current_ip()}")
    return get_current_ip()

def log_failed_email(sbd: str, name: str, email: str, error_msg: str):
    """Ghi log email gửi thất bại"""
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(FAILED_LOG_FILE, "a", encoding="utf-8") as f:
        log_entry = f"[{current_time}] SBD:{sbd} | Name:{name} | Email:{email} | Error:{error_msg}\n"
        f.write(log_entry)

def build_html_email(student: dict) -> str:
    """Tạo nội dung HTML cho email"""
    math_emoji = get_medal_emoji(student.get('toan', ''))
    science_emoji = get_medal_emoji(student.get('kh', ''))
    english_emoji = get_medal_emoji(student.get('ta', ''))
    
    return f"""<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
            line-height: 1.6;
            color: #333;
            background-color: #f5f5f5;
            padding: 30px 20px;
            margin: 0;
        }}
        .email-wrapper {{
            max-width: 700px;
            margin: 0 auto;
            background: #fffbf0;
            border: 3px solid #7b68ee;
            border-radius: 20px;
            box-shadow: 0 4px 15px rgba(123, 104, 238, 0.2);
            overflow: hidden;
        }}
        .header {{
            width: 100%;
            max-width: 100%;
        }}
        .header-top {{
            background: #fffbeb;
            padding: 20px;
            text-align: center;
        }}
        .header-top table {{
            width: 100%;
            max-width: 680px;
            margin: 0 auto;
        }}
        .logo {{
            max-width: 140px;
            height: auto;
            display: block;
        }}
        .header-title {{
            color: #1565c0;
            font-size: 42px;
            font-weight: 700;
            margin: 0;
            letter-spacing: 2px;
            text-align: center;
            padding: 10px;
        }}
        .header-bg {{
            width: 100%;
            max-width: 100%;
            display: block;
        }}
        .header-bg img {{
            width: 100%;
            max-width: 680px;
            height: auto;
            display: block;
            margin: 0 auto;
        }}
        .header-subtitle {{
            background: #fffbeb;
            padding: 20px;
            text-align: center;
        }}
        .subtitle-line {{
            color: #ff7f27;
            font-size: 28px;
            font-weight: 700;
            letter-spacing: 2px;
            text-transform: uppercase;
            line-height: 1.4;
            display: block;
        }}
        @keyframes shimmer {{
            0%, 100% {{ opacity: 1; }}
            50% {{ opacity: 0.85; }}
        }}
        .content {{
            padding: 40px 35px;
            background: #fffbf0;
        }}
        .greeting {{
            font-size: 16px;
            color: #444;
            margin-bottom: 20px;
            line-height: 1.8;
        }}
        .info-card {{
            background: #ffffff;
            border: 2px solid #d4c5f9;
            padding: 30px 25px;
            margin: 35px 0;
            border-radius: 15px;
            text-align: center;
            box-shadow: 0 2px 8px rgba(123, 104, 238, 0.1);
        }}
        .info-title {{
            color: #5a4a8f;
            font-size: 18px;
            font-weight: 600;
            margin-bottom: 25px;
            text-align: center;
        }}
        .info-item {{
            padding: 8px 0;
            border-bottom: 1px solid #e0e0e0;
            text-align: center;
        }}
        .info-item:last-child {{
            border-bottom: none;
        }}
        .info-label {{
            font-size: 13px;
            color: #666;
            margin-bottom: 3px;
        }}
        .info-value {{
            font-size: 15px;
            font-weight: 600;
            color: #5a4a8f;
        }}
        .results-title {{
            text-align: center;
            color: #5a4a8f;
            font-size: 20px;
            font-weight: 600;
            margin: 40px 0 30px 0;
            text-transform: uppercase;
            letter-spacing: 1px;
        }}
        .results-container {{
            display: flex;
            justify-content: center;
            align-items: center;
            margin: 40px auto;
            width: 100%;
        }}
        .results-grid {{
            display: flex;
            gap: 20px;
            justify-content: center;
            align-items: stretch;
            max-width: 600px;
            margin: 0 auto;
        }}
        .result-card {{
            flex: 0 0 auto;
            width: 180px;
            background: #ffffff;
            border: 2px solid #d4c5f9;
            border-radius: 15px;
            padding: 25px 15px;
            text-align: center;
            transition: all 0.3s;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            box-shadow: 0 2px 8px rgba(123, 104, 238, 0.1);
        }}
        .result-card:hover {{
            box-shadow: 0 6px 16px rgba(123, 104, 238, 0.25);
            transform: translateY(-3px);
            border-color: #7b68ee;
        }}
        .result-header {{
            width: 100%;
            margin-bottom: 15px;
        }}
        .result-icon {{
            font-size: 40px;
            margin-bottom: 10px;
            display: block;
        }}
        .result-subject {{
            font-size: 15px;
            color: #333;
            font-weight: 600;
            line-height: 1.3;
            margin: 0;
            display: block;
        }}
        .result-subject-en {{
            font-size: 11px;
            color: #999;
            font-style: italic;
            line-height: 1.3;
            margin: 3px 0 0 0;
            display: block;
        }}
        .result-content {{
            width: 100%;
            margin-top: 10px;
        }}
        .result-medal {{
            font-size: 60px;
            margin: 10px 0;
            line-height: 1;
            display: block;
        }}
        .result-text {{
            font-size: 16px;
            font-weight: 700;
            color: #5a4a8f;
            margin-top: 8px;
            line-height: 1.3;
            display: block;
        }}
        .cert-box {{
            background: linear-gradient(135deg, #f5f0ff 0%, #e8dcff 100%);
            border: 2px solid #9b8fd9;
            border-radius: 15px;
            padding: 25px;
            margin: 40px 0;
            text-align: center;
            box-shadow: 0 2px 8px rgba(123, 104, 238, 0.15);
        }}
        .cert-label {{
            font-size: 14px;
            color: #5a4a8f;
            font-weight: 600;
            margin-bottom: 10px;
        }}
        .cert-code {{
            font-size: 22px;
            font-weight: 700;
            color: #5a4a8f;
            letter-spacing: 1px;
        }}
        .qr-section {{
            background: #ffffff;
            border: 2px solid #d4c5f9;
            border-radius: 15px;
            padding: 30px;
            margin: 40px 0;
            text-align: center;
            box-shadow: 0 2px 8px rgba(123, 104, 238, 0.1);
        }}
        .qr-title {{
            color: #5a4a8f;
            font-size: 18px;
            font-weight: 600;
            margin-bottom: 15px;
        }}
        .qr-subtitle {{
            color: #666;
            font-size: 14px;
            margin-bottom: 20px;
        }}
        .qr-image {{
            max-width: 220px;
            height: auto;
            margin: 20px auto;
            display: block;
            padding: 15px;
            background: #fffbf0;
            border-radius: 12px;
            border: 2px solid #d4c5f9;
        }}
        .message-box {{
            background: #f5f0ff;
            border-left: 4px solid #7b68ee;
            padding: 20px;
            margin: 35px 0;
            border-radius: 12px;
            color: #5a4a8f;
            line-height: 1.8;
        }}
        .signature {{
            text-align: center;
            margin-top: 40px;
            padding-top: 30px;
            border-top: 2px solid #d4c5f9;
        }}
        .signature-line {{
            font-size: 14px;
            color: #666;
            font-weight: 400;
            line-height: 1.6;
            margin-bottom: 8px;
        }}
        .signature-org {{
            font-size: 16px;
            font-weight: 700;
            color: #5a4a8f;
            line-height: 1.6;
            margin-top: 5px;
        }}
        .footer {{
            background: linear-gradient(135deg, #5a4a8f 0%, #7b68ee 100%);
            color: white;
            padding: 30px;
            text-align: center;
            border-top: 4px solid #4a3a7f;
        }}
        .footer-org {{
            font-size: 16px;
            font-weight: 600;
            margin-bottom: 5px;
            color: white;
        }}
        .footer-subtitle {{
            font-size: 14px;
            font-weight: 500;
            margin-bottom: 15px;
            color: white;
            font-style: italic;
        }}
        .footer-contact {{
            font-size: 13px;
            color: white;
            margin: 5px 0;
        }}
        .footer-contact a {{
            color: white;
            text-decoration: none;
        }}
        .footer-contact a:hover {{
            text-decoration: underline;
        }}
        .footer-note {{
            font-size: 12px;
            color: white;
            margin-top: 15px;
            font-style: italic;
            opacity: 0.9;
        }}
    </style>
</head>
<body>
    <div class="email-wrapper">
        <div class="header">
            <div class="header-top">
                <table cellpadding="0" cellspacing="0" border="0" width="100%" style="max-width: 680px; margin: 0 auto;">
                    <tr>
                        <td width="160" valign="middle" align="left" style="padding: 10px;">
                            <img src="cid:logo_asmo" class="logo" alt="ASMO Logo" style="max-width: 140px; height: auto; display: block;" />
                        </td>
                        <td valign="middle" align="center">
                            <div class="header-title">ASMO VIETNAM</div>
                        </td>
                        <td width="160"></td>
                    </tr>
                </table>
            </div>
            <div class="header-bg">
                <img src="cid:background_image" alt="Background" style="width: 100%; max-width: 680px; height: auto; display: block; margin: 0 auto;" />
            </div>
            <div class="header-subtitle">
                <span class="subtitle-line">THÔNG BÁO GIẢI THƯỞNG</span>
                <span class="subtitle-line">AWARD NOTIFICATION</span>
            </div>
        </div>
    
        <div class="content">
            <div class="greeting">
                <strong>Kính gửi Phụ huynh học sinh: {student['fullName']},</strong><br/><br/>
                Ban Tổ chức Kỳ thi Olympic Quốc tế Khoa học, Toán & Tiếng Anh ASMO Việt Nam xin chúc mừng em đã hoàn thành xuất sắc Vòng Quốc gia và đạt kết quả ấn tượng!
            </div>
            
            <div class="info-card">
                <div class="info-title">📋 THÔNG TIN HỌC SINH • STUDENT INFORMATION</div>
                <div class="info-item">
                    <div class="info-label">Số báo danh • Candidate</div>
                    <div class="info-value">{student.get('candidate', 'N/A')}</div>
                </div>
                <div class="info-item">
                    <div class="info-label">Họ và tên • Full Name</div>
                    <div class="info-value">{student.get('fullName', 'N/A')}</div>
                </div>
                <div class="info-item">
                    <div class="info-label">Ngày sinh • Date of Birth</div>
                    <div class="info-value">{student.get('dob', 'N/A')}</div>
                </div>
                <div class="info-item">
                    <div class="info-label">Lớp • Grade</div>
                    <div class="info-value">{student.get('grade', 'N/A')}</div>
                </div>
                <div class="info-item">
                    <div class="info-label">Trường • School</div>
                    <div class="info-value">{student.get('school', 'N/A')}</div>
                </div>
            </div>
            
            <div class="results-title">🏆 KẾT QUẢ • RESULTS</div>
            
            <div class="results-container">
                <div class="results-grid">
                    <div class="result-card">
                        <div class="result-header">
                            <div class="result-icon">📐</div>
                            <div class="result-subject">Toán học</div>
                            <div class="result-subject-en">Mathematics</div>
                        </div>
                        <div class="result-content">
                            <div class="result-medal">{math_emoji}</div>
                            <div class="result-text">{student.get('toan', 'N/A')}</div>
                        </div>
                    </div>
                    <div class="result-card">
                        <div class="result-header">
                            <div class="result-icon">🔬</div>
                            <div class="result-subject">Khoa học</div>
                            <div class="result-subject-en">Science</div>
                        </div>
                        <div class="result-content">
                            <div class="result-medal">{science_emoji}</div>
                            <div class="result-text">{student.get('kh', 'N/A')}</div>
                        </div>
                    </div>
                    <div class="result-card">
                        <div class="result-header">
                            <div class="result-icon">🗣️</div>
                            <div class="result-subject">Tiếng Anh</div>
                            <div class="result-subject-en">English</div>
                        </div>
                        <div class="result-content">
                            <div class="result-medal">{english_emoji}</div>
                            <div class="result-text">{student.get('ta', 'N/A')}</div>
                        </div>
                    </div>
                </div>
            </div>
            
            <div class="cert-box">
                <div class="cert-label">📜 MÃ CHỨNG CHỈ • CERTIFICATE CODE</div>
                <div class="cert-code">{student.get('certCode', 'N/A')}</div>
            </div>
            
            <div class="qr-section">
                <div class="qr-title">📱 MÃ QR • QR CODE</div>
                <div class="qr-subtitle">Quét mã QR để xác minh thông tin • Scan QR code to verify information</div>
                <img src="cid:qr_code" alt="QR Code" class="qr-image" />
            </div>
            
            <div class="message-box">
                💫 Một lần nữa, BTC xin chúc mừng em đã đạt thành tích xuất sắc! Chúc em tiếp tục phát huy tài năng và đạt được nhiều thành công rực rỡ hơn trong tương lai!
            </div>
            
            <div class="signature">
                <div class="signature-line">Trân trọng • Best regards,</div>
                <div class="signature-org">Ban Tổ chức ASMO Vietnam</div>
            </div>
        </div>
        
        <div class="footer">
            <div class="footer-org">ASMO VIETNAM</div>
            <div class="footer-subtitle">Asian Science and Mathematics Olympiad</div>
            <div class="footer-contact">Email: <a href="mailto:asmo@rice-ins.com">asmo@rice-ins.com</a> | Website: <a href="https://www.asmo.edu.vn/">https://www.asmo.edu.vn/</a></div>
            <div class="footer-note">Email này được gửi tự động, vui lòng không trả lời • This is an automated email, please do not reply</div>
        </div>
    </div>
</body>
</html>"""

def load_sent_logs() -> set:
    """Load danh sách email đã gửi từ log file"""
    sent_logs = set()
    if os.path.exists(LOG_FILE):
        try:
            with open(LOG_FILE, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line and '] ' in line:
                        # Extract SBD|email from log
                        log_key = line.split('] ', 1)[1].split(' | ')[0]
                        sent_logs.add(log_key)
        except Exception as e:
            print(f"[⚠️] Lỗi đọc log file: {e}")
    return sent_logs

def main():
    print("🚀 BẮT ĐẦU GỬI EMAIL THÔNG BÁO GIẢI THƯỞNG...")
    print(f"📅 Thời gian: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print("="*60)
    
    # Lấy IP ban đầu
    current_ip = get_current_ip()
    print(f"🌐 IP hiện tại: {current_ip}")
    print(f"🔄 Tự động đề nghị đổi IP sau mỗi {EMAILS_PER_IP} email: {'BẬT' if AUTO_PAUSE_FOR_IP_CHANGE else 'TẮT'}")
    
    # Load logs
    sent_logs = load_sent_logs()
    print(f"📋 Đã load {len(sent_logs)} email đã gửi từ log")
    
    # Đọc dữ liệu học sinh (giữ SBD dạng string để giữ số 0 đầu)
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name='Sheet1', dtype={'SBD': str})
        print(f"📋 Đã load {len(df)} học sinh từ Excel")
    except Exception as e:
        print(f"[❌] Lỗi đọc file Excel: {e}")
        return
    
    # Thống kê
    sent_count = 0
    failed_count = 0
    skipped_count = 0
    emails_sent_with_current_ip = 0
    start_time = time.time()
    
    print("\n📧 BẮT ĐẦU GỬI EMAIL...")
    print("-"*60)
    
    for idx, row in df.iterrows():
        # Parse student data
        student = {
            'candidate': str(row.get('SBD', '')) if pd.notna(row.get('SBD')) else '',
            'fullName': str(row.get('FULL NAME', '')).strip() if pd.notna(row.get('FULL NAME')) else '',
            'dob': str(row.get('D.O.B', '')).strip() if pd.notna(row.get('D.O.B')) else '',
            'grade': str(row.get('KHỐI', '')).strip() if pd.notna(row.get('KHỐI')) else '',
            'school': str(row.get('TRƯỜNG', '')).strip() if pd.notna(row.get('TRƯỜNG')) else '',
            'toan': str(row.get('TOÁN', '')).strip() if pd.notna(row.get('TOÁN')) else '',
            'kh': str(row.get('KHOA HỌC', '')).strip() if pd.notna(row.get('KHOA HỌC')) else '',
            'ta': str(row.get('TIẾNG ANH', '')).strip() if pd.notna(row.get('TIẾNG ANH')) else '',
            'certCode': str(row.get('CERT CODE FULL', '')).strip() if pd.notna(row.get('CERT CODE FULL')) else '',
            'qrData': str(row.get('QR DATA', '')).strip() if pd.notna(row.get('QR DATA')) else '',
        }
        
        # Lấy email từ DS_KQ_WITH_QR.xlsx
        raw_email = ''
        if 'EMAIL' in df.columns and pd.notna(row.get('EMAIL')):
            raw_email = str(row.get('EMAIL')).strip()
        
        # Validate
        if not raw_email:
            print(f"[⚠️] Bỏ qua {student['candidate']} - {student['fullName']}: không có email")
            skipped_count += 1
            continue
        
        # Parse recipients
        recipients = [e.strip() for e in raw_email.replace(";", ",").split(",") if e.strip()]
        valid_recipients = [email for email in recipients if is_valid_email(email)]
        
        if not valid_recipients:
            print(f"[❌] Bỏ qua {student['candidate']} - {student['fullName']}: email không hợp lệ")
            skipped_count += 1
            continue
        
        # Check log
        log_key = f"{student['candidate']}|{valid_recipients[0].lower()}"
        if log_key in sent_logs:
            print(f"[⏩] Đã gửi: {student['candidate']} - {student['fullName']}")
            skipped_count += 1
            continue
        
        # Generate QR code from Excel data
        try:
            qr_data = student.get('qrData', '')
            if not qr_data:
                print(f"[⚠️] Bỏ qua {student['candidate']} - {student['fullName']}: không có QR data")
                skipped_count += 1
                continue
            qr_image_bytes = generate_qr_code(qr_data)
        except Exception as e:
            print(f"[❌] Lỗi tạo QR cho {student['candidate']}: {e}")
            log_failed_email(student['candidate'], student['fullName'], ', '.join(valid_recipients), f"QR Error: {e}")
            failed_count += 1
            continue
        
        # Build email
        subject = f"ASMO - THÔNG BÁO GIẢI THƯỞNG / AWARD NOTIFICATION — {student['fullName']}"
        html_body = build_html_email(student)
        
        # Create message
        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = f"SEAMO VIETNAM <{EMAIL}>"
        msg['To'] = ", ".join(valid_recipients)
        
        msg.set_content(f"Kính gửi Phụ huynh học sinh {student['fullName']},\n\nVui lòng xem email HTML để xem đầy đủ thông tin giải thưởng.\n\nTrân trọng,\nBan Tổ chức SEAMO Vietnam")
        msg.add_alternative(html_body, subtype='html')
        
        # Embed background image inline
        background_path = os.path.join(BASE_DIR, 'background.png')
        if os.path.exists(background_path):
            try:
                with open(background_path, 'rb') as f:
                    bg_data = f.read()
                    bg_part = MIMEImage(bg_data, _subtype='png')
                    bg_part.add_header('Content-ID', '<background_image>')
                    bg_part.add_header('Content-Disposition', 'inline')
                    msg.attach(bg_part)
            except Exception as e:
                print(f"[WARN] Khong the embed background: {e}")
        
        # Embed logo ASMO inline (không đính kèm file)
        logo_path = os.path.join(BASE_DIR, 'logo ASMO.jpg')
        if os.path.exists(logo_path):
            try:
                with open(logo_path, 'rb') as f:
                    logo_data = f.read()
                    # Embed inline với cid, KHÔNG có filename
                    logo_part = MIMEImage(logo_data, _subtype='jpeg')
                    logo_part.add_header('Content-ID', '<logo_asmo>')
                    logo_part.add_header('Content-Disposition', 'inline')
                    # KHÔNG set filename để tránh hiển thị như attachment
                    msg.attach(logo_part)
            except Exception as e:
                print(f"[WARN] Khong the embed logo: {e}")
        
        # Embed QR code inline (không đính kèm file)
        qr_part = MIMEImage(qr_image_bytes, _subtype='png')
        qr_part.add_header('Content-ID', '<qr_code>')
        qr_part.add_header('Content-Disposition', 'inline')
        # KHÔNG set filename để tránh hiển thị như attachment
        msg.attach(qr_part)
        
        # Attach photo if exists (optional)
        photo_path = os.path.join(PHOTOS_DIR, f'{student["candidate"]}.jpg')
        if os.path.exists(photo_path):
            try:
                with open(photo_path, 'rb') as photo_file:
                    photo_data = photo_file.read()
                    msg.add_attachment(
                        photo_data,
                        maintype='image',
                        subtype='jpeg',
                        filename=f'Photo_{student["candidate"]}.jpg'
                    )
            except Exception as e:
                print(f"[⚠️] Không thể đính kèm ảnh cho {student['candidate']}: {e}")
        
        # Send email
        success, message = send_email_with_retry(msg)
        
        if success:
            print(f"[✅] Gửi thành công: {student['candidate']} - {student['fullName']} ({valid_recipients[0]})")
            
            # Log
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with open(LOG_FILE, "a", encoding="utf-8") as f:
                f.write(f"[{current_time}] {log_key} | {student['fullName']} | {', '.join(valid_recipients)}\n")
            
            sent_count += 1
            emails_sent_with_current_ip += 1
            
            # Kiểm tra xem có cần đổi IP không
            if AUTO_PAUSE_FOR_IP_CHANGE and emails_sent_with_current_ip >= EMAILS_PER_IP:
                print(f"\n{'='*60}")
                print(f"📊 ĐÃ GỬI {emails_sent_with_current_ip} EMAIL VỚI IP HIỆN TẠI")
                print(f"{'='*60}")
                
                # Đề nghị đổi IP
                old_ip = current_ip
                current_ip = wait_for_ip_change(old_ip)
                
                # Reset counter
                emails_sent_with_current_ip = 0
                
                print(f"\n{'='*60}")
                print(f"🔄 ĐÃ ĐỔI IP: {old_ip} → {current_ip}")
                print(f"📧 Tiếp tục gửi email...")
                print(f"{'='*60}\n")
        else:
            print(f"[❌] Lỗi gửi {student['candidate']} - {student['fullName']}: {message}")
            log_failed_email(student['candidate'], student['fullName'], ', '.join(valid_recipients), message)
            failed_count += 1
        
        # Delay
        time.sleep(EMAIL_DELAY)
    
    # Summary
    print("\n" + "="*60)
    print("📊 TỔNG KẾT:")
    print(f"✅ Gửi thành công: {sent_count}")
    print(f"❌ Gửi thất bại: {failed_count}")
    print(f"⏩ Bỏ qua: {skipped_count}")
    print(f"📧 Tổng cộng: {sent_count + failed_count + skipped_count}")
    print(f"⏱️ Thời gian: {time.time() - start_time:.2f} giây")
    print("="*60)

def send_single_student_email(sbd: str) -> dict:
    """Gửi email cho 1 học sinh theo SBD"""
    try:
        # Read Excel
        df = pd.read_excel(EXCEL_FILE, dtype={'SBD': str}, sheet_name='Sheet1')
        
        # Find student
        student_row = df[df['SBD'] == sbd]
        if len(student_row) == 0:
            return {'status': 'error', 'message': f'Không tìm thấy học sinh SBD: {sbd}'}
        
        row = student_row.iloc[0]
        student = {
            'candidate': str(row.get('SBD', '')),
            'fullName': str(row.get('FULL NAME', '')).strip(),
            'dob': str(row.get('D.O.B', '')).strip(),
            'grade': str(row.get('KHỐI', '')).strip(),
            'school': str(row.get('TRƯỜNG', '')).strip(),
            'toan': str(row.get('TOÁN', '')).strip(),
            'kh': str(row.get('KHOA HỌC', '')).strip(),
            'ta': str(row.get('TIẾNG ANH', '')).strip(),
            'certCode': str(row.get('CERT CODE FULL', '')).strip(),
            'qrData': str(row.get('QR DATA', '')).strip(),
        }
        
        # Get email
        email = str(row.get('EMAIL', '')).strip() if pd.notna(row.get('EMAIL')) else ''
        if not email:
            return {'status': 'error', 'message': f'Học sinh {sbd} không có email'}
        
        if not is_valid_email(email):
            return {'status': 'error', 'message': f'Email không hợp lệ: {email}'}
        
        # Generate QR and email
        qr_image_bytes = generate_qr_code(student['qrData'])
        html_body = build_html_email(student)
        
        # Create message
        subject = f"ASMO - THÔNG BÁO GIẢI THƯỞNG / AWARD NOTIFICATION — {student['fullName']}"
        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = f"ASMO VIETNAM <{EMAIL}>"
        msg['To'] = email
        
        msg.set_content(f"Kính gửi Phụ huynh học sinh {student['fullName']}...")
        msg.add_alternative(html_body, subtype='html')
        
        # Embed background image inline
        background_path = os.path.join(BASE_DIR, 'background.png')
        if os.path.exists(background_path):
            with open(background_path, 'rb') as f:
                bg_data = f.read()
                bg_img = MIMEImage(bg_data, _subtype='png')
                bg_img.add_header('Content-ID', '<background_image>')
                bg_img.add_header('Content-Disposition', 'inline')
                msg.attach(bg_img)
        
        # Embed logo inline (không đính kèm file)
        logo_path = os.path.join(BASE_DIR, 'logo ASMO.jpg')
        if os.path.exists(logo_path):
            with open(logo_path, 'rb') as f:
                logo_data = f.read()
                logo_img = MIMEImage(logo_data, _subtype='jpeg')
                logo_img.add_header('Content-ID', '<logo_asmo>')
                logo_img.add_header('Content-Disposition', 'inline')
                # KHÔNG set filename để tránh hiển thị như attachment
                msg.attach(logo_img)
        
        # Embed QR code inline (không đính kèm file)
        qr_part = MIMEImage(qr_image_bytes, _subtype='png')
        qr_part.add_header('Content-ID', '<qr_code>')
        qr_part.add_header('Content-Disposition', 'inline')
        # KHÔNG set filename để tránh hiển thị như attachment
        msg.attach(qr_part)
        
        # Send
        success, error_msg = send_email_with_retry(msg)
        if success:
            return {'status': 'success', 'message': f'Da gui email den {email}'}
        else:
            return {'status': 'error', 'message': f'Loi gui: {error_msg}'}
            
    except Exception as e:
        return {'status': 'error', 'message': f'Loi: {str(e)}'}

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⚠️ Đã dừng chương trình!")
    except Exception as e:
        print(f"\n\n❌ Lỗi: {e}")
