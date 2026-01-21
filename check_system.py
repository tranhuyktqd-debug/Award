"""
Script kiểm tra toàn bộ hệ thống trước khi chạy
Chạy: python check_system.py
"""
import os
import sys

def check_file(filepath, description):
    """Kiểm tra file có tồn tại không"""
    if os.path.exists(filepath):
        size = os.path.getsize(filepath)
        print(f"✅ {description}: {filepath} ({size:,} bytes)")
        return True
    else:
        print(f"❌ THIẾU: {description}: {filepath}")
        return False

def check_directory(dirpath, description):
    """Kiểm tra thư mục có tồn tại không"""
    if os.path.exists(dirpath) and os.path.isdir(dirpath):
        count = len([f for f in os.listdir(dirpath) if os.path.isfile(os.path.join(dirpath, f))])
        print(f"✅ {description}: {dirpath} ({count} files)")
        return True
    else:
        print(f"⚠️  KHÔNG CÓ: {description}: {dirpath}")
        return False

def check_python_package(package_name):
    """Kiểm tra Python package đã cài chưa"""
    try:
        __import__(package_name)
        print(f"✅ Python package: {package_name}")
        return True
    except ImportError:
        print(f"❌ THIẾU package: {package_name}")
        return False

def check_excel_structure(filepath):
    """Kiểm tra cấu trúc file Excel"""
    try:
        import pandas as pd
        df = pd.read_excel(filepath, dtype={'SBD': str})
        
        required_columns = ['SBD', 'FULL NAME', 'D.O.B', 'KHỐI', 'TRƯỜNG', 'TOÁN', 'KHOA HỌC', 'TIẾNG ANH']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            print(f"❌ File Excel thiếu cột: {', '.join(missing_columns)}")
            return False
        else:
            print(f"✅ Excel structure OK: {len(df)} học sinh, {len(df.columns)} cột")
            
            # Check email column
            if 'EMAIL' in df.columns:
                email_count = df['EMAIL'].notna().sum()
                print(f"   📧 Email: {email_count}/{len(df)} học sinh có email")
            
            # Check QR DATA column
            if 'QR DATA' in df.columns:
                qr_count = df['QR DATA'].notna().sum()
                print(f"   🔲 QR: {qr_count}/{len(df)} học sinh có QR DATA")
            else:
                print(f"   ⚠️  Chưa có cột QR DATA - cần chạy create_qr_for_all_students.py")
            
            return True
    except Exception as e:
        print(f"❌ Lỗi đọc Excel: {e}")
        return False

def main():
    print("🔍 KIỂM TRA HỆ THỐNG TRA CỨU VÀ GỬI EMAIL")
    print("="*60)
    
    all_ok = True
    
    # 1. Kiểm tra Python packages
    print("\n📦 1. KIỂM TRA PYTHON PACKAGES:")
    packages = ['pandas', 'openpyxl', 'qrcode', 'PIL']
    for pkg in packages:
        if not check_python_package(pkg):
            all_ok = False
    
    # 2. Kiểm tra file web
    print("\n🌐 2. KIỂM TRA FILE WEB:")
    web_files = [
        ('index.html', 'Giao diện web'),
        ('script.js', 'Logic JavaScript'),
        ('styles.css', 'CSS styling'),
    ]
    for filepath, desc in web_files:
        if not check_file(filepath, desc):
            all_ok = False
    
    # 3. Kiểm tra file Python
    print("\n🐍 3. KIỂM TRA FILE PYTHON:")
    python_files = [
        ('web_server.py', 'Web server (PORT 8001)'),
        ('create_qr_for_all_students.py', 'Tạo QR codes'),
        ('send_student_awards.py', 'Gửi email'),
        ('email_config.py', 'Cấu hình email'),
    ]
    for filepath, desc in python_files:
        if not check_file(filepath, desc):
            all_ok = False
    
    # 4. Kiểm tra file data
    print("\n📊 4. KIỂM TRA FILE DỮ LIỆU:")
    
    # Check DATA KQ.xlsx
    if check_file('DATA KQ.xlsx', 'File dữ liệu gốc'):
        check_excel_structure('DATA KQ.xlsx')
    else:
        print("   ⚠️  Cần upload file DATA KQ.xlsx qua web interface")
    
    # Check DS_KQ_WITH_QR.xlsx
    if check_file('DS_KQ_WITH_QR.xlsx', 'File có QR DATA'):
        check_excel_structure('DS_KQ_WITH_QR.xlsx')
    else:
        print("   ℹ️  File DS_KQ_WITH_QR.xlsx chưa tồn tại")
        print("   💡 Chạy: python create_qr_for_all_students.py")
    
    # 5. Kiểm tra file assets
    print("\n🖼️  5. KIỂM TRA FILE ASSETS:")
    check_file('logo ASMO.jpg', 'Logo email')
    check_directory('photos', 'Thư mục ảnh học sinh')
    
    # 6. Kiểm tra port
    print("\n🔌 6. KIỂM TRA CẤU HÌNH PORT:")
    try:
        with open('script.js', 'r', encoding='utf-8') as f:
            content = f.read()
            if 'localhost:8001' in content:
                print("✅ Port trong script.js: 8001")
            else:
                print("⚠️  Port trong script.js không phải 8001")
    except:
        pass
    
    try:
        with open('web_server.py', 'r', encoding='utf-8') as f:
            content = f.read()
            if 'PORT = 8001' in content or "port=8001" in content:
                print("✅ Port trong web_server.py: 8001")
            else:
                print("⚠️  Kiểm tra lại port trong web_server.py")
    except:
        pass
    
    # 7. Kiểm tra email config
    print("\n📧 7. KIỂM TRA CẤU HÌNH EMAIL:")
    try:
        from email_config import EMAIL_SENDER, SMTP_SERVER, SMTP_PORT, EMAIL_PASSWORD
        print(f"✅ Email sender: {EMAIL_SENDER}")
        print(f"✅ SMTP server: {SMTP_SERVER}:{SMTP_PORT}")
        if EMAIL_PASSWORD:
            print(f"✅ Email password: {'*' * len(EMAIL_PASSWORD)}")
        else:
            print("⚠️  Chưa cấu hình EMAIL_PASSWORD")
            print("   💡 Xem: HUONG_DAN_APP_PASSWORD.md")
    except Exception as e:
        print(f"⚠️  Lỗi đọc email_config.py: {e}")
    
    # Summary
    print("\n" + "="*60)
    if all_ok:
        print("✅ TẤT CẢ KIỂM TRA THÀNH CÔNG!")
        print("\n🚀 SẴN SÀNG CHẠY HỆ THỐNG:")
        print("   1. python web_server.py")
        print("   2. Mở trình duyệt: http://localhost:8001/index.html")
    else:
        print("⚠️  CÓ MỘT SỐ VẤN ĐỀ CẦN KHẮC PHỤC")
        print("\n💡 XEM HƯỚNG DẪN:")
        print("   - HUONG_DAN_CHAY_HE_THONG.md")
        print("   - HUONG_DAN_SU_DUNG.md")
    
    print("="*60)

if __name__ == '__main__':
    main()

