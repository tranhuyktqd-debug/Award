# -*- coding: utf-8 -*-
"""
Ứng dụng tạo QR Code cho danh sách điểm danh
Đọc từ file Excel và tạo QR code cho mỗi người
"""
import os
import sys
import pandas as pd
import qrcode
from datetime import datetime

# Set UTF-8 encoding for console
if sys.platform == 'win32':
    try:
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='ignore')
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='ignore')
    except:
        pass

# === CẤU HÌNH ===
INPUT_FILE = r'D:\1 CHUAN BI KY THI\TEST_TRA_CUU_TRAO_GIAI\DS thi sinh SEAMO X.xlsx'
OUTPUT_FILE = r'D:\1 CHUAN BI KY THI\TEST_TRA_CUU_TRAO_GIAI\DS_SEAMO_X_WITH_QR.xlsx'
QR_FOLDER = r'D:\1 CHUAN BI KY THI\TEST_TRA_CUU_TRAO_GIAI\QR_SEAMO'  # Thư mục QR

# Tiêu đề QR
QR_TITLE = "Southeast Asian Mathematical Olympiad (SEAMO X) 2026"

# Mapping cột theo index (file có header ở dòng 1)
# STT, Candidate no, Name, Grade, School, ROLE, TEAM NO
COL_INDEX = {
    'STT': 0,
    'CANDIDATE_NO': 1,
    'NAME': 2,
    'GRADE': 3,
    'SCHOOL': 4,
    'ROLE': 5,
    'TEAM_NO': 6,
}

def create_qr_code(data: str, filename: str, size: int = 300):
    """Tạo QR code và lưu thành file"""
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10,
        border=2,
    )
    qr.add_data(data)
    qr.make(fit=True)
    
    img = qr.make_image(fill_color="black", back_color="white")
    img.save(filename)

def safe_str(value, is_integer=False):
    """Chuyển đổi giá trị thành string an toàn - để trống nếu không có dữ liệu"""
    if pd.isna(value) or value is None or str(value).strip() == '':
        return ''
    # Nếu là số và cần hiển thị dạng số nguyên
    if is_integer:
        try:
            num = float(value)
            if num == int(num):
                return str(int(num))
        except (ValueError, TypeError):
            pass
    return str(value).strip()

def main():
    print("="*70)
    print("🎯 ỨNG DỤNG TẠO QR CODE CHO DANH SÁCH ĐIỂM DANH")
    print("="*70)
    print(f"\n📅 Thời gian: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print(f"📁 File input: {INPUT_FILE}")
    print(f"📁 File output: {OUTPUT_FILE}")
    print(f"📁 Thư mục QR: {QR_FOLDER}")
    
    # Tạo thư mục QR nếu chưa có
    if not os.path.exists(QR_FOLDER):
        os.makedirs(QR_FOLDER)
        print(f"✅ Đã tạo thư mục: {QR_FOLDER}")
    
    # Đọc file Excel
    print("\n📖 Đang đọc file Excel...")
    try:
        df = pd.read_excel(INPUT_FILE)
        print(f"✅ Đã đọc {len(df)} dòng")
    except Exception as e:
        print(f"❌ Lỗi đọc file: {e}")
        return
    
    # Bỏ dòng tiêu đề (dòng 0) và dòng header (dòng 1 - đã dùng làm tên cột)
    # Dữ liệu bắt đầu từ dòng 1 (sau khi đọc, dòng 0 là header thực)
    df_data = df.iloc[1:].copy()
    df_data = df_data.reset_index(drop=True)
    
    print(f"📊 Số người cần tạo QR: {len(df_data)}")
    
    # Tạo cột QR DATA
    qr_data_list = []
    success_count = 0
    skip_count = 0
    
    print("\n" + "="*70)
    print("🔲 BẮT ĐẦU TẠO QR CODE...")
    print("="*70)
    
    for idx, row in df_data.iterrows():
        # Lấy dữ liệu từ các cột theo index
        stt = safe_str(row.iloc[COL_INDEX['STT']])
        candidate_no = safe_str(row.iloc[COL_INDEX['CANDIDATE_NO']])
        name = safe_str(row.iloc[COL_INDEX['NAME']])
        grade = safe_str(row.iloc[COL_INDEX['GRADE']], is_integer=True)
        school = safe_str(row.iloc[COL_INDEX['SCHOOL']])
        role = safe_str(row.iloc[COL_INDEX['ROLE']])
        team_no = safe_str(row.iloc[COL_INDEX['TEAM_NO']])
        
        # Bỏ qua nếu không có tên
        if name == '':
            skip_count += 1
            qr_data_list.append('')
            continue
        
        # Tạo nội dung QR (căn trái, dòng kẻ ngắn hơn)
        qr_content = f"""{QR_TITLE}
━━━━━━━━━━━━━━━
Candidate No: {candidate_no}
Full Name: {name}
━━━━━━━━━━━━━━━
Grade: {grade}
School: {school}
Team No: {team_no}
Role: {role}"""
        
        # Tạo tên file QR theo Candidate No
        # Loại bỏ ký tự không hợp lệ trong tên file
        safe_candidate_no = str(candidate_no).replace('/', '_').replace('\\', '_')
        if safe_candidate_no == '':
            safe_candidate_no = f'IDX_{idx}'
        
        # Tên file: CANDIDATE_NO + NAME
        safe_name = name.replace('/', '_').replace('\\', '_').replace(':', '_').replace('?', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_')
        qr_filename = f"{safe_candidate_no}_{safe_name}.png"
        qr_filepath = os.path.join(QR_FOLDER, qr_filename)
        
        try:
            # Tạo QR code
            create_qr_code(qr_content, qr_filepath)
            qr_data_list.append(qr_content)
            success_count += 1
            
            # Hiển thị tên
            print(f"[✅] {success_count}. {name} - {qr_filename}")
            
        except Exception as e:
            print(f"[❌] Lỗi tạo QR cho {name}: {e}")
            qr_data_list.append('')
    
    # Thêm cột QR DATA vào DataFrame
    df_data['QR_DATA'] = qr_data_list
    
    # Sử dụng trực tiếp df_data (không cần ghép header)
    df_final = df_data
    
    # Lưu file Excel mới
    print("\n" + "="*70)
    print("💾 Đang lưu file Excel...")
    try:
        df_final.to_excel(OUTPUT_FILE, index=False)
        print(f"✅ Đã lưu file: {OUTPUT_FILE}")
    except Exception as e:
        print(f"❌ Lỗi lưu file: {e}")
    
    # Tóm tắt
    print("\n" + "="*70)
    print("📊 TỔNG KẾT:")
    print("="*70)
    print(f"✅ Tạo QR thành công: {success_count}")
    print(f"⏩ Bỏ qua: {skip_count}")
    print(f"📁 File output: {OUTPUT_FILE}")
    print(f"📁 Thư mục QR: {QR_FOLDER}")
    print("="*70)
    print("🎉 HOÀN TẤT!")
    print("="*70)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⚠️ Đã dừng chương trình!")
    except Exception as e:
        print(f"\n\n❌ Lỗi: {e}")
        import traceback
        traceback.print_exc()

