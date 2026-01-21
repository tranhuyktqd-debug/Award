# -*- coding: utf-8 -*-
"""
Script tao cot QR DATA cho tat ca hoc sinh trong DATA KQ.xlsx
Output: DS_KQ_WITH_QR.xlsx (giu nguyen tat ca cot + them QR DATA)
"""
import pandas as pd
import sys
import os

# Set UTF-8 encoding for console output
if sys.platform == 'win32':
    try:
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    except:
        pass

print("[*] TAO QR DATA CHO TAT CA HOC SINH...")
print("="*60)

# Kiem tra file ton tai
if not os.path.exists('DATA KQ.xlsx'):
    print("[ERROR] Khong tim thay file 'DATA KQ.xlsx'")
    print("[INFO] Vui long upload file qua web interface truoc")
    sys.exit(1)

# Doc file goc (giu SBD dang string de giu so 0 dau)
try:
    df = pd.read_excel('DATA KQ.xlsx', dtype={'SBD': str})
    print(f"[OK] Doc DATA KQ.xlsx: {len(df)} hoc sinh")
    print(f"[INFO] Columns: {list(df.columns)}")
except Exception as e:
    print(f"[ERROR] Loi doc file Excel: {e}")
    sys.exit(1)

# Ham lay gia tri cot (ho tro ca co dau va khong dau)
def get_column_value(row, col_names):
    """Lay gia tri tu cot, thu nhieu ten khac nhau"""
    for col in col_names:
        if col in row.index and pd.notna(row.get(col)):
            return str(row[col]).strip()
    return ''

# Ham tao QR DATA
def create_qr_data(row):
    """Tao noi dung QR theo format"""
    # Giu nguyen SBD tu Excel (khong strip, giu so 0 dau)
    sbd = str(row.get('SBD', '')) if pd.notna(row.get('SBD')) else ''
    # Format SBD voi leading zeros (9 digits)
    sbd_formatted = sbd.zfill(9) if sbd else ''
    
    # Ho tro ca ten cot co dau va khong dau
    full_name = get_column_value(row, ['FULL NAME', 'Full Name', 'Ho ten', 'Họ tên'])
    dob = get_column_value(row, ['D.O.B', 'D.O.B2', 'DOB', 'Ngay sinh', 'Ngày sinh'])
    grade = get_column_value(row, ['KHOI', 'KHỐI', 'Grade', 'Lop', 'Lớp'])
    school = get_column_value(row, ['TRUONG', 'TRƯỜNG', 'School', 'Trường'])
    toan = get_column_value(row, ['TOAN', 'TOÁN', 'Toan', 'Toán'])
    kh = get_column_value(row, ['KHOA HOC', 'KHOA HỌC', 'KH', 'Khoa hoc', 'Khoa học'])
    ta = get_column_value(row, ['TIENG ANH', 'TIẾNG ANH', 'TA', 'Tieng Anh', 'Tiếng Anh'])
    cert_code = get_column_value(row, ['CERT CODE FULL', 'CERT CODE', 'Cert code', 'Ma chung chi'])
    
    qr_text = f"""STUDENT INFORMATION
Candidate: {sbd_formatted}
Name: {full_name}
Date of Birth: {dob}
Grade {grade} - {school}

RESULTS:
Math: {toan}
Science: {kh}
English: {ta}

Certificate: {cert_code}"""
    
    return qr_text

# Tao cot QR DATA
print("\n[*] Tao QR DATA cho tung hoc sinh...")
try:
    df['QR DATA'] = df.apply(create_qr_data, axis=1)
except Exception as e:
    print(f"[ERROR] Loi tao QR DATA: {e}")
    sys.exit(1)

# Kiem tra
has_qr = df['QR DATA'].notna().sum()
print(f"[OK] Da tao QR DATA: {has_qr}/{len(df)} hoc sinh")

# Luu file moi
output_file = 'DS_KQ_WITH_QR.xlsx'
try:
    df.to_excel(output_file, index=False)
    print(f"\n[OK] Da luu: {output_file}")
except Exception as e:
    print(f"[ERROR] Loi luu file: {e}")
    sys.exit(1)

print("\n[INFO] THONG KE FILE MOI:")
print(f"  - Tong so hoc sinh: {len(df)}")
print(f"  - Tong so cot: {len(df.columns)}")
print(f"  - Co EMAIL: {df['EMAIL'].notna().sum() if 'EMAIL' in df.columns else 0}")
print(f"  - Co QR DATA: {has_qr}")

print("\n[OK] HOAN TAT!")
print(f"[INFO] File moi: {output_file}")
print("   - Giu nguyen TAT CA cot tu DATA KQ.xlsx")
print("   - Them cot QR DATA")
print("   - San sang de gui email!")

print("\n" + "="*60)
print("[INFO] Sample QR DATA (hoc sinh dau tien):")
try:
    print(df['QR DATA'].iloc[0])
except:
    print("(Khong the hien thi)")
print("="*60)
