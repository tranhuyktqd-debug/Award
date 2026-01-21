# -*- coding: utf-8 -*-
"""
Script so sánh 2 file Excel theo cột SBD và tạo file mới
"""
import pandas as pd
import sys
import os

# Set UTF-8 encoding for console
if sys.platform == 'win32':
    try:
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    except:
        pass

print("="*80)
print("📊 SO SÁNH VÀ PHÂN TÁCH DỮ LIỆU")
print("="*80)

# Đường dẫn file
file_trao_giai = r'D:\ASMO\TEST_TRA_CUU_TRAO_GIAI\Awards_TRAO GIAI.xlsx'
file_full = r'D:\ASMO\TEST_TRA_CUU_TRAO_GIAI\Awards_Template_Full.xlsx'
file_output = r'D:\ASMO\TEST_TRA_CUU_TRAO_GIAI\Awards_Comparison_Result.xlsx'

try:
    # 1. Đọc file TRAO GIẢI
    print(f"\n1️⃣ Đọc file TRAO GIẢI...")
    if not os.path.exists(file_trao_giai):
        print(f"   ❌ File không tồn tại: {file_trao_giai}")
        sys.exit(1)
    
    df_trao_giai = pd.read_excel(file_trao_giai)
    print(f"   ✅ Đọc thành công: {len(df_trao_giai)} dòng")
    
    # Kiểm tra cột SBD
    if 'SBD' not in df_trao_giai.columns:
        print(f"   ❌ Không tìm thấy cột 'SBD'")
        print(f"   📋 Các cột có: {list(df_trao_giai.columns)}")
        sys.exit(1)
    
    # Lấy danh sách SBD trong file TRAO GIẢI
    sbd_trao_giai = set(df_trao_giai['SBD'].dropna().astype(str))
    print(f"   📊 Số SBD duy nhất: {len(sbd_trao_giai)}")
    
    # 2. Đọc file FULL
    print(f"\n2️⃣ Đọc file Full...")
    if not os.path.exists(file_full):
        print(f"   ❌ File không tồn tại: {file_full}")
        sys.exit(1)
    
    df_full = pd.read_excel(file_full)
    print(f"   ✅ Đọc thành công: {len(df_full)} dòng")
    
    # Kiểm tra cột SBD
    if 'SBD' not in df_full.columns:
        print(f"   ❌ Không tìm thấy cột 'SBD'")
        print(f"   📋 Các cột có: {list(df_full.columns)}")
        sys.exit(1)
    
    # 3. So sánh và phân tách
    print(f"\n3️⃣ So sánh dữ liệu...")
    
    # Chuyển SBD sang string để so sánh
    df_full['SBD_str'] = df_full['SBD'].astype(str)
    
    # Sheet 1: TRAO GIẢI - Học sinh có trong file TRAO GIẢI
    df_sheet1 = df_full[df_full['SBD_str'].isin(sbd_trao_giai)].copy()
    df_sheet1 = df_sheet1.drop('SBD_str', axis=1)  # Xóa cột tạm
    print(f"   ✅ Sheet 1 (TRAO GIẢI): {len(df_sheet1)} học sinh")
    
    # Sheet 2: KO ĐK - Học sinh KHÔNG có trong file TRAO GIẢI
    df_sheet2 = df_full[~df_full['SBD_str'].isin(sbd_trao_giai)].copy()
    df_sheet2 = df_sheet2.drop('SBD_str', axis=1)  # Xóa cột tạm
    print(f"   ✅ Sheet 2 (KO ĐK): {len(df_sheet2)} học sinh")
    
    # Kiểm tra tổng
    total_check = len(df_sheet1) + len(df_sheet2)
    print(f"   📊 Tổng kiểm tra: {total_check} (= {len(df_full)}? {total_check == len(df_full)})")
    
    # 4. Lưu file kết quả
    print(f"\n4️⃣ Lưu file kết quả...")
    print(f"   💾 {file_output}")
    
    with pd.ExcelWriter(file_output, engine='openpyxl') as writer:
        df_sheet1.to_excel(writer, sheet_name='TRAO GIẢI', index=False)
        df_sheet2.to_excel(writer, sheet_name='KO ĐK', index=False)
    
    print(f"   ✅ Lưu thành công!")
    
    # 5. Thống kê
    print(f"\n5️⃣ Thống kê:")
    print(f"   📊 File TRAO GIẢI: {len(df_trao_giai)} học sinh")
    print(f"   📊 File Full: {len(df_full)} học sinh")
    print(f"   📊 Sheet 'TRAO GIẢI': {len(df_sheet1)} học sinh ({len(df_sheet1)/len(df_full)*100:.1f}%)")
    print(f"   📊 Sheet 'KO ĐK': {len(df_sheet2)} học sinh ({len(df_sheet2)/len(df_full)*100:.1f}%)")
    
    # Hiển thị mẫu
    print(f"\n📋 Mẫu Sheet 1 - TRAO GIẢI (5 học sinh đầu):")
    if len(df_sheet1) > 0:
        display_cols = ['SBD', 'FULL NAME', 'KHỐI', 'TRƯỜNG']
        available_cols = [col for col in display_cols if col in df_sheet1.columns]
        print(df_sheet1[available_cols].head(5).to_string(index=False))
    else:
        print("   (Không có dữ liệu)")
    
    print(f"\n📋 Mẫu Sheet 2 - KO ĐK (5 học sinh đầu):")
    if len(df_sheet2) > 0:
        display_cols = ['SBD', 'FULL NAME', 'KHỐI', 'TRƯỜNG']
        available_cols = [col for col in display_cols if col in df_sheet2.columns]
        print(df_sheet2[available_cols].head(5).to_string(index=False))
    else:
        print("   (Không có dữ liệu)")
    
    print("\n" + "="*80)
    print("✅ HOÀN THÀNH!")
    print("="*80)
    print(f"\n📁 File kết quả: {file_output}")
    print(f"   - Sheet 1: TRAO GIẢI ({len(df_sheet1)} học sinh)")
    print(f"   - Sheet 2: KO ĐK ({len(df_sheet2)} học sinh)")
    
except Exception as e:
    print(f"\n❌ Lỗi: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)
