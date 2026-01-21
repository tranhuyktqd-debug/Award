# -*- coding: utf-8 -*-
"""
Script kiểm tra cấu trúc file Excel
"""
import pandas as pd
import sys

# Set UTF-8 encoding for console
if sys.platform == 'win32':
    try:
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='ignore')
    except:
        pass

# Đọc file Excel
file_path = r'D:\1 CHUAN BI KY THI\TEST_TRA_CUU_TRAO_GIAI\DS điểm danh - Chị Hòa.xlsx'

try:
    df = pd.read_excel(file_path, nrows=10)
    
    print("="*60)
    print("📊 THÔNG TIN FILE EXCEL")
    print("="*60)
    print(f"\n📁 File: {file_path}")
    print(f"📋 Tổng số dòng: {len(df)}")
    print(f"📋 Tổng số cột: {len(df.columns)}")
    
    print("\n" + "="*60)
    print("📋 TÊN CÁC CỘT (INDEX):")
    print("="*60)
    for idx, col in enumerate(df.columns):
        print(f"Cột {idx} ('{chr(65+idx)}'): {col}")
    
    print("\n" + "="*60)
    print("📋 DỮ LIỆU MẪU (5 dòng đầu):")
    print("="*60)
    print(df.head(5).to_string())
    
    print("\n" + "="*60)
    print("📋 DỮ LIỆU CÁC CỘT D, F, G, H, I, J:")
    print("="*60)
    # Cột D=3, F=5, G=6, H=7, I=8, J=9 (index bắt đầu từ 0)
    cols_to_show = [3, 5, 6, 7, 8, 9]
    for idx in cols_to_show:
        if idx < len(df.columns):
            print(f"\nCột {chr(65+idx)} (index {idx}): {df.columns[idx]}")
            print(f"Dữ liệu mẫu: {df.iloc[0:3, idx].tolist()}")
    
except Exception as e:
    print(f"❌ Lỗi: {e}")
    import traceback
    traceback.print_exc()

