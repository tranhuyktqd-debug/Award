"""
Script merge dữ liệu:
- Lấy QR DATA từ DS_KQ_WITH_QR.xlsx
- Lấy EMAIL từ DATA KQ.xlsx
- Merge theo SBD và tạo file mới
"""
import pandas as pd

print("🔄 MERGE DỮ LIỆU QR VÀ EMAIL...")
print("="*60)

# Đọc file có QR
df_qr = pd.read_excel('DS_KQ_WITH_QR.xlsx')
print(f"✅ Đọc DS_KQ_WITH_QR.xlsx: {len(df_qr)} rows")

# Đọc file có EMAIL
df_email = pd.read_excel('DATA KQ.xlsx')
print(f"✅ Đọc DATA KQ.xlsx: {len(df_email)} rows")

# Kiểm tra cột SBD
print(f"\n📋 Columns DS_KQ_WITH_QR: {list(df_qr.columns)}")
print(f"📋 Columns DATA KQ: {list(df_email.columns)[:15]}")

# Merge theo SBD
print(f"\n🔗 Merge dữ liệu theo SBD...")
df_merged = df_email.merge(
    df_qr[['SBD', 'QR DATA']], 
    on='SBD', 
    how='left'
)

print(f"✅ Merge thành công: {len(df_merged)} rows")

# Kiểm tra
has_qr = df_merged['QR DATA'].notna().sum()
has_email = df_merged['EMAIL'].notna().sum() if 'EMAIL' in df_merged.columns else 0

print(f"\n📊 THỐNG KÊ:")
print(f"  - Có QR DATA: {has_qr}/{len(df_merged)}")
print(f"  - Có EMAIL: {has_email}/{len(df_merged)}")

# Lưu file mới
output_file = 'DATA_KQ_FULL_WITH_QR.xlsx'
df_merged.to_excel(output_file, index=False)
print(f"\n💾 Đã lưu: {output_file}")
print("="*60)
print("\n✅ HOÀN TẤT!")
print(f"📝 File mới: {output_file}")
print(f"   - Có đầy đủ {len(df_merged)} học sinh")
print(f"   - Có EMAIL để gửi")
print(f"   - Có QR DATA cho {has_qr} học sinh")
