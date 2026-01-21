import pandas as pd

# Đọc file Excel
df = pd.read_excel('DATA KQ.xlsx', sheet_name='Sheet1')

print('='*60)
print('=== CẤU TRÚC FILE EXCEL ===')
print('='*60)
print(f'Tổng số dòng: {len(df)}')

print(f'\nCác cột có sẵn:')
for i, col in enumerate(df.columns.tolist(), 1):
    print(f'{i:2d}. {col}')

print(f'\n=== MẪU DỮ LIỆU 3 DÒNG ĐẦU ===')
print(df.head(3).to_string())

# Kiểm tra cột EMAIL
print(f'\n=== KIỂM TRA CỘT EMAIL ===')
has_email = 'EMAIL' in df.columns
has_email_phhs = 'EMAIL_PHHS' in df.columns

print(f'Có cột EMAIL: {has_email}')
print(f'Có cột EMAIL_PHHS: {has_email_phhs}')

if has_email:
    non_empty = df['EMAIL'].notna().sum()
    print(f'Số dòng có email: {non_empty}/{len(df)} ({non_empty/len(df)*100:.1f}%)')
    print(f'\n5 email mẫu:')
    print(df[df['EMAIL'].notna()]['EMAIL'].head(5).tolist())
elif has_email_phhs:
    non_empty = df['EMAIL_PHHS'].notna().sum()
    print(f'Số dòng có email: {non_empty}/{len(df)} ({non_empty/len(df)*100:.1f}%)')
else:
    print('⚠️ CẢNH BÁO: Không tìm thấy cột EMAIL hoặc EMAIL_PHHS!')
    print('Cần thêm cột EMAIL vào file Excel để gửi email.')
