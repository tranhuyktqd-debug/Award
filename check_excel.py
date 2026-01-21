import pandas as pd

# Read Excel file
df = pd.read_excel('DS KQ V2 VBD.xlsx')

print("Tất cả các cột trong file:")
for i, col in enumerate(df.columns):
    print(f"{i}: '{col}'")

print("\n\nDữ liệu 3 dòng đầu:")
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 200)
print(df.head(3))

print("\n\nKiểm tra cột Unnamed:")
unnamed_cols = [col for col in df.columns if 'Unnamed' in str(col)]
if unnamed_cols:
    print(f"Có {len(unnamed_cols)} cột Unnamed: {unnamed_cols}")
    print("\nDữ liệu trong các cột Unnamed:")
    print(df[unnamed_cols].head(3))
