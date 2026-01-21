import pandas as pd

# Check all sheets
xl = pd.ExcelFile('DS KQ V2 VBD.xlsx')
print(f"Số sheet: {len(xl.sheet_names)}")
print(f"Tên các sheet: {xl.sheet_names}")

# Read all sheets
for sheet_name in xl.sheet_names:
    print(f"\n\n{'='*60}")
    print(f"SHEET: {sheet_name}")
    print('='*60)
    df = pd.read_excel('DS KQ V2 VBD.xlsx', sheet_name=sheet_name)
    print(f"Số cột: {len(df.columns)}")
    print("Các cột:")
    for i, col in enumerate(df.columns):
        print(f"  {i}: '{col}'")
    print("\nDữ liệu mẫu:")
    print(df.head(2))
