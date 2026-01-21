import pandas as pd

# Đọc với dtype str cho SBD để giữ số 0 đầu
df = pd.read_excel('DS_KQ_WITH_QR.xlsx', dtype={'SBD': str})

print("Sample SBD values:")
print(df['SBD'].head())
print(f"\nSBD type: {df['SBD'].dtype}")
print(f"\nFirst SBD: '{df['SBD'].iloc[0]}'")
print(f"Length: {len(df['SBD'].iloc[0])}")
