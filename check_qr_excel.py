import pandas as pd

df1 = pd.read_excel('DS_KQ_WITH_QR.xlsx')
df2 = pd.read_excel('DATA KQ.xlsx')

print(f'DS_KQ_WITH_QR: {len(df1)} rows')
print(f'DATA KQ: {len(df2)} rows')
print(f'\nColumns DS_KQ_WITH_QR:')
print(list(df1.columns))
print(f'\nSample QR DATA:')
print(df1["QR DATA"].iloc[0][:300])
