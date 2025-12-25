import pandas as pd

f = 'To Ship order-2025-12-23-19_56.xlsx'
try:
    df = pd.read_excel(f, header=None)
    print(f"Shape: {df.shape}")
    print("Row 0, Col 0 value:", repr(df.iloc[0, 0]))
    print("Row 2, Col 0 value:", repr(df.iloc[2, 0]))
except Exception as e:
    print(f"Error: {e}")
