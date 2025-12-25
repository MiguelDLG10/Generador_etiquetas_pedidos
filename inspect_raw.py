import pandas as pd

f = 'To Ship order-2025-12-23-19_56.xlsx'
print(f"--- Raw dump for {f} ---")
try:
    df = pd.read_excel(f, header=None)
    print(df.head(5))
except Exception as e:
    print(f"Error reading {f}: {e}")
