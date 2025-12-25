import pandas as pd

files = [
    'Pedidos tiktok.xlsx',
    'To Ship order-2025-12-23-19_56.xlsx'
]

for f in files:
    print(f"\n--- Headers for {f} ---")
    try:
        df = pd.read_excel(f)
        print(list(df.columns))
        # Print first row to see if there's a header offset issue
        print("\nFirst row data:")
        print(df.iloc[0].to_dict())
    except Exception as e:
        print(f"Error reading {f}: {e}")
