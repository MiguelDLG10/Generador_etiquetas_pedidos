import pandas as pd

files = ['pedido_3.xlsx']

for f in files:
    print(f"\n--- Headers for {f} ---")
    try:
        df = pd.read_excel(f)
        print(list(df.columns))
        print("\nFirst row data:")
        print(df.iloc[0].to_dict())
    except Exception as e:
        print(f"Error reading {f}: {e}")
