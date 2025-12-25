import pandas as pd

f = 'To Ship order-2025-12-23-19_56.xlsx'
try:
    xl = pd.ExcelFile(f)
    print(f"Sheet names: {xl.sheet_names}")
    for sheet in xl.sheet_names:
        df = pd.read_excel(f, sheet_name=sheet)
        print(f"--- Sheet: {sheet} ---")
        print(f"Columns: {list(df.columns)}")
except Exception as e:
    print(f"Error: {e}")
