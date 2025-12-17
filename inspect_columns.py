import pandas as pd

# Load the Excel file with header=1 (skipping the first row)
file_path = 'pedidos shein.xlsx'
try:
    df = pd.read_excel(file_path, header=1)
    print("Columns found:")
    print(df.columns.tolist())
    print("\nFirst few rows:")
    print(df.head())
except Exception as e:
    print(f"Error reading file: {e}")
