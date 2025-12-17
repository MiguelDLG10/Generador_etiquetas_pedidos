import pandas as pd

# Load the Excel file
file_path = 'pedidos shein.xlsx'
try:
    df = pd.read_excel(file_path)
    print("File read successfully!")
    print(df.head())
except Exception as e:
    print(f"Error reading file: {e}")
