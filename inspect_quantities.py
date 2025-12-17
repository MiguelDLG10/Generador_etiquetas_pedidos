import pandas as pd

# Load the Excel file
file_path = 'pedidos shein.xlsx'
try:
    df = pd.read_excel(file_path, header=1)
    print("Inspecting 'Cantidad De Envío De Actividad':")
    print(df['Cantidad De Envío De Actividad'].head(20))
    print("\nUnique values in 'Cantidad De Envío De Actividad':")
    print(df['Cantidad De Envío De Actividad'].unique())
    
    # Also check if there are duplicate SKUs in the same order
    print("\nChecking for duplicate SKUs in orders:")
    grouped = df.groupby(['Número de pedido', 'SKU del vendedor']).size()
    print(grouped[grouped > 1].head())
except Exception as e:
    print(f"Error reading file: {e}")
