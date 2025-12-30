
import pandas as pd
import PyPDF2

excel_path = '/Users/migueldlg/Downloads/Para ordenar/To Ship order-2025-12-29-18_20.xlsx'
pdf_path = '/Users/migueldlg/Downloads/Para ordenar/etiquetas.pdf'

print("--- EXCEL HEADER ---")
try:
    df = pd.read_excel(excel_path)
    print(df.columns.tolist())
    print(df.head(3))
except Exception as e:
    print(f"Error reading Excel: {e}")

print("\n--- PDF CONTENT (First Page) ---")
try:
    with open(pdf_path, 'rb') as f:
        reader = PyPDF2.PdfReader(f)
        if len(reader.pages) > 0:
            page = reader.pages[0]
            text = page.extract_text()
            print(text[:1000])  # Print first 1000 chars
        else:
            print("PDF is empty")
except Exception as e:
    print(f"Error reading PDF: {e}")
