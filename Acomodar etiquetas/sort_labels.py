
import pandas as pd
import PyPDF2
import re
import os

# Configuration
excel_path = '/Users/migueldlg/Downloads/Para ordenar/To Ship order-2025-12-29-18_20.xlsx'
pdf_path = '/Users/migueldlg/Downloads/Para ordenar/etiquetas.pdf'
output_pdf_path = '/Users/migueldlg/Downloads/Para ordenar/etiquetas_ordenadas.pdf'

def normalize_text(text):
    """Normalize text data: remove hyphens and spaces."""
    if pd.isna(text):
        return ""
    return str(text).replace("-", "").replace(" ", "").strip()

def sort_pdf():
    print("Reading Excel file...")
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        print(f"Error reading Excel: {e}")
        return

    # Check for 'Tracking ID' column
    if 'Tracking ID' not in df.columns:
        print("Error: 'Tracking ID' column not found in Excel.")
        return

    # Extract relevant IDs, keeping order and removing duplicates
    # Use a dictionary to keep order and remove duplicates (Python 3.7+ dicts are ordered)
    target_ids = list(dict.fromkeys(df['Tracking ID'].dropna()))
    
    # Normalize target IDs for matching
    normalized_target_ids = [normalize_text(tid) for tid in target_ids]
    
    print(f"Found {len(target_ids)} unique Tracking IDs in Excel.")

    print("Reading PDF file...")
    try:
        reader = PyPDF2.PdfReader(pdf_path)
        total_pages = len(reader.pages)
        print(f"Total pages in PDF: {total_pages}")
    except Exception as e:
        print(f"Error reading PDF: {e}")
        return

    # Map Normalized ID -> Page Object
    # Since one ID might appear on multiple pages (though unlikely for labels, good to be safe) or vice versa,
    # but the requirement implies 1 label = 1 page usually.
    # We will search each page for our target IDs.
    
    # Strategy: 
    # 1. Iterate through all pages.
    # 2. Extract text from each page.
    # 3. Normalize page text.
    # 4. Check which target ID is present in the page text.
    # Optimization: Store page objects in a way we can retrieve them by ID.
    
    id_to_pages = {nid: [] for nid in normalized_target_ids}
    
    print("Indexing PDF pages (this might take a moment)...")
    unmatched_pages = []
    
    for i, page in enumerate(reader.pages):
        text = page.extract_text()
        normalized_page_text = normalize_text(text)
        
        found = False
        for nid in normalized_target_ids:
            if nid in normalized_page_text:
                id_to_pages[nid].append(page)
                found = True
                # Break if we assume one ID per page to speed up? 
                # Let's not break, just in case multiple IDs match somehow (subsets), 
                # but usually longer specific IDs are unique.
                # Given strict matching requirements, let's stop searching once found to avoid false positives if IDs are substrings of each other?
                # Actually, let's stick to finding the first match.
                break 
        
        if not found:
            unmatched_pages.append(i)

    # Create Writer
    writer = PyPDF2.PdfWriter()
    
    added_count = 0
    missing_ids = []

    print("Reordering pages...")
    for nid in normalized_target_ids:
        pages = id_to_pages.get(nid, [])
        if pages:
            for page in pages:
                writer.add_page(page)
                added_count += 1
        else:
            missing_ids.append(nid)

    # Save Output
    print(f"Writing output to {output_pdf_path}...")
    with open(output_pdf_path, "wb") as f:
        writer.write(f)

    print("-" * 30)
    print(f"Process Complete.")
    print(f"Total Excel IDs: {len(normalized_target_ids)}")
    print(f"Pages added to new PDF: {added_count}")
    print(f"IDs not found in PDF: {len(missing_ids)}")
    if unmatched_pages:
        print(f"Warning: {len(unmatched_pages)} pages from original PDF were not matched to any ID (possibly cover sheets or parse failures).")

if __name__ == "__main__":
    sort_pdf()
