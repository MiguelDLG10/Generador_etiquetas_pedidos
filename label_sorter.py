import pandas as pd
import PyPDF2
import re
import os

def normalize_text(text):
    """Normalize text data: remove hyphens and spaces."""
    if pd.isna(text):
        return ""
    return str(text).replace("-", "").replace(" ", "").strip()

def sort_tiktok_labels(excel_path, pdf_path, output_pdf_path):
    """
    Sorts PDF labels based on 'Tracking ID' from Excel file.
    """
    stats = {
        'total_excel_ids': 0,
        'matched_pages': 0,
        'missing_ids': [],
        'unmatched_pages': 0,
        'success': False,
        'error': None
    }

    print("Reading Excel file for sorting...")
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        stats['error'] = f"Error reading Excel: {e}"
        return stats

    # Check for 'Tracking ID' column
    if 'Tracking ID' not in df.columns:
        stats['error'] = "Error: 'Tracking ID' column not found in Excel. Cannot sort labels."
        return stats

    # Extract relevant IDs, keeping order and removing duplicates
    # Use a dictionary to keep order and remove duplicates
    target_ids = list(dict.fromkeys(df['Tracking ID'].dropna()))
    
    # Normalize target IDs for matching
    normalized_target_ids = [normalize_text(tid) for tid in target_ids]
    
    stats['total_excel_ids'] = len(target_ids)
    print(f"Found {len(target_ids)} unique Tracking IDs in Excel.")

    print("Reading PDF file...")
    try:
        reader = PyPDF2.PdfReader(pdf_path)
        total_pages = len(reader.pages)
        print(f"Total pages in PDF: {total_pages}")
    except Exception as e:
        stats['error'] = f"Error reading PDF: {e}"
        return stats

    # Map Normalized ID -> Page Object
    id_to_pages = {nid: [] for nid in normalized_target_ids}
    
    print("Indexing PDF pages...")
    unmatched_pages_count = 0
    
    for i, page in enumerate(reader.pages):
        text = page.extract_text()
        normalized_page_text = normalize_text(text)
        
        found = False
        for nid in normalized_target_ids:
            # Check if ID is in page text
            # Note: normalized_ids usually don't have spaces, page text might or might not.
            # Normalizing page text handles spaces/hyphens removal.
            if len(nid) > 5 and nid in normalized_page_text: # Ensure ID is not too short to avoid false positives?
                id_to_pages[nid].append(page)
                found = True
                break 
            elif nid in normalized_page_text: # Fallback for shorter IDs logic if necessary, or just standard check
                 id_to_pages[nid].append(page)
                 found = True
                 break
        
        if not found:
            unmatched_pages_count += 1

    stats['unmatched_pages'] = unmatched_pages_count

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
                # Count distinct labels/IDs matched, not just total pages
                # If multiple IDs are on one page, we might add the page multiple times, which is standard behavior for 'per order' printing
        else:
            missing_ids.append(nid)

    # Calculate matched labels count (unique IDs found)
    stats['matched_pages'] = added_count # This is actually pages added
    stats['matched_ids_count'] = len(normalized_target_ids) - len(missing_ids)


    # Save Output
    print(f"Writing output to {output_pdf_path}...")
    try:
        with open(output_pdf_path, "wb") as f:
            writer.write(f)
        stats['success'] = True
    except Exception as e:
        stats['error'] = f"Error writing output PDF: {e}"

    return stats
