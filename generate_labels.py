import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

def load_and_normalize_data(file_path):
    """
    Load data from Excel and normalize + return stats
    
    Returns:
        (DataFrame, dict): Normalized data and processing stats
    """
    stats = {
        'total_rows': 0,
        'valid_rows': 0,
        'dropped_rows': 0,
        'drop_reasons': [],
        'format_detected': 'Unknown'
    }

    # Try reading as TikTok first (header=0 usually)
    try:
        df = pd.read_excel(file_path, header=0)
        # Check for TikTok specific columns
        if 'Order ID' in df.columns and 'Seller SKU' in df.columns:
            print("Detected TikTok format")
            stats['format_detected'] = 'TikTok'
            stats['total_rows'] = len(df)
            
            normalized = pd.DataFrame()
            normalized['order_id'] = df['Order ID']
            normalized['package_id'] = df['Package ID'] if 'Package ID' in df.columns else 'N/A'
            normalized['tracking_id'] = df['Tracking ID'] if 'Tracking ID' in df.columns else 'N/A'
            
            # Fill NaNs in these columns immediately to prevent groupby dropping them later
            normalized['package_id'] = normalized['package_id'].fillna('N/A')
            normalized['tracking_id'] = normalized['tracking_id'].fillna('N/A')

            normalized['sku'] = df['Seller SKU']
            normalized['quantity'] = pd.to_numeric(df['Quantity'], errors='coerce')
            normalized['source'] = 'TIKTOK'
            
            # Identify drops
            # 1. Missing Quantity
            missing_qty = normalized['quantity'].isna()
            if missing_qty.any():
                drop_count = missing_qty.sum()
                stats['dropped_rows'] += int(drop_count)
                stats['drop_reasons'].append(f"{drop_count} rows dropped due to invalid/missing Quantity")
            
            # Drop rows with invalid quantity
            normalized = normalized.dropna(subset=['quantity'])
            
            stats['valid_rows'] = len(normalized)
            return normalized, stats
    except Exception:
        pass

    # Try reading as Shein (header=1)
    try:
        df = pd.read_excel(file_path, header=1)
        if 'Número de pedido' in df.columns and 'SKU del vendedor' in df.columns:
            print("Detected Shein format")
            stats['format_detected'] = 'Shein'
            stats['total_rows'] = len(df)
            
            normalized = pd.DataFrame()
            normalized['order_id'] = df['Número de pedido']
            normalized['package_id'] = df['Paquete del vendedor'].fillna('N/A')
            normalized['tracking_id'] = df['Número de guía'].fillna('N/A')
            normalized['sku'] = df['SKU del vendedor']
            
            # Check for potential drops (though count is forced to 1, effectively keeping all valid execution rows)
            # If SKU is missing, that's a problem
            missing_sku = normalized['sku'].isna()
            if missing_sku.any():
                 drop_count = missing_sku.sum()
                 stats['dropped_rows'] += int(drop_count)
                 stats['drop_reasons'].append(f"{drop_count} rows with missing SKU")
                 normalized = normalized.dropna(subset=['sku'])

            normalized['quantity'] = 1 
            normalized['source'] = 'SHEIN'
            
            stats['valid_rows'] = len(normalized)
            return normalized, stats
    except Exception as e:
        print(f"Error reading as Shein: {e}")

    raise ValueError("File format not recognized (neither TikTok with 'Order ID' nor Shein with 'Número de pedido' found)")

def generate_labels_and_summary(input_file, output_file):
    # Load and normalize data
    # Let exceptions propagate to the UI
    df, stats = load_and_normalize_data(input_file)

    # Aggregate data by order_id, package_id, tracking_id, sku, source
    # This sums up quantities for the same SKU in the same order
    df_agg = df.groupby(['order_id', 'package_id', 'tracking_id', 'sku', 'source'])['quantity'].sum().reset_index()

    # Get unique orders preserving order of appearance is a bit trickier after groupby
    # We can get unique orders from the normalized df before aggregation if we want strict original order
    # But usually sorting by something or just taking unique from agg is fine.
    # To be safe and close to original behavior:
    unique_orders = df['order_id'].drop_duplicates().tolist()

    # Create Canvas
    c = canvas.Canvas(output_file)
    
    # Label Dimensions
    label_width = 63 * mm
    label_height = 38 * mm
    margin = 2 * mm

    # Data for summary
    sku_summary = {}
    
    # Page counter
    page_number = 1

    print(f"Generating labels for {len(unique_orders)} orders...")

    for order_id in unique_orders:
        # Get all rows for this order from aggregated df
        group = df_agg[df_agg['order_id'] == order_id]
        
        if group.empty:
            continue

        # Extract Order Level Info (take from first row of group)
        first_row = group.iloc[0]
        paquete_vendedor = str(first_row['package_id']) if pd.notna(first_row['package_id']) else 'N/A'
        numero_guia = str(first_row['tracking_id']) if pd.notna(first_row['tracking_id']) else 'N/A'
        source_app = str(first_row['source'])
        
        # Helper function to draw header
        def draw_header(c, margin, label_width, label_height, paquete_vendedor, numero_guia):
            y_pos = label_height - margin - 10
            
            # Paquete del vendedor
            c.setFont("Helvetica-Bold", 10)
            c.drawString(margin, y_pos, f"Paq: {paquete_vendedor}")
            y_pos -= 12
            
            # Número de guía
            c.setFont("Helvetica-Bold", 12)
            guia_text = f"Guía: {numero_guia}"
            
            # Check width and wrap if necessary
            # Label width is 63mm, margin 2mm. Usable approx 59mm.
            # 12pt font is approx 4.2mm high.
            # We can use c.stringWidth to check.
            max_width = label_width - (2 * margin)
            
            if c.stringWidth(guia_text, "Helvetica-Bold", 12) > max_width:
                # Basic wrap strategy: split at arbitrary point or space
                # Since tracking numbers are often continuous, we might force split
                # Try to fit as much as possible
                split_idx = int(len(guia_text) * (max_width / c.stringWidth(guia_text, "Helvetica-Bold", 12)))
                # Adjust slightly to be safe
                split_idx = max(5, split_idx - 2) 
                
                line1 = guia_text[:split_idx]
                line2 = guia_text[split_idx:]
                
                c.drawString(margin, y_pos, line1)
                y_pos -= 12
                c.drawString(margin, y_pos, line2)
                y_pos -= 14
            else:
                c.drawString(margin, y_pos, guia_text)
                y_pos -= 14
            
            # Items Header
            # Header removed to save space
            
            return y_pos

        # Initial setup for this order
        c.setPageSize((label_width, label_height))
        y_pos = draw_header(c, margin, label_width, label_height, paquete_vendedor, numero_guia)
        
        # List Items
        c.setFont("Helvetica", 12) # Reduced to 12 non-bold
        
        # Iterate over aggregated SKUs
        for _, row in group.iterrows():
            sku = str(row['sku'])
            qty = int(row['quantity'])
            
            # Add to summary
            if sku in sku_summary:
                sku_summary[sku] += qty
            else:
                sku_summary[sku] = qty
            
            # Draw Item Line
            # Truncate SKU if too long
            display_sku = (sku[:25] + '..') if len(sku) > 25 else sku
            text = f"{display_sku}  (x{qty})"
            
            c.drawString(margin, y_pos, text)
            y_pos -= 15 # Adjusted spacing for font size 12
            
            # Check if we ran out of space
            if y_pos < margin + 8: # Increased buffer for larger footer
                # Add page number to current page
                c.setFont("Helvetica-Bold", 10) # Larger font
                page_text = f"{page_number} - {source_app}"
                text_width = c.stringWidth(page_text, "Helvetica-Bold", 10)
                c.drawString((label_width - text_width) / 2, margin / 2, page_text)
                
                c.showPage()
                page_number += 1
                
                # Start new page for same order
                c.setPageSize((label_width, label_height))
                y_pos = draw_header(c, margin, label_width, label_height, paquete_vendedor, numero_guia)
                c.setFont("Helvetica", 12) # Reset font for items
        
        # Add page number at the bottom center of the last page for this order
        c.setFont("Helvetica-Bold", 10) # Larger font
        page_text = f"{page_number} - {source_app}"
        text_width = c.stringWidth(page_text, "Helvetica-Bold", 10)
        c.drawString((label_width - text_width) / 2, margin / 2, page_text)
        
        c.showPage()
        page_number += 1

    # Summary Section
    print("Generating summary page...")
    c.setPageSize(A4)
    width, height = A4
    
    c.setFont("Helvetica-Bold", 14)
    c.drawString(20 * mm, height - 20 * mm, "Lista de Picking (Resumen de SKUs)")
    
    y_pos = height - 30 * mm
    c.setFont("Helvetica-Bold", 10)
    c.drawString(20 * mm, y_pos, "SKU del Vendedor")
    c.drawString(120 * mm, y_pos, "Cantidad Total")
    y_pos -= 5 * mm
    c.line(20 * mm, y_pos, 180 * mm, y_pos)
    y_pos -= 5 * mm
    
    c.setFont("Helvetica", 10)
    
    # Sort summary by SKU
    sorted_skus = sorted(sku_summary.items())
    
    for sku, total_qty in sorted_skus:
        sku_str = str(sku)
        c.drawString(20 * mm, y_pos, sku_str)
        
        # Calculate width of SKU text to start line right after it
        text_width = c.stringWidth(sku_str, "Helvetica", 10)
        start_line_x = 20 * mm + text_width + 2 * mm # 2mm padding
        
        # Draw dashed line connecting SKU and Quantity
        c.saveState()
        c.setDash(1, 3)
        c.line(start_line_x, y_pos + 1.5, 115 * mm, y_pos + 1.5)
        c.restoreState()
        
        c.drawString(120 * mm, y_pos, str(total_qty))
        y_pos -= 10 * mm # Increased spacing
        
        if y_pos < 20 * mm:
            c.showPage()
            c.setPageSize(A4)
            y_pos = height - 20 * mm
            c.setFont("Helvetica", 10)

    c.save()
    print(f"PDF generated: {output_file}")
    
    stats['unique_orders'] = len(unique_orders)
    return stats

if __name__ == "__main__":
    # Test block
    input_excel = 'pedidos shein.xlsx'
    output_pdf = 'etiquetas_pedidos_test.pdf'
    generate_labels_and_summary(input_excel, output_pdf)
