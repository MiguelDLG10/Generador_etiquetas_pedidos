import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

def generate_labels_and_summary(input_file, output_file):
    # Load data
    try:
        df = pd.read_excel(input_file, header=1)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    # Get unique order numbers while preserving original order
    # This ensures labels are generated in the same order as the Excel file
    unique_orders = df['Número de pedido'].drop_duplicates().tolist()

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
        # Get all rows for this order
        group = df[df['Número de pedido'] == order_id]
        # Set page size for label
        c.setPageSize((label_width, label_height))
        
        # Extract Order Level Info (take from first row of group)
        first_row = group.iloc[0]
        paquete_vendedor = str(first_row.get('Paquete del vendedor', 'N/A'))
        numero_guia = str(first_row.get('Número de guía', 'N/A'))
        
        # Helper function to draw header
        def draw_header(c, margin, label_height, paquete_vendedor, numero_guia):
            y_pos = label_height - margin - 10
            
            # Paquete del vendedor
            c.setFont("Helvetica-Bold", 10)
            c.drawString(margin, y_pos, f"Paq: {paquete_vendedor}")
            y_pos -= 12
            
            # Número de guía
            c.setFont("Helvetica-Bold", 12)
            c.drawString(margin, y_pos, f"Guía: {numero_guia}")
            y_pos -= 14
            
            # Items Header
            c.setFont("Helvetica", 8)
            c.drawString(margin, y_pos, "SKU / Cantidad")
            y_pos -= 10
            
            return y_pos

        # Initial setup for this order
        c.setPageSize((label_width, label_height))
        y_pos = draw_header(c, margin, label_height, paquete_vendedor, numero_guia)
        
        # List Items
        c.setFont("Helvetica", 9)
        
        # Count occurrences of each SKU in this order
        order_skus = group['SKU del vendedor'].value_counts()
        
        for sku, qty in order_skus.items():
            sku = str(sku)
            
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
            y_pos -= 12 # Increased spacing
            
            # Check if we ran out of space
            if y_pos < margin + 5: # Add a bit of buffer for page number
                # Add page number to current page
                c.setFont("Helvetica", 6)
                page_text = f"{page_number}"
                text_width = c.stringWidth(page_text, "Helvetica", 6)
                c.drawString((label_width - text_width) / 2, margin / 2, page_text)
                
                c.showPage()
                page_number += 1
                
                # Start new page for same order
                c.setPageSize((label_width, label_height))
                y_pos = draw_header(c, margin, label_height, paquete_vendedor, numero_guia)
                c.setFont("Helvetica", 9) # Reset font for items
        
        # Add page number at the bottom center of the last page for this order
        c.setFont("Helvetica", 6)
        page_text = f"{page_number}"
        text_width = c.stringWidth(page_text, "Helvetica", 6)
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

if __name__ == "__main__":
    input_excel = 'pedidos shein.xlsx'
    output_pdf = 'etiquetas_pedidos.pdf'
    generate_labels_and_summary(input_excel, output_pdf)
