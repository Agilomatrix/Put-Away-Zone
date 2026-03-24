import streamlit as st
import pandas as pd
import os
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph, PageBreak
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.graphics.barcode import code128
from reportlab.graphics.shapes import Drawing
import re
import tempfile

# Define sticker dimensions
STICKER_WIDTH = 10 * cm
STICKER_HEIGHT = 15 * cm
STICKER_PAGESIZE = (STICKER_WIDTH, STICKER_HEIGHT)
CONTENT_BOX_WIDTH = 10 * cm
CONTENT_BOX_HEIGHT = 10.0 * cm  # Increased to accommodate more fields

# Styles
bold_style = ParagraphStyle(name='Bold', fontName='Helvetica-Bold', fontSize=12, alignment=TA_CENTER)
desc_style = ParagraphStyle(name='Desc', fontName='Helvetica', fontSize=10, alignment=TA_CENTER, leading=11)
val_style = ParagraphStyle(name='Value', fontName='Helvetica-Bold', fontSize=14, alignment=TA_CENTER)

def generate_barcode(data_string, width):
    """Generate a Code128 barcode"""
    try:
        # Code128 can be wide, so we adjust barWidth to fit the box
        bc = code128.Code128(data_string, barHeight=1.0*cm, barWidth=0.7)
        d = Drawing(width, 1.2*cm)
        # Center the barcode in the drawing
        d.add(bc)
        return d
    except:
        return Paragraph("Barcode Error", desc_style)

def parse_location_string(location_str):
    location_parts = [''] * 4
    if not location_str or not isinstance(location_str, str): return location_parts
    matches = re.findall(r'([^_\s]+)', str(location_str))
    for i, match in enumerate(matches[:4]): location_parts[i] = match
    return location_parts

def generate_sticker_labels(df):
    def draw_border(canvas, doc):
        canvas.saveState()
        canvas.setLineWidth(1.8)
        # Draw border around the content
        canvas.rect(0.1*cm, STICKER_HEIGHT - 11.5*cm, STICKER_WIDTH - 0.2*cm, 11.3*cm)
        canvas.restoreState()

    temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    temp_pdf_path = temp_pdf.name
    
    doc = SimpleDocTemplate(temp_pdf_path, pagesize=STICKER_PAGESIZE,
                          topMargin=0.5*cm, leftMargin=0.1*cm, rightMargin=0.1*cm)

    all_elements = []
    cols = [c.upper() for c in df.columns]

    for index, row in df.iterrows():
        # Map columns (Flexible naming)
        grn_no = str(row.get(next((c for c in df.columns if 'GRN' in c.upper() and 'DATE' not in c.upper()), df.columns[0]), ""))
        grn_date = str(row.get(next((c for c in df.columns if 'DATE' in c.upper()), ""), ""))
        part_no = str(row.get(next((c for c in df.columns if 'PART' in c.upper()), ""), ""))
        desc = str(row.get(next((c for c in df.columns if 'DESC' in c.upper() or 'NAME' in c.upper()), ""), ""))
        qty = str(row.get(next((c for c in df.columns if 'QTY' in c.upper() or 'QUANTITY' in c.upper()), ""), ""))
        loc_raw = str(row.get(next((c for c in df.columns if 'LOC' in c.upper()), ""), ""))
        
        loc_parts = parse_location_string(loc_raw)
        
        # Data for Barcode (Scans all fields)
        barcode_data = f"{grn_no}|{part_no}|{qty}|{loc_raw}"
        
        # TABLE STRUCTURE (To prevent overlapping)
        content_w = STICKER_WIDTH - 0.4*cm
        
        # Create Store Location Sub-table
        loc_table = Table([loc_parts], colWidths=[(content_w*0.66)/4]*4, rowHeights=[0.8*cm])
        loc_table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.black),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))

        main_data = [
            ["GRN No", Paragraph(grn_no, val_style)],
            ["GRN Date", Paragraph(grn_date.split(' ')[0], val_style)],
            ["Part No", Paragraph(part_no, val_style)],
            ["Description", Paragraph(desc[:60], desc_style)],
            ["Quantity", Paragraph(qty, val_style)],
            ["Store Location", loc_table],
            ["Barcode", generate_barcode(barcode_data, content_w)]
        ]

        # Explicit Row Heights to ensure no overlap
        row_h = [0.9*cm, 0.9*cm, 1.1*cm, 1.4*cm, 0.9*cm, 1.0*cm, 1.8*cm]
        
        t = Table(main_data, colWidths=[content_w*0.33, content_w*0.66], rowHeights=row_h)
        t.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1.2, colors.black),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold'),
            ('SPAN', (1, 6), (1, 6)), # Barcode span
            ('SPAN', (0, 6), (1, 6)), # Make barcode full width
        ]))

        all_elements.append(t)
        all_elements.append(PageBreak())

    doc.build(all_elements, onFirstPage=draw_border, onLaterPages=draw_border)
    return temp_pdf_path

# Streamlit UI
def main():
    st.set_page_config(page_title="Warehouse Label Gen")
    st.title("🏷️ Put Away Label Generator")
    st.markdown("Developed by Agilomatrix")
    
    uploaded_file = st.file_uploader("Upload Excel/CSV", type=['xlsx', 'csv'])
    
    if uploaded_file:
        df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
        st.write("Preview:", df.head(3))
        
        if st.button("Generate Labels", type="primary"):
            pdf_path = generate_sticker_labels(df)
            with open(pdf_path, "rb") as f:
                st.download_button("📥 Download PDF", f, file_name="warehouse_labels.pdf")
            os.unlink(pdf_path)

if __name__ == "__main__":
    main()
