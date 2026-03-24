import streamlit as st
import pandas as pd
import os
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, PageBreak, Spacer
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.graphics.barcode import code128
from reportlab.graphics.shapes import Drawing
import re
import tempfile

# STICKER DIMENSIONS UPDATED
STICKER_WIDTH = 10 * cm
STICKER_HEIGHT = 7.5 * cm
STICKER_PAGESIZE = (STICKER_WIDTH, STICKER_HEIGHT)

# STYLES (Reduced font size to fit smaller sticker)
bold_style = ParagraphStyle(name='Bold', fontName='Helvetica-Bold', fontSize=9, alignment=TA_CENTER, leading=10)
desc_style = ParagraphStyle(name='Desc', fontName='Helvetica', fontSize=8, alignment=TA_CENTER, leading=9)
val_style = ParagraphStyle(name='Value', fontName='Helvetica-Bold', fontSize=10, alignment=TA_CENTER, leading=10)

def generate_barcode(data_string, width):
    """Generate a Code128 barcode scaled to fit"""
    try:
        bc = code128.Code128(data_string, barHeight=1.0*cm, barWidth=0.5)
        d = Drawing(width, 1.2*cm)
        d.add(bc)
        return d
    except:
        return Paragraph("Error", bold_style)

def parse_location_string(location_str):
    location_parts = [''] * 4
    if not location_str or not isinstance(location_str, str): return location_parts
    matches = re.findall(r'([^_\s]+)', str(location_str))
    for i, match in enumerate(matches[:4]): location_parts[i] = match
    return location_parts

def generate_sticker_labels(df):
    
    def draw_border(canvas, doc):
        canvas.saveState()
        canvas.setLineWidth(1.5)
        # Draw border leaving tiny margin from edges
        canvas.rect(0.2*cm, 0.2*cm, 9.6*cm, 7.1*cm)
        canvas.restoreState()

    temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    temp_pdf_path = temp_pdf.name
    
    # Doc margins set to minimum
    doc = SimpleDocTemplate(temp_pdf_path, pagesize=STICKER_PAGESIZE,
                          topMargin=0.4*cm, bottomMargin=0.4*cm, leftMargin=0.3*cm, rightMargin=0.3*cm)

    all_elements = []
    content_w = 9.4*cm # Adjusted to fit inside

    for index, row in df.iterrows():
        
        # Column auto-detection
        grn_no = str(row.get(next((c for c in df.columns if 'GRN' in c.upper() and 'DATE' not in c.upper()), df.columns[0]), ""))
        grn_date = str(row.get(next((c for c in df.columns if 'DATE' in c.upper()), ""), ""))
        part_no = str(row.get(next((c for c in df.columns if 'PART' in c.upper()), ""), ""))
        desc = str(row.get(next((c for c in df.columns if 'DESC' in c.upper() or 'NAME' in c.upper()), ""), ""))
        qty = str(row.get(next((c for c in df.columns if 'QTY' in c.upper() or 'QUANTITY' in c.upper()), ""), ""))
        loc_raw = str(row.get(next((c for c in df.columns if 'LOC' in c.upper()), ""), ""))
        
        loc_parts = parse_location_string(loc_raw)
        barcode_data = f"{grn_no}|{part_no}|{qty}|{loc_raw}"

        # Store Location table
        loc_table = Table([loc_parts], colWidths=[(content_w*0.65)/4]*4, rowHeights=[0.7*cm])
        loc_table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.8, colors.black),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTSIZE', (0,0), (-1,-1), 8),
        ]))

        # Main Data Rows
        main_data = [
            ["GRN No", Paragraph(grn_no, val_style)],
            ["GRN Date", Paragraph(grn_date.split(' ')[0], val_style)],
            ["Part No", Paragraph(part_no, val_style)],
            ["Desc", Paragraph(desc[:45], desc_style)],
            ["Qty", Paragraph(qty, val_style)],
            ["Location", loc_table],
            ["Barcode", generate_barcode(barcode_data, content_w*0.65)]
        ]

        # Compact Row heights to fit 7.5cm height
        row_h = [0.7*cm, 0.7*cm, 0.7*cm, 1.1*cm, 0.7*cm, 0.8*cm, 1.3*cm]

        t = Table(main_data, colWidths=[content_w*0.35, content_w*0.65], rowHeights=row_h)
        t.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.8, colors.black),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (0,-1), 8),
            # Spanning Barcode across the bottom
            ('SPAN', (1, 6), (1, 6)),
            ('ALIGN', (1,6), (1,6), 'CENTER')
        ]))

        all_elements.append(t)
        if index < len(df) - 1:
            all_elements.append(PageBreak())

    doc.build(all_elements, onFirstPage=draw_border, onLaterPages=draw_border)
    return temp_pdf_path

def main():
    st.set_page_config(page_title="Sticker Gen", page_icon="🏷️")
    st.title("🏷️ Small Sticker Generator")
    
    uploaded_file = st.file_uploader("Upload Data", type=['xlsx', 'csv'])
    if uploaded_file:
        df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
        if st.button("Generate Sticker", type="primary"):
            pdf_path = generate_sticker_labels(df)
            with open(pdf_path, "rb") as f:
                st.download_button("Download Labels", f, file_name="stickers.pdf")
            os.unlink(pdf_path)

if __name__ == "__main__":
    main()
