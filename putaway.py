import streamlit as st
import pandas as pd
import os
from reportlab.lib.pagesizes import landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph, PageBreak, Image
from reportlab.lib.units import cm, inch
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.graphics.barcode import code128
from reportlab.graphics.shapes import Drawing
from io import BytesIO
import re
import tempfile

# Auto-install check for PIL (needed for some ReportLab functions)
try:
    from PIL import Image as PILImage
except ImportError:
    st.error("PIL not available. Please install: pip install pillow")
    st.stop()

# Define sticker dimensions
STICKER_WIDTH = 10 * cm
STICKER_HEIGHT = 15 * cm
STICKER_PAGESIZE = (STICKER_WIDTH, STICKER_HEIGHT)

# Define content box dimensions
CONTENT_BOX_WIDTH = 10 * cm
CONTENT_BOX_HEIGHT = 7.2 * cm

# Fixed layout settings
DATE_WIDTH_RATIO = 0.40  # Reduced to give barcode more room
DATE_HEIGHT = 1.2  # cm
BARCODE_HEIGHT = 2.0  # cm

# Define paragraph styles
bold_style = ParagraphStyle(name='Bold', fontName='Helvetica-Bold', fontSize=16, alignment=TA_CENTER, leading=14)
desc_style = ParagraphStyle(name='Description', fontName='Helvetica', fontSize=11, alignment=TA_CENTER, leading=12)
qty_style = ParagraphStyle(name='Quantity', fontName='Helvetica', fontSize=11, alignment=TA_CENTER, leading=12)

def generate_barcode(data_string):
    """Generate a Code128 barcode from the given data string"""
    try:
        # Clean data string: Barcodes can't handle newlines well, use pipes/spaces
        clean_data = data_string.replace('\n', ' | ')
        
        # Create the barcode
        # barWidth controls the thickness of lines. 
        # Since we want to "scan everything", we keep it thin to fit the width.
        bc = code128.Code128(clean_data, barHeight=1.2*cm, barWidth=0.8)
        
        # Wrap the barcode in a Drawing object so it can be placed in a Table
        d = Drawing(bc.width, bc.height)
        d.add(bc)
        return d
    except Exception as e:
        st.error(f"Error generating Barcode: {e}")
        return None

def parse_location_string(location_str):
    """Parse a location string into components for table display - 4 boxes"""
    location_parts = [''] * 4
    if not location_str or not isinstance(location_str, str):
        return location_parts

    location_str = location_str.strip()
    pattern = r'([^_\s]+)'
    matches = re.findall(pattern, location_str)

    for i, match in enumerate(matches[:4]):
        location_parts[i] = match

    return location_parts

def generate_sticker_labels(df):
    """Generate sticker labels with Barcode from DataFrame"""
    
    def draw_border(canvas, doc):
        canvas.saveState()
        x_offset = (STICKER_WIDTH - CONTENT_BOX_WIDTH) / 2
        y_offset = STICKER_HEIGHT - CONTENT_BOX_HEIGHT - 0.2*cm
        canvas.setStrokeColor(colors.Color(0, 0, 0, alpha=0.95))
        canvas.setLineWidth(1.8)
        canvas.rect(
            x_offset + doc.leftMargin,
            y_offset,
            CONTENT_BOX_WIDTH - 0.2*cm,
            CONTENT_BOX_HEIGHT
        )
        canvas.restoreState()

    # Identify columns (case-insensitive)
    df_copy = df.copy()
    df_copy.columns = [col.upper() if isinstance(col, str) else col for col in df_copy.columns]
    cols = df_copy.columns.tolist()

    grn_col = next((col for col in cols if 'GRN' in col), cols[0])
    part_no_col = next((col for col in cols if 'PART' in col), cols[0])
    desc_col = next((col for col in cols if 'DESC' in col or 'NAME' in col), cols[1] if len(cols) > 1 else part_no_col)
    store_location_col = next((col for col in cols if 'LOC' in col or 'STORE' in col), cols[2] if len(cols) > 2 else desc_col)
    receipt_date_col = next((col for col in cols if 'DATE' in col), None)

    temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    temp_pdf_path = temp_pdf.name
    temp_pdf.close()

    doc = SimpleDocTemplate(temp_pdf_path, pagesize=STICKER_PAGESIZE,
                          topMargin=0.2*cm,
                          bottomMargin=(STICKER_HEIGHT - CONTENT_BOX_HEIGHT - 0.2*cm),
                          leftMargin=0.1*cm, rightMargin=0.1*cm)

    content_width = CONTENT_BOX_WIDTH - 0.2*cm
    all_elements = []

    progress_bar = st.progress(0)
    status_placeholder = st.empty()
    
    total_rows = len(df_copy)
    for index, row in df_copy.iterrows():
        progress = (index + 1) / total_rows
        progress_bar.progress(progress)
        status_placeholder.text(f"Creating sticker {index+1} of {total_rows}")
        
        elements = []

        grn_no = str(row[grn_col]) if grn_col in row and pd.notna(row[grn_col]) else ""
        part_no = str(row[part_no_col])
        desc = str(row[desc_col])
        store_location = str(row[store_location_col]) if store_location_col in row else ""
        receipt_date = str(row[receipt_date_col]) if receipt_date_col and pd.notna(row[receipt_date_col]) else ""
        
        location_parts = parse_location_string(store_location)

        # Main Table Layout
        grn_row_height = 0.9*cm
        header_row_height = 0.9*cm
        desc_row_height = 1.4*cm
        location_row_height = 0.8*cm

        # Barcode Data: Combine all info to "scan everything"
        barcode_data = f"{grn_no}|{part_no}|{store_location}|{receipt_date}"
        barcode_graphic = generate_barcode(barcode_data)

        main_table_data = [
            ["GRN No", Paragraph(f"{grn_no}", bold_style)],
            ["Part No", Paragraph(f"{part_no}", bold_style)],
            ["Description", Paragraph(desc[:47] + "..." if len(desc) > 50 else desc, desc_style)]
        ]

        main_table = Table(main_table_data,
                         colWidths=[content_width/3, content_width*2/3],
                         rowHeights=[grn_row_height, header_row_height, desc_row_height])

        main_table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1.2, colors.Color(0, 0, 0, alpha=0.95)),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (0, -1), 11),
        ]))
        elements.append(main_table)

        # Store Location Section (UNCHANGED as requested)
        store_location_label = Paragraph("Store Location", ParagraphStyle(name='SL', fontName='Helvetica-Bold', fontSize=11, alignment=TA_CENTER))
        inner_table_width = content_width * 2 / 3
        inner_col_widths = [inner_table_width / 4] * 4
        store_location_inner_table = Table([location_parts], colWidths=inner_col_widths, rowHeights=[location_row_height])
        store_location_inner_table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1.2, colors.Color(0, 0, 0, alpha=0.95)),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
        ]))

        store_location_table = Table([[store_location_label, store_location_inner_table]], 
                                    colWidths=[content_width/3, inner_table_width], 
                                    rowHeights=[location_row_height])
        store_location_table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1.2, colors.Color(0, 0, 0, alpha=0.95)),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        elements.append(store_location_table)

        # Bottom Section: Date and Barcode
        date_width = content_width * DATE_WIDTH_RATIO
        barcode_width = content_width - date_width

        date_table = Table(
            [["Date:", Paragraph(str(receipt_date).split(' ')[0], qty_style)]],
            colWidths=[date_width*0.4, date_width*0.6],
            rowHeights=[BARCODE_HEIGHT]
        )
        date_table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1.2, colors.Color(0, 0, 0, alpha=0.95)),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (0, 0), 'Helvetica-Bold'),
        ]))

        barcode_container = Table([[barcode_graphic]], colWidths=[barcode_width], rowHeights=[BARCODE_HEIGHT])
        barcode_container.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1.2, colors.Color(0, 0, 0, alpha=0.95)),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0,0), (-1,-1), 2),
            ('RIGHTPADDING', (0,0), (-1,-1), 2),
        ]))

        bottom_table = Table([[date_table, barcode_container]], colWidths=[date_width, barcode_width])
        bottom_table.setStyle(TableStyle([
            ('LEFTPADDING', (0,0), (-1,-1), 0),
            ('RIGHTPADDING', (0,0), (-1,-1), 0),
            ('BOTTOMPADDING', (0,0), (-1,-1), 0),
            ('TOPPADDING', (0,0), (-1,-1), 0),
        ]))

        elements.append(Spacer(1, 0.2*cm))
        elements.append(bottom_table)
        all_elements.extend(elements)

        if index < len(df_copy) - 1:
            all_elements.append(PageBreak())

    try:
        doc.build(all_elements, onFirstPage=draw_border, onLaterPages=draw_border)
        return temp_pdf_path
    except Exception as e:
        st.error(f"Error building PDF: {e}")
        return None

def main():
    st.set_page_config(page_title="Put Away Label Gen", page_icon="🏷️", layout="wide")
    st.title("🏷️ Put Away Zone Label Generator")
    st.markdown("Designed and Developed by Agilomatrix")
    st.markdown("---")
    
    uploaded_file = st.file_uploader("Upload Excel/CSV", type=['xlsx', 'xls', 'csv'])
    
    if uploaded_file:
        try:
            df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
            st.success(f"File loaded: {len(df)} labels ready.")
            
            if st.button("🚀 Generate Barcode Labels", type="primary"):
                pdf_path = generate_sticker_labels(df)
                if pdf_path:
                    with open(pdf_path, "rb") as f:
                        st.download_button("📥 Download PDF", f, file_name="labels.pdf", mime="application/pdf")
                    os.unlink(pdf_path)
        except Exception as e:
            st.error(f"Error: {e}")

if __name__ == "__main__":
    main()
