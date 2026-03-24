import streamlit as st
import pandas as pd
import os
from reportlab.lib.pagesizes import landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph, PageBreak, Image
from reportlab.lib.units import cm, inch
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.lib.utils import ImageReader
from reportlab.graphics.barcode import code128
from reportlab.graphics.shapes import Drawing
from reportlab.graphics import renderPDF
from io import BytesIO
import re
import tempfile

# Auto-install required packages
try:
    from PIL import Image as PILImage
except ImportError:
    st.error("PIL not available. Please install: pip install pillow")
    st.stop()

# ── Sticker page dimensions ──────────────────────────────────────────────────
STICKER_WIDTH  = 10 * cm
STICKER_HEIGHT = 15 * cm
STICKER_PAGESIZE = (STICKER_WIDTH, STICKER_HEIGHT)

# ── Paragraph styles ──────────────────────────────────────────────────────────
header_label_style = ParagraphStyle(
    name='HeaderLabel',
    fontName='Helvetica-Bold',
    fontSize=10,
    alignment=TA_LEFT,
    leading=12,
)
header_value_style = ParagraphStyle(
    name='HeaderValue',
    fontName='Helvetica-Bold',
    fontSize=11,
    alignment=TA_LEFT,
    leading=12,
)
big_value_style = ParagraphStyle(
    name='BigValue',
    fontName='Helvetica-Bold',
    fontSize=13,
    alignment=TA_CENTER,
    leading=14,
)
desc_value_style = ParagraphStyle(
    name='DescValue',
    fontName='Helvetica',
    fontSize=10,
    alignment=TA_CENTER,
    leading=12,
)
loc_label_style = ParagraphStyle(
    name='LocLabel',
    fontName='Helvetica-Bold',
    fontSize=10,
    alignment=TA_CENTER,
    leading=12,
)


# ── Helpers ───────────────────────────────────────────────────────────────────

def generate_barcode(data_string, width_cm=9.0, height_cm=1.8):
    """Return a ReportLab Image of a Code-128 barcode."""
    try:
        barcode = code128.Code128(
            data_string,
            barWidth=0.5,
            barHeight=height_cm * cm,
            humanReadable=True,
            fontSize=8,
            fontName='Helvetica',
        )
        buf = BytesIO()
        barcode_width  = barcode.width
        barcode_height = barcode.height

        d = Drawing(barcode_width, barcode_height)
        barcode.drawOn(d._canvas if hasattr(d, '_canvas') else None, 0, 0)

        # Use renderPDF to a temp canvas then grab image via PIL approach
        # Simpler: draw directly via platypus flowable trick
        return barcode          # we'll use it as a Flowable directly
    except Exception as e:
        st.error(f"Barcode error: {e}")
        return None


def parse_location_string(location_str):
    """Split a location string into up to 4 parts."""
    parts = [''] * 4
    if not location_str or not isinstance(location_str, str):
        return parts
    matches = re.findall(r'([^_\s]+)', location_str.strip())
    for i, m in enumerate(matches[:4]):
        parts[i] = m
    return parts


def clean_date(date_val):
    """Return a clean date string (date part only)."""
    s = str(date_val) if date_val and str(date_val) != 'nan' else ''
    if not s:
        return ''
    return s.split(' ')[0] if ' ' in s else s


def row_label_value(label, value, label_w, value_w, row_h, value_style=None):
    """Return a 2-column Table row: bold label | value."""
    vs = value_style or header_value_style
    return Table(
        [[Paragraph(label, header_label_style), Paragraph(str(value), vs)]],
        colWidths=[label_w, value_w],
        rowHeights=[row_h],
    )


# ── Main generator ────────────────────────────────────────────────────────────

def generate_sticker_labels(df):

    def draw_border(canvas, doc):
        canvas.saveState()
        canvas.setStrokeColor(colors.Color(0, 0, 0, alpha=0.95))
        canvas.setLineWidth(1.8)
        canvas.rect(
            doc.leftMargin,
            doc.bottomMargin,
            STICKER_WIDTH  - doc.leftMargin  - doc.rightMargin,
            STICKER_HEIGHT - doc.topMargin   - doc.bottomMargin,
        )
        canvas.restoreState()

    # ── Normalise column names ────────────────────────────────────────────────
    df_copy = df.copy()
    df_copy.columns = [c.upper().strip() if isinstance(c, str) else c for c in df_copy.columns]
    cols = df_copy.columns.tolist()

    def find_col(*keywords, fallback=None):
        for kw_group in keywords:
            if isinstance(kw_group, str):
                kw_group = [kw_group]
            for col in cols:
                if all(k in col for k in kw_group):
                    return col
        return fallback

    grn_no_col       = find_col(['GRN', 'NO'], ['GRN', 'NUM'], 'GRN', fallback=cols[0])
    grn_date_col     = find_col(['GRN', 'DATE'], ['RECEIPT', 'DATE'], 'DATE', fallback=None)
    part_no_col      = find_col(['PART', 'NO'], ['PART', 'NUM'], 'PART', fallback=cols[0])
    desc_col         = find_col('DESC', 'NAME', fallback=cols[1] if len(cols) > 1 else cols[0])
    qty_col          = find_col('QTY', 'QUANTITY', fallback=None)
    store_loc_col    = find_col(['STORE', 'LOC'], 'STORELOCATION', 'LOCATION', 'LOC',
                                fallback=cols[2] if len(cols) > 2 else None)

    # ── Document setup ────────────────────────────────────────────────────────
    temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    temp_pdf_path = temp_pdf.name
    temp_pdf.close()

    MARGIN = 0.25 * cm
    doc = SimpleDocTemplate(
        temp_pdf_path,
        pagesize=STICKER_PAGESIZE,
        topMargin=MARGIN,
        bottomMargin=MARGIN,
        leftMargin=MARGIN,
        rightMargin=MARGIN,
    )

    content_width = STICKER_WIDTH - 2 * MARGIN
    label_w = content_width * 0.38   # left column (field name)
    value_w = content_width * 0.62   # right column (value)
    row_h   = 1.1 * cm

    all_elements = []
    progress_bar       = st.progress(0)
    status_placeholder = st.empty()
    total_rows = len(df_copy)

    for idx, (_, row) in enumerate(df_copy.iterrows()):
        progress = (idx + 1) / total_rows
        progress_bar.progress(progress)
        status_placeholder.text(f"Creating sticker {idx+1} of {total_rows} ({int(progress*100)}%)")

        # ── Extract values ────────────────────────────────────────────────────
        def get(col):
            if col and col in row and pd.notna(row[col]):
                return str(row[col])
            return ''

        grn_no       = get(grn_no_col)
        grn_date     = clean_date(get(grn_date_col)) if grn_date_col else ''
        part_no      = get(part_no_col)
        desc         = get(desc_col)
        qty          = get(qty_col) if qty_col else ''
        store_loc    = get(store_loc_col) if store_loc_col else ''
        loc_parts    = parse_location_string(store_loc)

        # Truncate description if very long
        desc_display = desc[:55] + '…' if len(desc) > 58 else desc

        # ── Build table rows ──────────────────────────────────────────────────
        table_data = [
            [Paragraph("GRN No.",     header_label_style), Paragraph(grn_no,       big_value_style)],
            [Paragraph("GRN Date",    header_label_style), Paragraph(grn_date,      header_value_style)],
            [Paragraph("Part No.",    header_label_style), Paragraph(part_no,       big_value_style)],
            [Paragraph("Description", header_label_style), Paragraph(desc_display,  desc_value_style)],
            [Paragraph("Quantity",    header_label_style), Paragraph(qty,           big_value_style)],
        ]

        row_heights = [row_h, row_h, row_h, 1.4 * cm, row_h]

        main_table = Table(
            table_data,
            colWidths=[label_w, value_w],
            rowHeights=row_heights,
        )
        main_table.setStyle(TableStyle([
            ('GRID',      (0, 0), (-1, -1), 1.2,  colors.black),
            ('VALIGN',    (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN',     (0, 0), (0,  -1), 'LEFT'),
            ('ALIGN',     (1, 0), (1,  -1), 'CENTER'),
            ('LEFTPADDING',  (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING',   (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING',(0, 0), (-1, -1), 2),
            ('BACKGROUND', (0, 0), (0, -1), colors.Color(0.93, 0.93, 0.93)),
        ]))

        # ── Store Location row ────────────────────────────────────────────────
        inner_col_w = value_w / 4
        inner_table = Table(
            [loc_parts],
            colWidths=[inner_col_w] * 4,
            rowHeights=[row_h],
        )
        inner_table.setStyle(TableStyle([
            ('GRID',     (0, 0), (-1, -1), 1.2, colors.black),
            ('ALIGN',    (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN',   (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
        ]))

        loc_table = Table(
            [[Paragraph("Store Location", loc_label_style), inner_table]],
            colWidths=[label_w, value_w],
            rowHeights=[row_h],
        )
        loc_table.setStyle(TableStyle([
            ('GRID',       (0, 0), (-1, -1), 1.2, colors.black),
            ('VALIGN',     (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN',      (0, 0), (0, 0),   'CENTER'),
            ('BACKGROUND', (0, 0), (0, 0),   colors.Color(0.93, 0.93, 0.93)),
            ('LEFTPADDING',  (0, 0), (0, 0), 4),
        ]))

        # ── Barcode ───────────────────────────────────────────────────────────
        barcode_data = grn_no if grn_no else (part_no if part_no else "NO-DATA")
        try:
            bc = code128.Code128(
                barcode_data,
                barWidth=1.05,
                barHeight=1.5 * cm,
                humanReadable=True,
                fontSize=8,
                fontName='Helvetica',
            )
        except Exception as e:
            bc = None
            st.warning(f"Barcode generation failed for row {idx+1}: {e}")

        # ── Assemble page elements ────────────────────────────────────────────
        elements = [
            Spacer(1, 0.1 * cm),
            main_table,
            loc_table,
            Spacer(1, 0.25 * cm),
        ]

        if bc:
            # Wrap barcode in a centred single-cell table
            bc_table = Table(
                [[bc]],
                colWidths=[content_width],
                rowHeights=[2.0 * cm],
            )
            bc_table.setStyle(TableStyle([
                ('ALIGN',  (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('BOX',    (0, 0), (-1, -1), 1.2, colors.black),
            ]))
            elements.append(bc_table)
        else:
            elements.append(
                Table(
                    [[Paragraph("[ BARCODE ]", big_value_style)]],
                    colWidths=[content_width],
                    rowHeights=[2.0 * cm],
                )
            )

        all_elements.extend(elements)
        if idx < total_rows - 1:
            all_elements.append(PageBreak())

    # ── Build PDF ─────────────────────────────────────────────────────────────
    try:
        doc.build(all_elements, onFirstPage=draw_border, onLaterPages=draw_border)
        status_placeholder.text("PDF generated successfully!")
        progress_bar.progress(1.0)
        return temp_pdf_path
    except Exception as e:
        st.error(f"Error building PDF: {e}")
        return None


# ── Streamlit UI ──────────────────────────────────────────────────────────────

def main():
    st.set_page_config(
        page_title="Put Away Zone Label Generator",
        page_icon="🏷️",
        layout="wide",
    )

    st.title("🏷️ Put Away Zone Label Generator")
    st.markdown(
        "<p style='font-size:18px; font-style:italic; margin-top:-10px;'>"
        "Designed and Developed by Agilomatrix</p>",
        unsafe_allow_html=True,
    )
    st.markdown("---")

    col1, col2 = st.columns([2, 1])

    with col1:
        st.header("📁 Upload File")
        uploaded_file = st.file_uploader(
            "Choose an Excel or CSV file",
            type=['xlsx', 'xls', 'csv'],
            help="Upload your data file containing GRN information",
        )

        if uploaded_file is not None:
            try:
                if uploaded_file.name.lower().endswith('.csv'):
                    df = pd.read_csv(uploaded_file)
                else:
                    df = pd.read_excel(uploaded_file)

                st.success(f"✅ File loaded! {len(df)} rows × {len(df.columns)} columns.")
                st.subheader("📊 Data Preview")
                st.write(f"**Columns:** {', '.join(df.columns.tolist())}")
                st.dataframe(df.head(), use_container_width=True)

                st.subheader("🎯 Generate Labels")
                if st.button("🚀 Generate Sticker Labels", type="primary", use_container_width=True):
                    with st.spinner("Generating sticker labels…"):
                        pdf_path = generate_sticker_labels(df)

                    if pdf_path:
                        st.success("🎉 Sticker labels generated successfully!")
                        with open(pdf_path, "rb") as f:
                            pdf_bytes = f.read()

                        filename = f"{uploaded_file.name.rsplit('.', 1)[0]}_sticker_labels.pdf"
                        st.markdown("""
                        <div style="border:2px solid #4CAF50;border-radius:10px;padding:20px;
                                    text-align:center;background:#f0f8ff;margin:10px 0;">
                            <h4 style="color:#4CAF50;margin-bottom:15px;">📄 Your PDF is Ready!</h4>
                            <p>Click the button below to download your sticker labels</p>
                        </div>""", unsafe_allow_html=True)

                        st.download_button(
                            label="📥 Download PDF File",
                            data=pdf_bytes,
                            file_name=filename,
                            mime="application/pdf",
                            type="primary",
                            use_container_width=True,
                        )
                        st.info(f"📊 File size: {len(pdf_bytes)/1024/1024:.2f} MB | Labels: {len(df)}")

                        try:
                            os.unlink(pdf_path)
                        except Exception:
                            pass
                    else:
                        st.error("❌ Failed to generate sticker labels.")

            except Exception as e:
                st.error(f"❌ Error reading file: {e}")

    with col2:
        st.header("ℹ️ Instructions")
        st.markdown("""
**How to use:**
1. Upload your Excel or CSV file
2. Preview the data
3. Click **Generate Sticker Labels**
4. Download the PDF

**Expected columns:**
- GRN No. / GRN Number
- GRN Date / Receipt Date
- Part No. / Part Number
- Description / Name
- Quantity / Qty
- Store Location

**Label layout (top → bottom):**
- GRN No.
- GRN Date
- Part No.
- Description
- Quantity
- Store Location
- **Barcode** (Code-128, GRN No.)
""")

        st.header("⚙️ Settings")
        st.markdown("""
**Fixed configuration:**
- Sticker size: 10 × 15 cm
- Barcode: Code-128 (GRN No.)
- Professional border & shading
- Auto column detection
""")


if __name__ == "__main__":
    main()
