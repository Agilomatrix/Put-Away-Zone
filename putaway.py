import streamlit as st
import pandas as pd
import os
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph, PageBreak
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.graphics.barcode import code128
import re
import tempfile

try:
    from PIL import Image as PILImage
except ImportError:
    st.error("PIL not available. Please install: pip install pillow")
    st.stop()

# ── Page dimensions ───────────────────────────────────────────────────────────
STICKER_WIDTH   = 10 * cm
STICKER_HEIGHT  = 15 * cm
STICKER_PAGESIZE = (STICKER_WIDTH, STICKER_HEIGHT)

# ── Content box: exactly 10 cm wide × 7.5 cm tall ────────────────────────────
CONTENT_W = 10 * cm
CONTENT_H =  7.5 * cm

# Tiny page margins so the content box butts up to the edges
SIDE_MARGIN = 0.0 * cm
TOP_MARGIN  = 0.0 * cm   # distance from top of page to top of box

# ── Inner usable width ────────────────────────────────────────────────────────
INNER_W  = CONTENT_W          # full 10 cm
LABEL_W  = INNER_W * 0.36     # 3.6 cm — field-name column
VALUE_W  = INNER_W - LABEL_W  # 6.4 cm — value column

# ── Row heights: must sum to exactly CONTENT_H = 7.5 cm ──────────────────────
# 5 standard rows × 0.78 cm = 3.90 cm
# Description row              = 0.90 cm
# Store Location row           = 0.82 cm
# Barcode row                  = 1.88 cm
# Total = 3.90 + 0.90 + 0.82 + 1.88 = 7.50 cm ✓
ROW_H      = 0.78 * cm
DESC_ROW_H = 0.90 * cm
LOC_ROW_H  = 0.82 * cm
BC_ROW_H   = 1.88 * cm

# ── Paragraph styles ──────────────────────────────────────────────────────────
lbl_style = ParagraphStyle(
    'Lbl', fontName='Helvetica-Bold', fontSize=8.5,
    alignment=TA_LEFT, leading=10, leftIndent=3,
)
val_large = ParagraphStyle(
    'ValLarge', fontName='Helvetica-Bold', fontSize=11,
    alignment=TA_CENTER, leading=12,
)
val_bold = ParagraphStyle(
    'ValBold', fontName='Helvetica-Bold', fontSize=10,
    alignment=TA_CENTER, leading=11,
)
val_normal = ParagraphStyle(
    'ValNorm', fontName='Helvetica', fontSize=9,
    alignment=TA_CENTER, leading=10,
)
loc_lbl_style = ParagraphStyle(
    'LocLbl', fontName='Helvetica-Bold', fontSize=8.5,
    alignment=TA_CENTER, leading=10,
)

# ── Helpers ───────────────────────────────────────────────────────────────────

def parse_location(loc_str):
    parts = [''] * 4
    if not loc_str or not isinstance(loc_str, str):
        return parts
    matches = re.findall(r'([^_\s]+)', loc_str.strip())
    for i, m in enumerate(matches[:4]):
        parts[i] = m
    return parts


def clean_date(val):
    s = str(val) if val and str(val) != 'nan' else ''
    return s.split(' ')[0] if ' ' in s else s


# ── Main generator ────────────────────────────────────────────────────────────

def generate_sticker_labels(df):

    def draw_border(canvas, doc):
        """Draw border exactly around the 10 × 7.5 cm content box."""
        canvas.saveState()
        canvas.setStrokeColor(colors.black)
        canvas.setLineWidth(1.5)
        # box starts at bottom-left = (0, STICKER_HEIGHT - CONTENT_H)
        box_x = 0
        box_y = STICKER_HEIGHT - CONTENT_H
        canvas.rect(box_x, box_y, CONTENT_W, CONTENT_H)
        canvas.restoreState()

    # ── Normalise column names ─────────────────────────────────────────────────
    df_copy = df.copy()
    df_copy.columns = [c.upper().strip() if isinstance(c, str) else c for c in df_copy.columns]
    cols = df_copy.columns.tolist()

    def find_col(*kw_groups, fallback=None):
        for kwg in kw_groups:
            kwg = [kwg] if isinstance(kwg, str) else kwg
            for col in cols:
                if all(k in col for k in kwg):
                    return col
        return fallback

    grn_no_col    = find_col(['GRN', 'NO'], ['GRN', 'NUM'], 'GRN',        fallback=cols[0])
    grn_date_col  = find_col(['GRN', 'DATE'], ['RECEIPT', 'DATE'], 'DATE', fallback=None)
    part_no_col   = find_col(['PART', 'NO'], ['PART', 'NUM'], 'PART',      fallback=cols[0])
    desc_col      = find_col('DESC', 'NAME',                               fallback=cols[1] if len(cols) > 1 else cols[0])
    qty_col       = find_col('QTY', 'QUANTITY',                            fallback=None)
    store_loc_col = find_col(['STORE', 'LOC'], 'STORELOCATION', 'LOCATION', 'LOC',
                              fallback=cols[2] if len(cols) > 2 else None)

    # ── Document: place flow area exactly inside the content box ─────────────
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    tmp_path = tmp.name
    tmp.close()

    doc = SimpleDocTemplate(
        tmp_path,
        pagesize=STICKER_PAGESIZE,
        topMargin=TOP_MARGIN,
        bottomMargin=STICKER_HEIGHT - TOP_MARGIN - CONTENT_H,
        leftMargin=SIDE_MARGIN,
        rightMargin=SIDE_MARGIN,
    )

    all_elements = []
    progress_bar = st.progress(0)
    status_ph    = st.empty()
    total_rows   = len(df_copy)

    for idx, (_, row) in enumerate(df_copy.iterrows()):
        progress_bar.progress((idx + 1) / total_rows)
        status_ph.text(f"Creating sticker {idx+1} of {total_rows} ({int((idx+1)/total_rows*100)}%)")

        def get(col):
            if col and col in row and pd.notna(row[col]):
                return str(row[col])
            return ''

        grn_no    = get(grn_no_col)
        grn_date  = clean_date(get(grn_date_col)) if grn_date_col else ''
        part_no   = get(part_no_col)
        desc      = get(desc_col)
        qty       = get(qty_col) if qty_col else ''
        store_loc = get(store_loc_col) if store_loc_col else ''
        loc_parts = parse_location(store_loc)

        desc_display = desc[:52] + '…' if len(desc) > 55 else desc

        # ── Main 5-row table ──────────────────────────────────────────────────
        table_data = [
            [Paragraph("GRN No.",     lbl_style), Paragraph(grn_no,       val_large)],
            [Paragraph("GRN Date",    lbl_style), Paragraph(grn_date,      val_bold)],
            [Paragraph("Part No.",    lbl_style), Paragraph(part_no,       val_large)],
            [Paragraph("Description", lbl_style), Paragraph(desc_display,  val_normal)],
            [Paragraph("Quantity",    lbl_style), Paragraph(qty,           val_large)],
        ]
        main_table = Table(
            table_data,
            colWidths=[LABEL_W, VALUE_W],
            rowHeights=[ROW_H, ROW_H, ROW_H, DESC_ROW_H, ROW_H],
        )
        main_table.setStyle(TableStyle([
            ('GRID',          (0, 0), (-1, -1), 1.0,  colors.black),
            ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
            ('BACKGROUND',    (0, 0), (0,  -1), colors.Color(0.91, 0.91, 0.91)),
            ('LEFTPADDING',   (0, 0), (-1, -1), 3),
            ('RIGHTPADDING',  (0, 0), (-1, -1), 3),
            ('TOPPADDING',    (0, 0), (-1, -1), 1),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
        ]))

        # ── Store Location row ────────────────────────────────────────────────
        loc_inner_w = VALUE_W / 4
        inner_table = Table(
            [loc_parts],
            colWidths=[loc_inner_w] * 4,
            rowHeights=[LOC_ROW_H],
        )
        inner_table.setStyle(TableStyle([
            ('GRID',    (0, 0), (-1, -1), 1.0, colors.black),
            ('ALIGN',   (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN',  (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME',(0, 0), (-1, -1), 'Helvetica-Bold'),
            ('FONTSIZE',(0, 0), (-1, -1), 9),
            ('TOPPADDING',    (0, 0), (-1, -1), 1),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
        ]))

        loc_table = Table(
            [[Paragraph("Store\nLocation", loc_lbl_style), inner_table]],
            colWidths=[LABEL_W, VALUE_W],
            rowHeights=[LOC_ROW_H],
        )
        loc_table.setStyle(TableStyle([
            ('GRID',          (0, 0), (-1, -1), 1.0, colors.black),
            ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN',         (0, 0), (0,  0),  'CENTER'),
            ('BACKGROUND',    (0, 0), (0,  0),  colors.Color(0.91, 0.91, 0.91)),
            ('LEFTPADDING',   (0, 0), (-1, -1), 2),
            ('RIGHTPADDING',  (0, 0), (-1, -1), 2),
            ('TOPPADDING',    (0, 0), (-1, -1), 1),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
        ]))

        # ── Barcode row ───────────────────────────────────────────────────────
        bc_data = grn_no if grn_no else (part_no if part_no else "NO-DATA")

        # Auto-scale barWidth so barcode fits within INNER_W
        char_count = max(len(bc_data), 1)
        target_w   = INNER_W - 0.4 * cm     # leave tiny padding
        bar_w      = target_w / (char_count * 11 + 35 + 20)
        bar_w      = max(0.55, min(bar_w, 1.5))   # clamp

        try:
            bc = code128.Code128(
                bc_data,
                barWidth=bar_w,
                barHeight=BC_ROW_H * 0.58,
                humanReadable=True,
                fontSize=7,
                fontName='Helvetica',
            )
            bc_cell = bc
        except Exception as e:
            bc_cell = Paragraph(f"[BARCODE: {bc_data}]", val_normal)
            st.warning(f"Barcode error row {idx+1}: {e}")

        bc_table = Table(
            [[bc_cell]],
            colWidths=[INNER_W],
            rowHeights=[BC_ROW_H],
        )
        bc_table.setStyle(TableStyle([
            ('BOX',           (0, 0), (-1, -1), 1.0, colors.black),
            ('ALIGN',         (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING',    (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ]))

        # ── Assemble ──────────────────────────────────────────────────────────
        all_elements.extend([main_table, loc_table, bc_table])
        if idx < total_rows - 1:
            all_elements.append(PageBreak())

    # ── Build PDF ─────────────────────────────────────────────────────────────
    try:
        doc.build(all_elements, onFirstPage=draw_border, onLaterPages=draw_border)
        status_ph.text("PDF generated successfully!")
        progress_bar.progress(1.0)
        return tmp_path
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
        "<p style='font-size:18px;font-style:italic;margin-top:-10px;'>"
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
                df = (pd.read_csv(uploaded_file)
                      if uploaded_file.name.lower().endswith('.csv')
                      else pd.read_excel(uploaded_file))

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
2. Review the data preview
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
1. GRN No.
2. GRN Date
3. Part No.
4. Description
5. Quantity
6. Store Location (4-box grid)
7. Barcode (Code-128)
""")

        st.header("⚙️ Layout")
        st.markdown("""
**Fixed configuration:**
- Sticker page  : 10 × 15 cm
- Content box   : 10 × 7.5 cm (exact)
- Barcode       : Code-128 (GRN No.)
- Auto column detection
- Shaded label column
""")


if __name__ == "__main__":
    main()
