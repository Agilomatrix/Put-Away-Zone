import streamlit as st
import pandas as pd
import os
import re
import tempfile

from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.graphics.barcode import code128
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER

try:
    from PIL import Image as PILImage
except ImportError:
    st.error("PIL not available. Please install: pip install pillow")
    st.stop()

# ═══════════════════════════════════════════════════════════════════════════════
# FIXED DIMENSIONS  —  everything derived from these 3 numbers
# ═══════════════════════════════════════════════════════════════════════════════
PAGE_W  = 10.0 * cm
PAGE_H  = 15.0 * cm

BOX_W   = 9.6  * cm   # content box width
BOX_H   = 7.5  * cm   # content box height  (ALL rows must fit inside)

BOX_X   = (PAGE_W - BOX_W) / 2   # 0.20 cm — left edge of box
BOX_Y   = PAGE_H - 0.20*cm - BOX_H  # top of page minus small gap minus box height

# ── Column split (same for EVERY row, including Store Location) ───────────────
LABEL_W = BOX_W * 0.33          # 3.456 cm
VALUE_W = BOX_W - LABEL_W       # 6.144 cm

# ── Row heights — verified to sum exactly to BOX_H ───────────────────────────
#   GRN No.        0.80 cm
#   GRN Date       0.80 cm
#   Part No.       0.80 cm
#   Description    0.95 cm
#   Quantity       0.80 cm
#   Store Loc      0.80 cm
#   Barcode        2.55 cm
#   ───────────────────────
#   Total          7.50 cm  ✓
R  = 0.80 * cm   # standard row height
RD = 0.95 * cm   # description row
RL = 0.80 * cm   # store location row
RB = 2.55 * cm   # barcode row
# Check: 5*R + RD + RL + RB = 4.00+0.95+0.80+2.55 = 8.30  ← wrong, recount
# Rows: GRN No, GRN Date, Part No = 3 standard rows
#       Description = 1 desc row
#       Quantity = 1 standard row
#       Store Loc = 1 loc row
#       Barcode = 1 barcode row
# = 4*R + RD + RL + RB
# = 4*0.80 + 0.95 + 0.80 + 2.55
# = 3.20 + 0.95 + 0.80 + 2.55 = 7.50 ✓
R  = 0.80 * cm
RD = 0.95 * cm
RL = 0.80 * cm
RB = 2.55 * cm
# Final verification: 4*0.80 + 0.95 + 0.80 + 2.55 = 3.20+0.95+0.80+2.55 = 7.50 ✓

ROWS = [R, R, R, RD, R, RL, RB]  # top to bottom order: GRN No, Date, Part, Desc, Qty, Loc, BC

# ── Font sizes ────────────────────────────────────────────────────────────────
F_LABEL     = 12   # field name (left col)
F_LARGE     = 13   # GRN No., Part No., Quantity value
F_MEDIUM    = 11   # GRN Date value
F_SMALL     = 9    # Description value
F_LOC_VAL   = 10   # location box values
F_BC_NUM    = 8    # barcode human-readable number

LW = 1.0   # line width for grid

# ── Helpers ───────────────────────────────────────────────────────────────────

def parse_location(loc_str):
    parts = ['', '', '', '']
    if not loc_str or not isinstance(loc_str, str):
        return parts
    matches = re.findall(r'([^_\s]+)', loc_str.strip())
    for i, m in enumerate(matches[:4]):
        parts[i] = m
    return parts

def clean_date(val):
    s = str(val) if val and str(val) != 'nan' else ''
    return s.split(' ')[0] if ' ' in s else s

def clean_num(val):
    """Strip trailing .0 from floats read as strings."""
    if val.endswith('.0') and val[:-2].isdigit():
        return val[:-2]
    return val

def draw_centered_text(c, text, x, y, w, h, font, size, bold=False):
    """Draw text centred both horizontally and vertically in a cell."""
    fname = 'Helvetica-Bold' if bold else 'Helvetica'
    c.setFont(fname, size)
    # Truncate if too wide
    while c.stringWidth(text, fname, size) > w - 6 and size > 6:
        size -= 0.5
        c.setFont(fname, size)
    text_w = c.stringWidth(text, fname, size)
    tx = x + (w - text_w) / 2
    ty = y + (h - size * 0.35) / 2   # approximate vertical centre
    c.drawString(tx, ty, text)

def draw_left_text(c, text, x, y, w, h, font_size, bold=True):
    """Draw text left-aligned, vertically centred in a cell."""
    fname = 'Helvetica-Bold' if bold else 'Helvetica'
    c.setFont(fname, font_size)
    ty = y + (h - font_size * 0.35) / 2
    c.drawString(x + 4, ty, text)

def draw_wrapped_centered(c, text, x, y, w, h, font_size):
    """Draw wrapped text centred in cell (for Description)."""
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_CENTER
    from reportlab.platypus import Paragraph
    style = ParagraphStyle('tmp', fontName='Helvetica', fontSize=font_size,
                           alignment=TA_CENTER, leading=font_size + 2)
    p = Paragraph(text, style)
    pw, ph = p.wrap(w - 8, h)
    # Centre vertically
    py = y + (h - ph) / 2
    p.drawOn(c, x + 4, py)

# ═══════════════════════════════════════════════════════════════════════════════
# DRAW ONE STICKER onto canvas  (draws at top of page every time)
# ═══════════════════════════════════════════════════════════════════════════════

def draw_sticker(c, grn_no, grn_date, part_no, desc, qty, loc_parts):
    """Draw the complete sticker content box on the current canvas page."""

    # ── Compute row Y positions ───────────────────────────────────────────────
    row_tops = []
    y = BOX_Y + BOX_H   
    for rh in ROWS:
        y -= rh
        row_tops.append(y)   

    # ── Draw outer border ─────────────────────────────────────────────────────
    c.setStrokeColor(colors.black)
    c.setLineWidth(1.5)
    c.rect(BOX_X, BOX_Y, BOX_W, BOX_H)

    # ── Draw each row (0 to 4) ────────────────────────────────────────────────
    row_defs = [
        (0, "GRN No.",     grn_no,   F_LARGE,  True),
        (1, "GRN Date",    grn_date, F_MEDIUM, True),
        (2, "Part No.",    part_no,  F_LARGE,  True),
        (3, "Description", desc,     F_SMALL,  False),
        (4, "Quantity",    qty,      F_LARGE,  True),
    ]

    c.setLineWidth(LW)
    for ri, label, value, vsize, vbold in row_defs:
        ry  = row_tops[ri]
        rh  = ROWS[ri]
        if ri > 0:
            c.line(BOX_X, ry + rh, BOX_X + BOX_W, ry + rh)
        c.line(BOX_X + LABEL_W, ry, BOX_X + LABEL_W, ry + rh)
        draw_left_text(c, label, BOX_X, ry, LABEL_W, rh, F_LABEL, bold=True)
        if ri == 3:
            draw_wrapped_centered(c, value, BOX_X + LABEL_W, ry, VALUE_W, rh, vsize)
        else:
            draw_centered_text(c, value, BOX_X + LABEL_W, ry, VALUE_W, rh, None, vsize, bold=vbold)

    # ── Store Location row (index 5) ──────────────────────────────────────────
    ri, ry, rh = 5, row_tops[5], ROWS[5]
    c.line(BOX_X, ry + rh, BOX_X + BOX_W, ry + rh)
    c.line(BOX_X + LABEL_W, ry, BOX_X + LABEL_W, ry + rh)
    draw_left_text(c, "Store Location", BOX_X, ry, LABEL_W, rh, F_LABEL, bold=True)
    loc_box_w = VALUE_W / 4
    for i, part in enumerate(loc_parts):
        lx = BOX_X + LABEL_W + i * loc_box_w
        if i > 0: c.line(lx, ry, lx, ry + rh)
        draw_centered_text(c, part, lx, ry, loc_box_w, rh, None, F_LOC_VAL, bold=True)

    # ── Barcode row (index 6) ── SCAN ALL FIELDS ──────────────────────────────
    ri, ry, rh = 6, row_tops[6], ROWS[6]
    c.line(BOX_X, ry + rh, BOX_X + BOX_W, ry + rh)

    # Create the Combined Data String for the Barcode
    # Format: GRN|Date|Part|Desc|Qty|Loc
    loc_combined = "-".join([p for p in loc_parts if p.strip()])
    # We truncate Description to 15 chars in the barcode to keep the barcode width scan-able
    bc_data = f"{grn_no}|{grn_date}|{part_no}|{desc[:15]}|{qty}|{loc_combined}"
    
    # Filter out characters that Code128 might struggle with in basic scanners
    bc_data = re.sub(r'[^a-zA-Z0-9|.-]', ' ', bc_data)

    try:
        # Dynamic bar width calculation to ensure it fits 9.6cm
        # Code 128 uses roughly 11 modules per character + overhead
        char_len = len(bc_data)
        # Calculate bar width: available width (9.6cm - padding) / modules
        # A character in Code128 is 11 modules wide. 
        # Total modules approx = (chars * 11) + 35
        estimated_modules = (char_len * 11) + 35
        avail_points = (BOX_W - 0.6*cm) # 9.6cm minus some padding
        calculated_bar_w = avail_points / estimated_modules
        
        # Ensure bar_w is within scannable limits (usually 0.4 to 1.2)
        final_bar_w = max(0.42, min(calculated_bar_w, 1.2))

        bc = code128.Code128(
            bc_data,
            barWidth=final_bar_w,
            barHeight=rh * 0.60,
            humanReadable=True, # Shows the text below the bars
            fontSize=F_BC_NUM,
            fontName='Helvetica',
        )
        
        # Center the barcode
        bc_x = BOX_X + (BOX_W - bc.width) / 2
        bc_y = ry + (rh - bc.height) / 2
        bc.drawOn(c, bc_x, bc_y)
        
    except Exception as e:
        c.setFont('Helvetica', 8)
        c.drawCentredString(BOX_X + BOX_W / 2, ry + rh / 2, f"Error: {str(e)[:40]}")


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN GENERATOR
# ═══════════════════════════════════════════════════════════════════════════════

def generate_sticker_labels(df):
    # ── Normalise columns ─────────────────────────────────────────────────────
    df_copy = df.copy()
    df_copy.columns = [col.upper().strip() if isinstance(col, str) else col
                       for col in df_copy.columns]
    cols = df_copy.columns.tolist()

    def find_col(*kw_groups, fallback=None):
        for kwg in kw_groups:
            kwg = [kwg] if isinstance(kwg, str) else list(kwg)
            for col in cols:
                if all(k in col for k in kwg):
                    return col
        return fallback

    grn_no_col    = find_col(['GRN','NO'], ['GRN','NUM'], ['GRN','#'],
                              'GRNNO', 'GRN_NO', 'GRN', fallback=cols[0])
    grn_date_col  = find_col(['GRN','DATE'], ['RECEIPT','DATE'], ['GRN','DT'], 'DATE',
                              fallback=None)
    part_no_col   = find_col(['PART','NO'], ['PART','NUM'], ['PART','#'],
                              'PARTNO', 'PART_NO', 'PART', fallback=cols[0])
    desc_col      = find_col('DESC', 'DESCRIPTION', 'NAME',
                              fallback=cols[1] if len(cols) > 1 else cols[0])
    qty_col       = find_col('QTY', 'QUANTITY', fallback=None)
    store_loc_col = find_col(['STORE','LOC'], 'STORELOCATION', 'STORE_LOCATION',
                              'LOCATION', 'LOC',
                              fallback=cols[2] if len(cols) > 2 else None)

    # ── Create PDF ────────────────────────────────────────────────────────────
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    tmp_path = tmp.name
    tmp.close()

    c = canvas.Canvas(tmp_path, pagesize=(PAGE_W, PAGE_H))

    progress_bar = st.progress(0)
    status_ph    = st.empty()
    total_rows   = len(df_copy)

    for idx, (_, row) in enumerate(df_copy.iterrows()):
        progress_bar.progress((idx + 1) / total_rows)
        status_ph.text(f"Creating sticker {idx+1} of {total_rows} ({int((idx+1)/total_rows*100)}%)")

        def get(col):
            if col and col in row and pd.notna(row[col]):
                v = str(row[col]).strip()
                return clean_num(v)
            return ''

        grn_no    = get(grn_no_col)
        grn_date  = clean_date(get(grn_date_col)) if grn_date_col else ''
        part_no   = get(part_no_col)
        desc      = get(desc_col)
        qty       = get(qty_col) if qty_col else ''
        store_loc = get(store_loc_col) if store_loc_col else ''
        loc_parts = parse_location(store_loc)

        draw_sticker(c, grn_no, grn_date, part_no, desc, qty, loc_parts)
        c.showPage()   # new page for next sticker

    c.save()
    status_ph.text("PDF generated successfully!")
    progress_bar.progress(1.0)
    return tmp_path


# ═══════════════════════════════════════════════════════════════════════════════
# STREAMLIT UI
# ═══════════════════════════════════════════════════════════════════════════════

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
7. Barcode (Code-128, GRN No.)
""")

        st.header("⚙️ Layout")
        st.markdown("""
**Fixed configuration:**
- Sticker page  : 10 × 15 cm
- Content box   : 9.6 × 7.5 cm
- All white background
- Pure canvas drawing — no overflow possible
- Barcode: Code-128, encodes GRN No.
""")


if __name__ == "__main__":
    main()
