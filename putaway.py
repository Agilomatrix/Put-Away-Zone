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
from io import BytesIO

try:
    import qrcode
    HAS_QR = True
except ImportError:
    HAS_QR = False

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
LABEL_W = BOX_W * 0.36          # 3.456 cm
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
F_LABEL     = 9    # field name (left col)
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

def draw_sticker(c, grn_no, grn_date, part_no, desc_full, qty, loc_parts):
    """Draw the complete sticker content box on the current canvas page."""

    # Truncated description for display only — QR uses full version
    desc_display = desc_full[:52] + '…' if len(desc_full) > 55 else desc_full

    # ── Compute row Y positions (top → bottom) ────────────────────────────────
    # BOX_Y is the bottom of the box; top of box = BOX_Y + BOX_H
    row_tops = []
    y = BOX_Y + BOX_H   # start at top of box
    for rh in ROWS:
        y -= rh
        row_tops.append(y)   # y = bottom of this row

    # row_tops[0] = bottom y of row 0 (GRN No.)
    # row_tops[6] = bottom y of row 6 (Barcode)

    # ── Draw outer border ─────────────────────────────────────────────────────
    c.setStrokeColor(colors.black)
    c.setLineWidth(1.5)
    c.rect(BOX_X, BOX_Y, BOX_W, BOX_H)

    # ── Draw each row ─────────────────────────────────────────────────────────
    row_defs = [
        (0, "GRN No.",     grn_no,        F_LARGE,  True),
        (1, "GRN Date",    grn_date,      F_MEDIUM, True),
        (2, "Part No.",    part_no,       F_LARGE,  True),
        (3, "Description", desc_display,  F_SMALL,  False),
        (4, "Quantity",    qty,           F_LARGE,  True),
    ]

    c.setLineWidth(LW)

    for ri, label, value, vsize, vbold in row_defs:
        ry  = row_tops[ri]          # bottom of row
        rh  = ROWS[ri]

        # Horizontal line at top of this row (skip for row 0 — that's the box top)
        if ri > 0:
            c.line(BOX_X, ry + rh, BOX_X + BOX_W, ry + rh)

        # Vertical divider between label and value
        c.line(BOX_X + LABEL_W, ry, BOX_X + LABEL_W, ry + rh)

        # Label text (left-aligned)
        draw_left_text(c, label, BOX_X, ry, LABEL_W, rh, F_LABEL, bold=True)

        # Value text
        if ri == 3:   # Description — may wrap
            draw_wrapped_centered(c, value, BOX_X + LABEL_W, ry, VALUE_W, rh, vsize)
        else:
            draw_centered_text(c, value, BOX_X + LABEL_W, ry, VALUE_W, rh, None, vsize, bold=vbold)

    # ── Store Location row (index 5) ──────────────────────────────────────────
    ri  = 5
    ry  = row_tops[ri]
    rh  = ROWS[ri]

    # Top border line of this row
    c.line(BOX_X, ry + rh, BOX_X + BOX_W, ry + rh)

    # Vertical divider after label
    c.line(BOX_X + LABEL_W, ry, BOX_X + LABEL_W, ry + rh)

    # "Store Location" label — left-aligned, same style as other labels
    draw_left_text(c, "Store Location", BOX_X, ry, LABEL_W, rh, F_LABEL, bold=True)

    # 4 equal boxes inside VALUE_W — each = VALUE_W/4
    loc_box_w = VALUE_W / 4
    for i, part in enumerate(loc_parts):
        lx = BOX_X + LABEL_W + i * loc_box_w
        # Draw vertical divider (left edge of each box, skip first — already drawn)
        if i > 0:
            c.line(lx, ry, lx, ry + rh)
        # Draw value
        draw_centered_text(c, part, lx, ry, loc_box_w, rh, None, F_LOC_VAL, bold=True)

    # ── Barcode / QR row (index 6) ────────────────────────────────────────────
    ri  = 6
    ry  = row_tops[ri]
    rh  = ROWS[ri]

    # Top border of barcode area
    c.line(BOX_X, ry + rh, BOX_X + BOX_W, ry + rh)

    # Build full data string with ALL sticker fields
    store_loc_str = ' | '.join([p for p in loc_parts if p])
    full_data = (
        f"GRN No: {grn_no} | "
        f"GRN Date: {grn_date} | "
        f"Part No: {part_no} | "
        f"Description: {desc_full} | "
        f"Quantity: {qty} | "
        f"Store Location: {store_loc_str}"
    )

    if HAS_QR:
        # ── QR Code — encodes ALL fields, scannable by any QR reader ─────────
        try:
            qr = qrcode.QRCode(
                version=None,
                error_correction=qrcode.constants.ERROR_CORRECT_M,
                box_size=10,
                border=2,
            )
            qr.add_data(full_data)
            qr.make(fit=True)
            qr_img = qr.make_image(fill_color="black", back_color="white")

            buf = BytesIO()
            qr_img.save(buf, format='PNG')
            buf.seek(0)

            # Draw QR centred in the barcode row
            qr_size = rh * 0.88          # QR square size (fits within row height)
            qr_x = BOX_X + (BOX_W - qr_size) / 2
            qr_y = ry + (rh - qr_size) / 2

            from reportlab.lib.utils import ImageReader
            c.drawImage(ImageReader(buf), qr_x, qr_y, width=qr_size, height=qr_size)

            # Label below QR: show GRN No. for quick human reference
            label_y = qr_y - 0.30 * cm
            c.setFont('Helvetica', 7)
            c.drawCentredString(BOX_X + BOX_W / 2, label_y, f"Scan for full details  |  GRN: {grn_no}")

        except Exception as e:
            st.warning(f"QR error row: {e}")
            c.setFont('Helvetica', 8)
            c.drawCentredString(BOX_X + BOX_W / 2, ry + rh / 2, f"[QR: {grn_no}]")
    else:
        # ── Fallback: Code128 barcode encoding GRN No. only ──────────────────
        bc_data = grn_no if grn_no else (part_no if part_no else "NO-DATA")
        pad = 0.5 * cm
        avail_w = BOX_W - 2 * pad
        char_count = max(len(bc_data), 1)
        bar_w = avail_w / (char_count * 11 + 35 + 20)
        bar_w = max(0.55, min(bar_w, 1.8))
        try:
            bc = code128.Code128(
                bc_data,
                barWidth=bar_w,
                barHeight=rh * 0.60,
                humanReadable=True,
                fontSize=F_BC_NUM,
                fontName='Helvetica',
            )
            bc_x = BOX_X + (BOX_W - bc.width) / 2
            bc_y = ry + (rh - bc.height) / 2
            bc.drawOn(c, bc_x, bc_y)
        except Exception as e:
            c.setFont('Helvetica', 9)
            c.drawCentredString(BOX_X + BOX_W / 2, ry + rh / 2, f"[BARCODE: {bc_data}]")
            st.warning(f"Barcode error: {e}")


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
- QR Code encodes ALL fields:
  GRN No, GRN Date, Part No,
  Description, Quantity, Store Location
""")


if __name__ == "__main__":
    main()
