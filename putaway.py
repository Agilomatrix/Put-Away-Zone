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
BOX_H   = 7.5  * cm   # content box height

BOX_X   = (PAGE_W - BOX_W) / 2
BOX_Y   = PAGE_H - 0.20*cm - BOX_H

# ── Column split ──────────────────────────────────────────────────────────────
LABEL_W = BOX_W * 0.33
VALUE_W = BOX_W - LABEL_W

# ── Row heights (8 rows) ──────────────────────────────────────────────────────
R  = 0.72 * cm
RD = 0.90 * cm
RB = 2.28 * cm

ROWS = [R, R, R, RD, R, R, R, RB]

# ── Font sizes ────────────────────────────────────────────────────────────────
F_LABEL   = 12
F_LARGE   = 13
F_MEDIUM  = 11
F_SMALL   = 9
F_LOC_VAL = 10
F_BC_NUM  = 8

LW = 1.0


# ── Helpers ───────────────────────────────────────────────────────────────────

def parse_location(loc_str):
    parts = ['', '', '', '']
    if not loc_str or not isinstance(loc_str, str):
        return parts
    segments = loc_str.strip().split('-')
    for i, s in enumerate(segments[:4]):
        parts[i] = s.strip()
    return parts

def clean_date(val):
    s = str(val) if val and str(val) != 'nan' else ''
    return s.split(' ')[0] if ' ' in s else s

def clean_num(val):
    if val.endswith('.0') and val[:-2].isdigit():
        return val[:-2]
    return val

def draw_left_value(c, text, x, y, w, h, font, size, bold=False):
    """Draw value text left-aligned inside the value cell."""
    fname = 'Helvetica-Bold' if bold else 'Helvetica'
    c.setFont(fname, size)
    while c.stringWidth(text, fname, size) > w - 8 and size > 6:
        size -= 0.5
        c.setFont(fname, size)
    ty = y + (h - size * 0.35) / 2
    c.drawString(x + 4, ty, text)

def draw_left_text(c, text, x, y, w, h, font_size, bold=True):
    fname = 'Helvetica-Bold' if bold else 'Helvetica'
    c.setFont(fname, font_size)
    ty = y + (h - font_size * 0.35) / 2
    c.drawString(x + 4, ty, text)

def draw_wrapped_left(c, text, x, y, w, h, font_size):
    """Draw description text left-aligned with wrapping."""
    style = ParagraphStyle('tmp', fontName='Helvetica', fontSize=font_size,
                           alignment=TA_LEFT, leading=font_size + 2)
    p = Paragraph(text, style)
    pw, ph = p.wrap(w - 8, h)
    py = y + (h - ph) / 2
    p.drawOn(c, x + 4, py)

def draw_centered_text(c, text, x, y, w, h, font, size, bold=False):
    fname = 'Helvetica-Bold' if bold else 'Helvetica'
    c.setFont(fname, size)
    while c.stringWidth(text, fname, size) > w - 6 and size > 6:
        size -= 0.5
        c.setFont(fname, size)
    text_w = c.stringWidth(text, fname, size)
    tx = x + (w - text_w) / 2
    ty = y + (h - size * 0.35) / 2
    c.drawString(tx, ty, text)


# ═══════════════════════════════════════════════════════════════════════════════
# DRAW ONE STICKER
# ═══════════════════════════════════════════════════════════════════════════════

def draw_sticker(c, grn_no, grn_date, part_no, desc, qty, uom,
                 loc_parts, store_loc_raw):
    """Draw sticker. Barcode encodes the raw store location string."""

    # Build the y-coordinate (bottom) of each row, top-down
    row_tops = []
    y = BOX_Y + BOX_H
    for rh in ROWS:
        y -= rh
        row_tops.append(y)

    # Outer border
    c.setStrokeColor(colors.black)
    c.setLineWidth(1.5)
    c.rect(BOX_X, BOX_Y, BOX_W, BOX_H)

    # ── Rows 0-5 (GRN No, Date, Part No, Desc, Qty, UOM) ─────────────────────
    row_defs = [
        (0, "GRN No.",     grn_no,   F_LARGE,  True),
        (1, "GRN Date",    grn_date, F_MEDIUM, True),
        (2, "Part No.",    part_no,  F_LARGE,  True),
        (3, "Description", desc,     F_SMALL,  False),
        (4, "Quantity",    qty,      F_LARGE,  True),
        (5, "UOM",         uom,      F_LARGE,  True),
    ]

    c.setLineWidth(LW)

    for ri, label, value, vsize, vbold in row_defs:
        ry = row_tops[ri]
        rh = ROWS[ri]

        if ri > 0:
            c.line(BOX_X, ry + rh, BOX_X + BOX_W, ry + rh)

        c.line(BOX_X + LABEL_W, ry, BOX_X + LABEL_W, ry + rh)

        draw_left_text(c, label, BOX_X, ry, LABEL_W, rh, F_LABEL, bold=True)

        # ── VALUES: all left-aligned ──────────────────────────────────────────
        if ri == 3:
            draw_wrapped_left(c, value, BOX_X + LABEL_W, ry, VALUE_W, rh, vsize)
        else:
            draw_left_value(c, value, BOX_X + LABEL_W, ry, VALUE_W, rh,
                            None, vsize, bold=vbold)

    # ── Store Location row (index 6) — UNCHANGED (centered per-box) ──────────
    ri  = 6
    ry  = row_tops[ri]
    rh  = ROWS[ri]

    c.line(BOX_X, ry + rh, BOX_X + BOX_W, ry + rh)
    c.line(BOX_X + LABEL_W, ry, BOX_X + LABEL_W, ry + rh)

    draw_left_text(c, "Storage Loc", BOX_X, ry, LABEL_W, rh, F_LABEL, bold=True)

    loc_box_w = VALUE_W / 4
    for i, part in enumerate(loc_parts):
        lx = BOX_X + LABEL_W + i * loc_box_w
        if i > 0:
            c.line(lx, ry, lx, ry + rh)
        draw_centered_text(c, part, lx, ry, loc_box_w, rh, None, F_LOC_VAL, bold=True)

    # ── Barcode row (index 7) — encodes Store Location ────────────────────────
    ri  = 7
    ry  = row_tops[ri]
    rh  = ROWS[ri]

    c.line(BOX_X, ry + rh, BOX_X + BOX_W, ry + rh)

    bc_data = (store_loc_raw.strip() if store_loc_raw and store_loc_raw.strip()
               else (part_no if part_no else (grn_no if grn_no else "NO-DATA")))

    pad     = 0.5 * cm
    avail_w = BOX_W - 2 * pad
    char_count = max(len(bc_data), 1)
    bar_w = avail_w / (char_count * 11 + 35 + 20)
    bar_w = max(0.55, min(bar_w, 1.8))

    try:
        bc = code128.Code128(
            bc_data,
            barWidth=bar_w,
            barHeight=rh * 0.62,
            humanReadable=True,
            fontSize=F_BC_NUM,
            fontName='Helvetica',
        )
        bc_w = bc.width
        bc_x = BOX_X + (BOX_W - bc_w) / 2
        bc_y = ry + (rh - bc.height) / 2
        bc.drawOn(c, bc_x, bc_y)
    except Exception as e:
        c.setFont('Helvetica', 9)
        c.drawCentredString(BOX_X + BOX_W / 2, ry + rh / 2, f"[BARCODE: {bc_data}]")
        st.warning(f"Barcode draw error: {e}")


# ═══════════════════════════════════════════════════════════════════════════════
# COLUMN FINDER HELPER
# ═══════════════════════════════════════════════════════════════════════════════

def find_col(cols, *kw_groups, fallback=None):
    for kwg in kw_groups:
        kwg = [kwg] if isinstance(kwg, str) else list(kwg)
        for col in cols:
            if all(k in col for k in kwg):
                return col
    return fallback


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN GENERATOR
# ═══════════════════════════════════════════════════════════════════════════════

def generate_sticker_labels(df_grn, df_loc):
    """
    df_grn  — File 1: GRN Date, GRN No, Part No, Description, Quantity
    df_loc  — File 2: Part No, Part Description, Store Location, UOM
    Merge on Part No, then generate one sticker per row of df_grn.
    """

    def norm_cols(df):
        df = df.copy()
        df.columns = [c.upper().strip() if isinstance(c, str) else c
                      for c in df.columns]
        return df

    df_grn = norm_cols(df_grn)
    df_loc = norm_cols(df_loc)

    g_cols = df_grn.columns.tolist()
    l_cols = df_loc.columns.tolist()

    grn_no_col   = find_col(g_cols, ['GRN','NO'], ['GRN','NUM'], ['GRN','#'],
                             'GRNNO', 'GRN_NO', 'GRN', fallback=g_cols[0])
    grn_date_col = find_col(g_cols, ['GRN','DATE'], ['RECEIPT','DATE'],
                             ['GRN','DT'], 'DATE', fallback=None)
    part_no_col1 = find_col(g_cols, ['PART','NO'], ['PART','NUM'], ['PART','#'],
                             'PARTNO', 'PART_NO', 'PART', fallback=g_cols[0])
    desc_col1    = find_col(g_cols, 'DESC', 'DESCRIPTION', 'NAME',
                             fallback=g_cols[1] if len(g_cols) > 1 else g_cols[0])
    qty_col      = find_col(g_cols, 'QTY', 'QUANTITY', fallback=None)

    part_no_col2  = find_col(l_cols, ['PART','NO'], ['PART','NUM'], ['PART','#'],
                              'PARTNO', 'PART_NO', 'PART', fallback=l_cols[0])
    store_loc_col = find_col(l_cols, ['STORE','LOC'], 'STORELOCATION',
                              'STORE_LOCATION', 'LOCATION', 'LOC',
                              fallback=l_cols[-1] if l_cols else None)
    uom_col       = find_col(l_cols, 'UOM', 'UNIT OF MEASURE', 'UNIT',
                              fallback=None)

    loc_lookup = {}
    uom_lookup = {}

    for _, row in df_loc.iterrows():
        pn = str(row[part_no_col2]).strip() if pd.notna(row[part_no_col2]) else ''
        sl = (str(row[store_loc_col]).strip()
              if store_loc_col and pd.notna(row[store_loc_col]) else '')
        um = (str(row[uom_col]).strip()
              if uom_col and pd.notna(row[uom_col]) else '')
        if pn:
            for key in (pn, pn.upper(), pn.lower()):
                loc_lookup[key] = sl
                uom_lookup[key] = um

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    tmp_path = tmp.name
    tmp.close()

    c = canvas.Canvas(tmp_path, pagesize=(PAGE_W, PAGE_H))

    progress_bar = st.progress(0)
    status_ph    = st.empty()
    total_rows   = len(df_grn)

    def get(row, col):
        if col and col in row and pd.notna(row[col]):
            v = str(row[col]).strip()
            return clean_num(v)
        return ''

    for idx, (_, row) in enumerate(df_grn.iterrows()):
        progress_bar.progress((idx + 1) / total_rows)
        status_ph.text(f"Creating sticker {idx+1} of {total_rows} "
                       f"({int((idx+1)/total_rows*100)}%)")

        grn_no   = get(row, grn_no_col)
        grn_date = clean_date(get(row, grn_date_col)) if grn_date_col else ''
        part_no  = get(row, part_no_col1)
        desc     = get(row, desc_col1)
        qty      = get(row, qty_col) if qty_col else ''

        store_loc_raw = (loc_lookup.get(part_no)
                         or loc_lookup.get(part_no.upper())
                         or loc_lookup.get(part_no.lower())
                         or '')
        uom = (uom_lookup.get(part_no)
               or uom_lookup.get(part_no.upper())
               or uom_lookup.get(part_no.lower())
               or '')

        loc_parts = parse_location(store_loc_raw)

        draw_sticker(c, grn_no, grn_date, part_no, desc, qty, uom,
                     loc_parts, store_loc_raw)
        c.showPage()

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
        st.header("📁 Upload Files")

        st.subheader("File 1 — GRN Data")
        file1 = st.file_uploader(
            "Upload GRN file (Excel / CSV)  ·  columns: GRN No, GRN Date, Part No, Description, Quantity",
            type=['xlsx', 'xls', 'csv'],
            key="file1",
        )

        st.subheader("File 2 — Store Location Master")
        file2 = st.file_uploader(
            "Upload Location file (Excel / CSV)  ·  columns: Part No, Part Description, Storage Loc, UOM",
            type=['xlsx', 'xls', 'csv'],
            key="file2",
        )

        df_grn = None
        df_loc = None

        if file1:
            try:
                df_grn = (pd.read_csv(file1)
                          if file1.name.lower().endswith('.csv')
                          else pd.read_excel(file1))
                st.success(f"✅ GRN file loaded — {len(df_grn)} rows × {len(df_grn.columns)} columns")
                st.write(f"**Columns:** {', '.join(df_grn.columns.tolist())}")
                st.dataframe(df_grn.head(5), use_container_width=True)
            except Exception as e:
                st.error(f"❌ Error reading GRN file: {e}")

        if file2:
            try:
                df_loc = (pd.read_csv(file2)
                          if file2.name.lower().endswith('.csv')
                          else pd.read_excel(file2))
                st.success(f"✅ Location file loaded — {len(df_loc)} rows × {len(df_loc.columns)} columns")
                st.write(f"**Columns:** {', '.join(df_loc.columns.tolist())}")
                st.dataframe(df_loc.head(5), use_container_width=True)
            except Exception as e:
                st.error(f"❌ Error reading Location file: {e}")

        if df_grn is not None and df_loc is not None:
            st.subheader("🎯 Generate Labels")
            if st.button("🚀 Generate Sticker Labels", type="primary", use_container_width=True):
                with st.spinner("Generating sticker labels…"):
                    pdf_path = generate_sticker_labels(df_grn, df_loc)

                if pdf_path:
                    st.success("🎉 Sticker labels generated successfully!")
                    with open(pdf_path, "rb") as f:
                        pdf_bytes = f.read()

                    filename = f"{file1.name.rsplit('.', 1)[0]}_sticker_labels.pdf"
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
                    st.info(f"📊 File size: {len(pdf_bytes)/1024/1024:.2f} MB | "
                            f"Labels: {len(df_grn)}")

                    try:
                        os.unlink(pdf_path)
                    except Exception:
                        pass
                else:
                    st.error("❌ Failed to generate sticker labels.")

        elif file1 is None and file2 is None:
            st.info("⬆️ Please upload both files above to get started.")
        elif file1 is None:
            st.warning("⚠️ Please also upload File 1 (GRN Data).")
        elif file2 is None:
            st.warning("⚠️ Please also upload File 2 (Store Location Master).")

    with col2:
        st.header("ℹ️ Instructions")
        st.markdown("""
**How to use:**
1. Upload **File 1** — GRN data
2. Upload **File 2** — Store Location master
3. The app matches on **Part No** to fill Store Location & UOM
4. Click **Generate Sticker Labels**
5. Download the PDF

**File 1 expected columns:**
- GRN No. / GRN Number
- GRN Date / Receipt Date
- Part No. / Part Number
- Description / Name
- Quantity / Qty

**File 2 expected columns:**
- Part No. / Part Number
- Part Description
- Storage Loc / Store Location
- UOM / Unit / Unit of Measure

**Label layout (top → bottom):**
1. GRN No.
2. GRN Date
3. Part No.
4. Description
5. Quantity
6. UOM
7. Storage Loc (4-box grid)
8. Barcode (Code-128, **encodes Store Location**)
""")

        st.header("⚙️ Layout")
        st.markdown("""
**Fixed configuration:**
- Sticker page  : 10 × 15 cm
- Content box   : 9.6 × 7.5 cm
- 8 rows, all white background
- Pure canvas drawing — no overflow possible
- Barcode: Code-128, encodes **Storage Loc**
- UOM looked up from File 2 by Part No
""")


if __name__ == "__main__":
    main()
