import streamlit as st
import pandas as pd
from datetime import date, timedelta
from fpdf import FPDF
import tempfile
import os
from PIL import Image
import io
import copy as _copy
import base64
import re

# -- Logo 2 — Agilomatrix logo loaded from fixed path "Image.png"
_LOGO2_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Image.png")

def _load_logo2_bytes():
    try:
        with open(_LOGO2_PATH, "rb") as f:
            return f.read()
    except FileNotFoundError:
        return None

LOGO2_BYTES = _load_logo2_bytes()

# --- App Configuration ---
st.set_page_config(
    page_title="RFQ Generator",
    page_icon="🏭",
    layout="wide"
)

# --- Category suggestion hints ---
CATEGORY_HINTS = {
    "Furniture": ["Office Desk", "Ergonomic Chair", "Conference Table", "Storage Cabinet", "Bookshelf"],
    "Electrical": ["MCB Panel", "Cable Tray", "DB Box", "Power Socket", "LED Fixture"],
    "IT / Electronics": ["Laptop", "Desktop PC", "Network Switch", "UPS", "CCTV Camera"],
    "Civil / Construction": ["Cement Bags", "TMT Steel Bars", "AAC Blocks", "Ready Mix Concrete", "Tiles"],
    "HVAC": ["Split AC Unit", "Ducted AC", "AHU", "Chiller Unit", "FCU"],
    "Plumbing": ["CPVC Pipes", "Ball Valve", "Water Pump", "Pressure Gauge", "Flow Meter"],
    "Interior / Fit-Out": ["False Ceiling", "Partition Wall", "Glass Partition", "Vinyl Flooring", "Acoustic Panel"],
    "Stationery & Office Supplies": ["A4 Paper Ream", "Ballpoint Pens", "Whiteboard", "File Folders", "Stapler"],
    "Warehouse Equipment": [
        "Heavy-duty Racks", "Pallet Racking Systems", "Industrial Shelving",
        "Cantilever Racks", "Mezzanine Floors", "Tabular Racks", "Mobile Storage Racks",
        "Forklifts", "Hand Pallet Trucks", "Electric Pallet Trucks", "Stackers",
        "Trolleys", "Conveyor Systems", "Scissor Lifts",
        "Vertical Carousel System", "Horizontal Carousel System",
        "Dock Levellers", "Dock Plates", "Loading Ramps",
        "Plastic Bins", "Crates", "Pallets (Wood)", "Pallets (Plastic)", "Pallets (Metal)", "Storage Boxes",
        "Rack Protectors", "Column Guards", "Safety Barriers",
        "Safety Mirrors", "Fire Extinguishers", "Safety Signage",
    ],
}

STORAGE_CONTAINERS_ITEMS = [
    "Plastic Bins", "Crates", "Pallets (Wood)", "Pallets (Plastic)", "Pallets (Metal)", "Storage Boxes",
]

UNIT_OPTIONS = ["Nos", "Pieces", "Sets", "Meters", "Sq.Ft", "Sq.M", "Kg", "Tons", "Liters", "Boxes", "Rolls", "Pairs", "Lots"]

# --- SPEC TABLE DATA ---
MODEL_DETAILS_ROWS = [
    {"Sr.no": 1,  "Category": "Dimensions of VStore",        "Description": "Height (mm)",                       "UNIT": "mm",     "Requirement": ""},
    {"Sr.no": "",  "Category": "",                  "Description": "Width (mm)",                        "UNIT": "mm",     "Requirement": ""},
    {"Sr.no": "",  "Category": "",                  "Description": "Depth (mm)",                        "UNIT": "mm",     "Requirement": ""},
    {"Sr.no": "",  "Category": "",                  "Description": "Floor area (m2)",                   "UNIT": "m2",     "Requirement": ""},
    {"Sr.no": "",  "Category": "",                  "Description": "1st Access Point Height (mm)",      "UNIT": "mm",     "Requirement": ""},
    {"Sr.no": "",  "Category": "",                  "Description": "2nd Access Point Height (mm)",      "UNIT": "mm",     "Requirement": ""},
    {"Sr.no": "",  "Category": "",                  "Description": "3rd Access Point Height (mm)",      "UNIT": "mm",     "Requirement": ""},
    {"Sr.no": "",  "Category": "",                  "Description": "4th Access Point Height (mm)",      "UNIT": "mm",     "Requirement": ""},
    {"Sr.no": "",  "Category": "",                  "Description": "Dead weight of Machine (Kg)",       "UNIT": "Kg",     "Requirement": ""},
    {"Sr.no": "",  "Category": "",                  "Description": "Total Weight of Tray (Kg)",         "UNIT": "Kg",     "Requirement": "MS Steel"},
    {"Sr.no": "",  "Category": "",                  "Description": "Total Weight of Machine (Kg)",      "UNIT": "Kg",     "Requirement": "Powder Coated"},
    {"Sr.no": "",  "Category": "",                  "Description": "Storage capacity (Kg)",             "UNIT": "Kg",     "Requirement": ""},
    {"Sr.no": "",  "Category": "",                  "Description": "Total Machine carrying capacity",   "UNIT": "Kg",     "Requirement": ""},
    {"Sr.no": "",  "Category": "",                  "Description": "Total full weight (Kg)",            "UNIT": "Kg",     "Requirement": ""},
    {"Sr.no": 2,  "Category": "Floor Load",         "Description": "Total (Kgs/sqm)",                  "UNIT": "Kg/m2",  "Requirement": ""},
    {"Sr.no": 3,  "Category": "Tray Details",       "Description": "Usable width (mm)",                "UNIT": "mm",     "Requirement": ""},
    {"Sr.no": "",  "Category": "",                  "Description": "Usable depth (mm)",                "UNIT": "mm",     "Requirement": ""},
    {"Sr.no": "",  "Category": "",                  "Description": "Empty Tray weight",                "UNIT": "Kg",     "Requirement": ""},
    {"Sr.no": "",  "Category": "",                  "Description": "Area of each Trays (mm)",          "UNIT": "mm2",    "Requirement": ""},
    {"Sr.no": "",  "Category": "",                  "Description": "Maximum Load capacity (Kg)",       "UNIT": "Kg",     "Requirement": ""},
    {"Sr.no": "",  "Category": "",                  "Description": "Number of Trays (Nos.)",           "UNIT": "Nos",    "Requirement": ""},
    {"Sr.no": "",  "Category": "",                  "Description": "Total area of all Trays (m2)",     "UNIT": "m2",     "Requirement": ""},
    {"Sr.no": 4,  "Category": "Access time",        "Description": "Maximum (Sec.)",                   "UNIT": "Sec",    "Requirement": ""},
    {"Sr.no": "",  "Category": "",                  "Description": "Average (Sec.)",                   "UNIT": "Sec",    "Requirement": ""},
    {"Sr.no": 5,  "Category": "No Trays can Fetch", "Description": "No trays / Hour",                  "UNIT": "Nos/hr", "Requirement": ""},
    {"Sr.no": 6,  "Category": "Power Supply",       "Description": "",                                 "UNIT": "",       "Requirement": ""},
    {"Sr.no": 7,  "Category": "Maximum Power rating","Description": "",                                "UNIT": "kW",     "Requirement": ""},
    {"Sr.no": 8,  "Category": "Control Panel",      "Description": "Standard control panel",          "UNIT": "",       "Requirement": ""},
    {"Sr.no": 9,  "Category": "Height Optimisation","Description": "Provided for storage",             "UNIT": "",       "Requirement": ""},
    {"Sr.no": 10, "Category": "Operator Panel",     "Description": "",                                 "UNIT": "",       "Requirement": ""},
    {"Sr.no": 11, "Category": "Accessories",        "Description": "Emergency stop",                  "UNIT": "",       "Requirement": ""},
    {"Sr.no": "",  "Category": "",                  "Description": "Accident protection light curtains","UNIT": "",      "Requirement": ""},
    {"Sr.no": "",  "Category": "",                  "Description": "Lighting in the accessing area",   "UNIT": "",       "Requirement": ""},
]

KEY_FEATURES_ROWS = [
    {"Sr.no": 1,  "Description": "Material Tracking",                               "Status": "", "Remarks": "All key features to be confirmed by vendor."},
    {"Sr.no": 2,  "Description": "Tray Details",                                    "Status": "", "Remarks": ""},
    {"Sr.no": 3,  "Description": "Inventory List",                                  "Status": "", "Remarks": ""},
    {"Sr.no": 4,  "Description": "Tray Call History",                               "Status": "", "Remarks": ""},
    {"Sr.no": 5,  "Description": "Alarm History",                                   "Status": "", "Remarks": ""},
    {"Sr.no": 6,  "Description": "Item Code Search",                                "Status": "", "Remarks": ""},
    {"Sr.no": 7,  "Description": "Bar Code Search",                                 "Status": "", "Remarks": ""},
    {"Sr.no": 8,  "Description": "Pick from BOM",                                   "Status": "", "Remarks": ""},
    {"Sr.no": 9,  "Description": "BOM Items List",                                  "Status": "", "Remarks": ""},
    {"Sr.no": 10, "Description": "User Management, with backup and restore options","Status": "", "Remarks": ""},
]

INBUILT_FEATURES_ROWS = [
    {"Sr.no": 1,  "Description": "Ergonomic tray positioning",                                     "Vendor Scope (Yes/No)": "", "Remarks": "All features to be included at vendor side."},
    {"Sr.no": 2,  "Description": "Variable frequency drives",                                      "Vendor Scope (Yes/No)": "", "Remarks": ""},
    {"Sr.no": 3,  "Description": "Tray uneven positioning sensor",                                 "Vendor Scope (Yes/No)": "", "Remarks": ""},
    {"Sr.no": 4,  "Description": "Light barrier for sensing material and operator intervention",   "Vendor Scope (Yes/No)": "", "Remarks": ""},
    {"Sr.no": 5,  "Description": "Operator Panel with IPC",                                        "Vendor Scope (Yes/No)": "", "Remarks": ""},
    {"Sr.no": 6,  "Description": "Weight management system for sensing tray overload",             "Vendor Scope (Yes/No)": "", "Remarks": ""},
    {"Sr.no": 7,  "Description": "Tray Block option for Multiple users",                           "Vendor Scope (Yes/No)": "", "Remarks": ""},
    {"Sr.no": 8,  "Description": "Password authentication",                                        "Vendor Scope (Yes/No)": "", "Remarks": ""},
    {"Sr.no": 9,  "Description": "Tray guide rail @ 50 pitch",                                    "Vendor Scope (Yes/No)": "", "Remarks": ""},
    {"Sr.no": 10, "Description": "Total machine capacity 60 tone",                                 "Vendor Scope (Yes/No)": "", "Remarks": ""},
    {"Sr.no": 11, "Description": "Expansion at later stage is possible",                           "Vendor Scope (Yes/No)": "", "Remarks": ""},
    {"Sr.no": 12, "Description": "Inventory management software",                                  "Vendor Scope (Yes/No)": "", "Remarks": ""},
]

INSTALLATION_ROWS = [
    {"Sr.no": 1,  "Category": "Inventory Management Suite (IPC)",                        "Vendor Scope (Yes/No)": "", "Customer Scope (Yes/No)": "", "Remarks": ""},
    {"Sr.no": 2,  "Category": "Packing, Freight & Transit Insurance",                    "Vendor Scope (Yes/No)": "", "Customer Scope (Yes/No)": "", "Remarks": ""},
    {"Sr.no": 3,  "Category": "Installation & Commissioning",                            "Vendor Scope (Yes/No)": "", "Customer Scope (Yes/No)": "", "Remarks": ""},
    {"Sr.no": 4,  "Category": "Training",                                                "Vendor Scope (Yes/No)": "", "Customer Scope (Yes/No)": "", "Remarks": ""},
    {"Sr.no": 5,  "Category": "Warranty Period",                                         "Vendor Scope (Yes/No)": "", "Customer Scope (Yes/No)": "", "Remarks": ""},
    {"Sr.no": 6,  "Category": "Unloading of material",                                  "Vendor Scope (Yes/No)": "", "Customer Scope (Yes/No)": "", "Remarks": ""},
    {"Sr.no": 7,  "Category": "Material handling during the installation",               "Vendor Scope (Yes/No)": "", "Customer Scope (Yes/No)": "", "Remarks": ""},
    {"Sr.no": 8,  "Category": "Power cable cost main junction Box to Machine",           "Vendor Scope (Yes/No)": "", "Customer Scope (Yes/No)": "", "Remarks": ""},
    {"Sr.no": 9,  "Category": "Biometric Access, Barcode Scanner",                      "Vendor Scope (Yes/No)": "", "Customer Scope (Yes/No)": "", "Remarks": ""},
    {"Sr.no": 10, "Category": "MS Office",                                               "Vendor Scope (Yes/No)": "", "Customer Scope (Yes/No)": "", "Remarks": ""},
    {"Sr.no": 11, "Category": "Software Customization",                                  "Vendor Scope (Yes/No)": "", "Customer Scope (Yes/No)": "", "Remarks": ""},
    {"Sr.no": 12, "Category": "Machine Integration with ERP system (extra at Actual)",   "Vendor Scope (Yes/No)": "", "Customer Scope (Yes/No)": "", "Remarks": ""},
    {"Sr.no": 13, "Category": "UPS and Stabilizer with accessories Installation",        "Vendor Scope (Yes/No)": "", "Customer Scope (Yes/No)": "", "Remarks": ""},
    {"Sr.no": 14, "Category": "Equipment Movement & Installation location",              "Vendor Scope (Yes/No)": "", "Customer Scope (Yes/No)": "", "Remarks": ""},
    {"Sr.no": 15, "Category": "PEB Cladding and Civil Floor for outside installation",   "Vendor Scope (Yes/No)": "", "Customer Scope (Yes/No)": "", "Remarks": ""},
]

SPEC_TEMPLATE = {
    "Model Details": MODEL_DETAILS_ROWS,
    "Key Features": KEY_FEATURES_ROWS,
    "Inbuilt features": INBUILT_FEATURES_ROWS,
    "Installation Accountability": INSTALLATION_ROWS,
}

ITEM_TABLE_HEADERS = [
    "Sr.No", "Description", "OL (mm)", "OW (mm)", "OH (mm)",
    "Base Type", "Color", "Weight Kg", "Load Capacity", "LID", "Qty",
    "Conceptual Image"
]
ITEM_TABLE_COL_WIDTHS = [8, 34, 13, 13, 13, 17, 13, 15, 18, 11, 9, 26]

def _empty_container_row(sr=1):
    return {
        "Sr.No": sr, "Description": "",
        "OL (mm)": "", "OW (mm)": "", "OH (mm)": "",
        "Base Type": "Flat", "Color": "",
        "Weight Kg": "", "Load capacity": "", "LID": "No",
        "Qty": 1
    }


# ==============================================================
# TEXT CLEANING UTILITIES
# ==============================================================

def _safe_text(t):
    if not t:
        return ""
    return str(t).encode('latin-1', errors='replace').decode('latin-1')


def _normalize_paragraph(text):
    if not text:
        return ""
    text = re.sub(r'[\t\r]+', ' ', text)
    text = re.sub(r'([.!?])([A-Za-z(])', r'\1 \2', text)
    text = re.sub(r'([,;:])([A-Za-z(])', r'\1 \2', text)
    text = re.sub(r' {2,}', ' ', text)
    return text.strip()


def _prepare_purpose_text(raw_text):
    if not raw_text:
        return []
    raw_paragraphs = re.split(r'\n\s*\n', raw_text)
    result = []
    for para in raw_paragraphs:
        lines = [line.strip() for line in para.split('\n') if line.strip()]
        joined = ' '.join(lines)
        if joined:
            cleaned = _normalize_paragraph(_safe_text(joined))
            if cleaned:
                result.append(cleaned)
    return result


# ==============================================================
# SPEC FILTERING
# ==============================================================

def _is_blank(v):
    return str(v).strip().lower() in ("", "nan", "none")


def _filter_model_details(df):
    """
    FIX: Only keep rows that have a non-empty Requirement value.
    For groups (rows sharing a Sr.no / Category header), we carry the Sr.no
    and Category forward ONLY onto the first kept row in each group, so the
    PDF still shows the group label — but empty-requirement rows are dropped.
    """
    if df is None or df.empty:
        return df

    # ── Step 1: build a flat list with group membership ──────────────────────
    rows_list = []
    current_sr  = ""
    current_cat = ""
    for _, r in df.iterrows():
        sr  = str(r.get("Sr.no",    "")).strip()
        cat = str(r.get("Category", "")).strip()
        req = str(r.get("Requirement", "")).strip()
        desc = str(r.get("Description", "")).strip()
        unit = str(r.get("UNIT", r.get("Unit", ""))).strip()

        # A new group starts when Sr.no or Category is non-empty
        if sr != "" or cat != "":
            current_sr  = sr
            current_cat = cat

        rows_list.append({
            "sr":   current_sr,
            "cat":  current_cat,
            "desc": desc,
            "unit": unit,
            "req":  req,
        })

    # ── Step 2: keep only rows where Requirement is filled ───────────────────
    kept = [r for r in rows_list if not _is_blank(r["req"])]

    if not kept:
        return pd.DataFrame()   # nothing to show

    # ── Step 3: for each group that has kept rows, show Sr.no & Category only
    #            on the FIRST kept row of that group ──────────────────────────
    seen_groups = set()
    result_rows = []
    for r in kept:
        group_key = (r["sr"], r["cat"])
        if group_key not in seen_groups:
            seen_groups.add(group_key)
            result_rows.append({
                "Sr.no":       r["sr"],
                "Category":    r["cat"],
                "Description": r["desc"],
                "UNIT":        r["unit"],
                "Requirement": r["req"],
            })
        else:
            result_rows.append({
                "Sr.no":       "",
                "Category":    "",
                "Description": r["desc"],
                "UNIT":        r["unit"],
                "Requirement": r["req"],
            })

    return pd.DataFrame(result_rows)


def _filter_navy_df(df, value_cols):
    if df is None or df.empty:
        return df

    def _has_value(row):
        return any(not _is_blank(row.get(c, "")) for c in value_cols)

    return df[df.apply(_has_value, axis=1)].reset_index(drop=True)


# ==============================================================
# PDF GENERATION
# ==============================================================
def create_advanced_rfq_pdf(data):

    class PDF(FPDF):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self._data = data
            self.set_auto_page_break(auto=True, margin=38)

        def header(self):
            if self.page_no() == 1:
                return

            logo1_data = self._data.get('logo1_data')
            logo2_data = LOGO2_BYTES
            logo1_w = self._data.get('logo1_w', 35)
            logo1_h = self._data.get('logo1_h', 18)
            logo2_w = 45
            logo2_h = 20
            header_h = max(logo1_h if logo1_data else 0, logo2_h if logo2_data else 0, 10)
            logo_y = 6

            if logo1_data:
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                        tmp.write(logo1_data)
                        tmp.flush()
                        self.image(tmp.name, x=self.l_margin,
                                   y=logo_y + (header_h - logo1_h) / 2,
                                   w=logo1_w, h=logo1_h)
                    os.remove(tmp.name)
                except Exception:
                    pass

            if logo2_data:
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                        tmp.write(logo2_data)
                        tmp.flush()
                        self.image(tmp.name, x=self.w - self.r_margin - logo2_w,
                                   y=logo_y + (header_h - logo2_h) / 2,
                                   w=logo2_w, h=logo2_h)
                    os.remove(tmp.name)
                except Exception:
                    pass

            title_text = 'Request for Quotation (RFQ)'
            self.set_font('Arial', 'B', 11)
            left_end = self.l_margin + (logo1_w + 4 if logo1_data else 0)
            right_start = self.w - self.r_margin - (logo2_w + 4 if logo2_data else 0)
            mid_w = right_start - left_end
            if mid_w > 20:
                title_h = 6
                title_y = logo_y + (header_h - title_h) / 2
                self.set_xy(left_end, title_y)
                self.cell(mid_w, title_h, title_text, 0, 0, 'C')

            self.set_y(logo_y + header_h + 3)
            self.set_draw_color(180, 180, 180)
            self.line(self.l_margin, self.get_y(), self.w - self.r_margin, self.get_y())
            self.set_draw_color(0, 0, 0)
            self.ln(3)

        def footer(self):
            self.set_y(-30)
            self.set_draw_color(180, 180, 180)
            self.line(self.l_margin, self.get_y(), self.w - self.r_margin, self.get_y())
            self.set_draw_color(0, 0, 0)
            self.ln(1)

            self.set_font('Arial', 'B', 9)
            self.set_text_color(80, 80, 80)
            self.cell(0, 5, 'APL-Confidential', 0, 1, 'R')
            self.ln(1)

            fn = self._data.get('footer_company_name', 'Agilomatrix Private Ltd')
            self.set_font('Arial', 'B', 13)
            self.set_text_color(0, 0, 0)
            self.cell(0, 6, fn, 0, 1, 'C')

            fa = self._data.get('footer_company_address',
                                'Registered Office: F1403, 7 Plumeria Drive, 7PD Street, Tathawade, Pune - 411033')
            self.set_font('Arial', '', 8)
            self.set_text_color(120, 120, 120)
            self.cell(0, 5, fa, 0, 1, 'C')

            self.set_font('Arial', '', 8)
            self.cell(0, 5, f'Page {self.page_no()}/{{nb}}', 0, 0, 'C')
            self.set_text_color(0, 0, 0)

        def section_title(self, title):
            self.set_font('Arial', 'B', 12)
            self.set_fill_color(26, 58, 92)
            self.set_text_color(255, 255, 255)
            usable = self.w - self.l_margin - self.r_margin
            sx = self.l_margin
            sy = self.get_y()
            sh = 9
            self.rect(sx, sy, usable, sh, 'F')
            self.set_xy(sx + 2, sy + 2)
            self.multi_cell(usable - 4, 6, f'  {title}', border=0, align='L')
            self.set_text_color(0, 0, 0)
            self.set_y(sy + sh)
            self.ln(3)

    def _clean(v):
        if v is None or str(v).strip().lower() in ("nan", "none", ""):
            return ""
        try:
            f = float(v)
            return str(int(f)) if f == int(f) else str(f)
        except Exception:
            return str(v).strip()

    def _write_logo(pdf, logo_data, x, y, w, h):
        if not logo_data:
            return
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                tmp.write(logo_data)
                tmp.flush()
                pdf.image(tmp.name, x=x, y=y, w=w, h=h)
            os.remove(tmp.name)
        except Exception:
            pass

    # ── COVER PAGE ────────────────────────────────────────────────────────────
    def create_cover_page(pdf):
        pdf.add_page()
        logo2_w = 45
        logo2_h = 20
        _write_logo(pdf, data.get('logo1_data'), pdf.l_margin, 12,
                    data.get('logo1_w', 35), data.get('logo1_h', 18))
        _write_logo(pdf, LOGO2_BYTES, pdf.w - pdf.r_margin - logo2_w, 12, logo2_w, logo2_h)

        pdf.set_y(35)
        pdf.set_font('Arial', 'B', 14)
        pdf.set_text_color(200, 0, 0)
        pdf.cell(0, 8, 'CONFIDENTIAL', 0, 1, 'L')
        pdf.set_text_color(0, 0, 0)
        pdf.ln(8)

        pdf.set_font('Arial', 'B', 28)
        pdf.cell(0, 14, 'Request for Quotation', 0, 1, 'C')
        pdf.ln(4)
        pdf.set_font('Arial', 'I', 16)
        pdf.cell(0, 8, 'for', 0, 1, 'C')
        pdf.ln(4)
        pdf.set_font('Arial', 'B', 20)
        pdf.cell(0, 10, data.get('Type_of_items', ''), 0, 1, 'C')
        pdf.ln(6)
        pdf.set_font('Arial', '', 16)
        pdf.cell(0, 8, 'At', 0, 1, 'C')
        pdf.ln(4)
        pdf.set_font('Arial', 'B', 20)
        pdf.cell(0, 10, data.get('Storage', ''), 0, 1, 'C')
        pdf.ln(8)
        pdf.set_font('Arial', 'B', 22)
        pdf.cell(0, 10, data.get('company_name', ''), 0, 1, 'C')
        pdf.ln(2)
        pdf.set_font('Arial', '', 14)
        pdf.cell(0, 8, data.get('company_address', ''), 0, 1, 'C')

    # ── MODEL DETAILS TABLE ───────────────────────────────────────────────────
    def render_model_details(pdf, df, subtitle=""):
        if df is None or df.empty:
            return

        df = _filter_model_details(df)
        if df is None or df.empty:
            return

        cw = [10, 42, 72, 22, 44]
        total_w = sum(cw)
        rh = 8
        header_fill = (220, 230, 241)
        req_fill = (255, 255, 204)

        def draw_col_headers():
            pdf.set_fill_color(*header_fill)
            pdf.set_font('Arial', 'B', 9)
            col_y = pdf.get_y()
            col_h = 14
            labels = ["Sr.no", "Category", "Description", "UNIT", "Requirement"]
            for i, c in enumerate(labels):
                cpl = max(1, int(cw[i] / 2.2))
                col_h = max(col_h, max(1, -(-len(c) // cpl)) * 5 + 6)
            cx = pdf.l_margin
            for i, c in enumerate(labels):
                cpl = max(1, int(cw[i] / 2.2))
                n_lines = max(1, -(-len(c) // cpl))
                top_pad = max(1, (col_h - n_lines * 5) / 2)
                pdf.rect(cx, col_y, cw[i], col_h, 'FD')
                pdf.set_xy(cx + 1, col_y + top_pad)
                pdf.multi_cell(cw[i] - 2, 5, c, border=0, align='C')
                cx += cw[i]
                pdf.set_xy(cx, col_y)
            pdf.set_y(col_y + col_h)

        pdf.set_font('Arial', 'B', 12)
        pdf.set_fill_color(26, 58, 92)
        pdf.set_text_color(255, 255, 255)
        hdr_x = pdf.l_margin
        hdr_y = pdf.get_y()
        hdr_h = 9
        pdf.rect(hdr_x, hdr_y, total_w, hdr_h, 'F')
        pdf.set_xy(hdr_x + 2, hdr_y + 2)
        pdf.multi_cell(total_w - 4, 6, '  Model Details', border=0, align='L')
        pdf.set_text_color(0, 0, 0)
        pdf.set_y(hdr_y + hdr_h)

        if subtitle:
            pdf.set_font('Arial', 'B', 9)
            pdf.set_fill_color(240, 240, 240)
            pdf.cell(total_w, 7, subtitle, border=1, ln=1, align='C', fill=True)

        draw_col_headers()

        # ── Render rows ──────────────────────────────────────────────────────
        # After _filter_model_details, the DataFrame is already flat:
        # Sr.no & Category are set only on the first row of each group.
        # We re-group here so we can still span the left cells vertically.

        rows_list = []
        for _, r in df.iterrows():
            rows_list.append({
                "sr":   _clean(r.get("Sr.no",       "")),
                "cat":  _clean(r.get("Category",    "")),
                "desc": _clean(r.get("Description", "")),
                "unit": _clean(r.get("UNIT", r.get("Unit", ""))),
                "req":  _clean(r.get("Requirement", ""))
            })

        # Re-group consecutive rows that share a (sr, cat) header
        groups = []
        if rows_list:
            curr = [rows_list[0]]
            for item in rows_list[1:]:
                if item["sr"] != "" or item["cat"] != "":
                    groups.append(curr)
                    curr = [item]
                else:
                    curr.append(item)
            groups.append(curr)

        pdf.set_font('Arial', '', 9)
        for grp in groups:
            group_h = len(grp) * rh
            if pdf.get_y() + group_h > pdf.page_break_trigger:
                pdf.add_page()
                draw_col_headers()
                pdf.set_font('Arial', '', 9)

            sy = pdf.get_y()
            sx = pdf.l_margin

            pdf.rect(sx, sy, cw[0], group_h)
            pdf.set_xy(sx, sy + (group_h - 5) / 2)
            pdf.cell(cw[0], 5, grp[0]["sr"], align='C')

            pdf.rect(sx + cw[0], sy, cw[1], group_h)
            pdf.set_xy(sx + cw[0] + 1, sy + 1)
            pdf.multi_cell(cw[1] - 2, 5, grp[0]["cat"], border=0, align='L')

            for idx, item in enumerate(grp):
                ry = sy + idx * rh
                rx = sx + cw[0] + cw[1]

                pdf.rect(rx, ry, cw[2], rh)
                pdf.set_xy(rx + 1, ry + 1)
                pdf.multi_cell(cw[2] - 2, 5, item["desc"], border=0, align='L')
                pdf.set_xy(rx + cw[2], ry)

                pdf.rect(rx + cw[2], ry, cw[3], rh)
                pdf.set_xy(rx + cw[2], ry + 1)
                pdf.cell(cw[3], rh - 2, item["unit"], align='C')

                pdf.set_fill_color(*req_fill)
                pdf.rect(rx + cw[2] + cw[3], ry, cw[4], rh, 'FD')
                pdf.set_xy(rx + cw[2] + cw[3] + 1, ry + 1)
                pdf.multi_cell(cw[4] - 2, 5, item["req"], border=0, align='C')
                pdf.set_fill_color(255, 255, 255)

            pdf.set_y(sy + group_h)
        pdf.ln(5)

    # ── CUSTOM SPEC TABLE ─────────────────────────────────────────────────────
    def render_custom_spec_table(pdf, custom_tables):
        """
        Renders one or more user-defined spec tables in the PDF.

        custom_tables: list of dicts, each with:
            {
              'title':   str,           # dark navy header text
              'columns': [str, ...],    # user-defined column names (2-5)
              'df':      pd.DataFrame,  # rows; columns match 'columns' list
            }
        """
        if not custom_tables:
            return

        USABLE_W   = 190   # A4 usable width in mm (10mm margins each side)
        SR_W       = 10    # fixed Sr.No column width
        header_fill = (220, 230, 241)
        rh_min      = 8

        for tbl in custom_tables:
            title    = _safe_text(tbl.get('title', 'Technical Specification'))
            user_cols = tbl.get('columns', [])
            df        = tbl.get('df', pd.DataFrame())

            if not user_cols or df is None or df.empty:
                continue

            # Drop rows where ALL user columns are blank
            def _row_has_data(row):
                return any(not _is_blank(row.get(c, '')) for c in user_cols)
            df = df[df.apply(_row_has_data, axis=1)].reset_index(drop=True)
            if df.empty:
                continue

            # Build column list: Sr.No always first, then user columns
            all_cols = ['Sr.No'] + user_cols
            n_user   = len(user_cols)

            # Distribute remaining width evenly across user columns
            remaining = USABLE_W - SR_W
            # First user col gets slightly more (it's usually the label column)
            if n_user == 1:
                col_widths = [remaining]
            elif n_user == 2:
                col_widths = [round(remaining * 0.45), round(remaining * 0.55)]
            elif n_user == 3:
                col_widths = [round(remaining * 0.38), round(remaining * 0.32), round(remaining * 0.30)]
            elif n_user == 4:
                col_widths = [round(remaining * 0.32), round(remaining * 0.25), round(remaining * 0.23), round(remaining * 0.20)]
            else:  # 5
                col_widths = [round(remaining * 0.28), round(remaining * 0.20), round(remaining * 0.20), round(remaining * 0.17), round(remaining * 0.15)]
            # Fix rounding so total = remaining
            col_widths[-1] += remaining - sum(col_widths)
            all_widths = [SR_W] + col_widths
            total_w    = sum(all_widths)

            # ── Title bar ────────────────────────────────────────────────────
            if pdf.get_y() + 30 > pdf.page_break_trigger:
                pdf.add_page()

            pdf.set_fill_color(26, 58, 92)
            pdf.set_text_color(255, 255, 255)
            pdf.set_font('Arial', 'B', 11)
            ty = pdf.get_y()
            pdf.rect(pdf.l_margin, ty, total_w, 9, 'F')
            pdf.set_xy(pdf.l_margin + 3, ty + 1.5)
            pdf.cell(total_w - 6, 6, f'  {title}', border=0)
            pdf.set_text_color(0, 0, 0)
            pdf.set_y(ty + 9)
            pdf.ln(1)

            # ── Column header row ─────────────────────────────────────────────
            def draw_col_headers():
                pdf.set_fill_color(*header_fill)
                pdf.set_font('Arial', 'B', 9)
                hy  = pdf.get_y()
                hh  = 12
                # Compute needed height across all headers
                for i, c in enumerate(all_cols):
                    cpl = max(1, int(all_widths[i] / 2.0))
                    hh  = max(hh, -(-len(c) // cpl) * 5 + 4)
                cx = pdf.l_margin
                for i, c in enumerate(all_cols):
                    pdf.rect(cx, hy, all_widths[i], hh, 'FD')
                    pdf.set_xy(cx + 1, hy + 2)
                    pdf.multi_cell(all_widths[i] - 2, 5, _safe_text(c), border=0, align='C')
                    cx += all_widths[i]
                    pdf.set_xy(cx, hy)
                pdf.set_y(hy + hh)

            draw_col_headers()
            pdf.set_font('Arial', '', 9)

            # ── Data rows ─────────────────────────────────────────────────────
            for row_i, (_, row) in enumerate(df.iterrows()):
                vals = [str(row_i + 1)] + [_clean(row.get(c, '')) for c in user_cols]

                # Compute row height
                row_h = rh_min
                for j, val in enumerate(vals):
                    cpl   = max(1, int(all_widths[j] / 1.85))
                    row_h = max(row_h, -(-len(val) // cpl) * 5 + 3)

                if pdf.get_y() + row_h > pdf.page_break_trigger:
                    pdf.add_page()
                    draw_col_headers()
                    pdf.set_font('Arial', '', 9)

                row_y = pdf.get_y()
                cx    = pdf.l_margin
                for j, val in enumerate(vals):
                    pdf.rect(cx, row_y, all_widths[j], row_h)
                    pdf.set_xy(cx + 1, row_y + 1)
                    pdf.multi_cell(all_widths[j] - 2, 5, val, border=0,
                                   align='C' if j == 0 else 'L')
                    cx += all_widths[j]
                    pdf.set_xy(cx, row_y)
                pdf.set_y(row_y + row_h)

            pdf.ln(5)

    # ── NAVY SECTION TABLE ────────────────────────────────────────────────────
    def render_navy_section(pdf, title, df, cols, widths):
        if df is None or df.empty:
            return

        value_cols = [c for c in cols if c not in ("Sr.no", "Category", "Description", "Remarks")]
        df = _filter_navy_df(df, value_cols)
        if df is None or df.empty:
            return

        total_w = sum(widths)
        remarks_col = cols[-1]
        remark_text = ""
        for _, row in df.iterrows():
            v = _clean(row.get(remarks_col, ""))
            if v:
                remark_text = v
                break

        if pdf.get_y() + 35 > pdf.page_break_trigger:
            pdf.add_page()

        pdf.set_fill_color(26, 58, 92)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font('Arial', 'B', 11)
        ty = pdf.get_y()
        pdf.rect(pdf.l_margin, ty, total_w, 9, 'F')
        pdf.set_xy(pdf.l_margin + 3, ty + 1.5)
        pdf.cell(total_w - 6, 6, f'  {title}', border=0)
        pdf.set_text_color(0, 0, 0)
        pdf.set_y(ty + 9)
        pdf.ln(1)

        def draw_headers():
            pdf.set_fill_color(220, 230, 241)
            pdf.set_font('Arial', 'B', 9)
            hy = pdf.get_y()
            hh = 14
            for i, c in enumerate(cols):
                lbl = c.strip()
                cpl = max(1, int(widths[i] / 2.2))
                hh = max(hh, -(-len(lbl) // cpl) * 5 + 6)
            cx = pdf.l_margin
            for i, c in enumerate(cols):
                lbl = c.strip()
                pdf.rect(cx, hy, widths[i], hh, 'FD')
                cpl = max(1, int(widths[i] / 2.2))
                n_lines = -(-len(lbl) // cpl)
                top_pad = max(1, (hh - n_lines * 5) / 2)
                pdf.set_xy(cx + 1, hy + top_pad)
                pdf.multi_cell(widths[i] - 2, 5, lbl, border=0, align='C')
                cx += widths[i]
                pdf.set_xy(cx, hy)
            pdf.set_y(hy + hh)
            return hh

        draw_headers()
        pdf.set_font('Arial', '', 9)

        for row_num, (_, row) in enumerate(df.iterrows()):
            vals = []
            for c in cols:
                if c == remarks_col:
                    vals.append(remark_text if row_num == 0 else "")
                else:
                    vals.append(_clean(row.get(c, "")))

            rh = 8
            for i, val in enumerate(vals):
                cpl = max(1, int(widths[i] / 1.85))
                rh = max(rh, -(-len(val) // cpl) * 5 + 3)

            if pdf.get_y() + rh > pdf.page_break_trigger:
                pdf.add_page()
                draw_headers()
                pdf.set_font('Arial', '', 9)

            row_y = pdf.get_y()
            cx = pdf.l_margin
            for i, val in enumerate(vals):
                pdf.rect(cx, row_y, widths[i], rh)
                pdf.set_xy(cx + 1, row_y + 1)
                pdf.multi_cell(widths[i] - 2, 5, val, border=0,
                               align='L' if i <= 1 else 'C')
                cx += widths[i]
                pdf.set_xy(cx, row_y)
            pdf.set_y(row_y + rh)

        pdf.ln(4)

    # ── STORAGE CONTAINER TABLE ───────────────────────────────────────────────
    def render_container_table(pdf, df, images_dict=None):
        headers = ITEM_TABLE_HEADERS
        cw = ITEM_TABLE_COL_WIDTHS
        hh = 14
        rh = 30
        IMG_W, IMG_H = 22, 21

        def draw_header():
            pdf.set_font("Arial", "B", 10)
            pdf.set_fill_color(220, 230, 241)
            sy = pdf.get_y()
            cx = pdf.l_margin
            for i, h in enumerate(headers):
                pdf.rect(cx, sy, cw[i], hh, 'FD')
                pdf.set_xy(cx, sy + 2)
                pdf.multi_cell(cw[i], 4, h, align="C")
                cx += cw[i]
            pdf.set_y(sy + hh)

        draw_header()
        pdf.set_font("Arial", "", 10)

        if df is not None and not df.empty:
            for idx, row in df.iterrows():
                ry = pdf.get_y()
                if ry + rh > pdf.page_break_trigger:
                    pdf.add_page()
                    draw_header()
                    pdf.set_font("Arial", "", 10)
                    ry = pdf.get_y()

                vals = [
                    str(idx + 1), _clean(row.get("Description")),
                    _clean(row.get("OL (mm)")), _clean(row.get("OW (mm)")),
                    _clean(row.get("OH (mm)")), _clean(row.get("Base Type")),
                    _clean(row.get("Color")), _clean(row.get("Weight Kg")),
                    _clean(row.get("Load capacity")), _clean(row.get("LID")),
                    _clean(row.get("Qty")), ""
                ]

                cx = pdf.l_margin
                for i, val in enumerate(vals):
                    pdf.rect(cx, ry, cw[i], rh)
                    if headers[i] == "Conceptual Image":
                        img_bytes = row.get("image_data_bytes")
                        if not isinstance(img_bytes, bytes) and images_dict:
                            img_bytes = images_dict.get(idx)
                        if isinstance(img_bytes, bytes):
                            try:
                                img = Image.open(io.BytesIO(img_bytes))
                                ix = cx + (cw[i] - IMG_W) / 2
                                iy = ry + (rh - IMG_H) / 2
                                with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                                    img.save(tmp.name, "PNG")
                                    pdf.image(tmp.name, x=ix, y=iy, w=IMG_W, h=IMG_H)
                                os.remove(tmp.name)
                            except Exception:
                                pass
                    else:
                        pdf.set_xy(cx + 1, ry + (rh - 6) / 2)
                        pdf.multi_cell(cw[i] - 2, 6, val, align="C")
                    cx += cw[i]
                pdf.set_y(ry + rh)
        pdf.ln(6)

    # ── GENERIC ITEMS TABLE ───────────────────────────────────────────────────
    def render_generic_items(pdf, df):
        if df is None or df.empty:
            return
        cols = ["Sr.No", "Item Name", "Description / Specification", "Quantity", "Unit", "Remarks"]
        widths = [12, 40, 70, 18, 18, 32]
        total_w = sum(widths)

        pdf.set_fill_color(220, 230, 241)
        pdf.set_font('Arial', 'B', 11)
        ch_y = pdf.get_y()
        ch_h = 9
        for i, c in enumerate(cols):
            ch_h = max(ch_h, max(1, -(-len(c) // max(1, int(widths[i] / 2.5)))) * 6 + 3)
        cx = pdf.l_margin
        for i, c in enumerate(cols):
            pdf.rect(cx, ch_y, widths[i], ch_h, 'FD')
            pdf.set_xy(cx + 1, ch_y + 1)
            pdf.multi_cell(widths[i] - 2, 6, c, border=0, align='C')
            cx += widths[i]
            pdf.set_xy(cx, ch_y)
        pdf.set_y(ch_y + ch_h)
        pdf.set_font('Arial', '', 9)

        for i, (_, row) in enumerate(df.iterrows()):
            vals = [
                str(i + 1), _clean(row.get("Item Name")),
                _clean(row.get("Description / Specification")),
                _clean(row.get("Quantity")), _clean(row.get("Unit")),
                _clean(row.get("Remarks"))
            ]
            rh = 8
            if pdf.get_y() + rh > pdf.page_break_trigger:
                pdf.add_page()
                pdf.set_fill_color(220, 230, 241)
                pdf.set_font('Arial', 'B', 11)
                cy2 = pdf.get_y()
                cx2 = pdf.l_margin
                for j, c in enumerate(cols):
                    pdf.rect(cx2, cy2, widths[j], ch_h, 'FD')
                    pdf.set_xy(cx2 + 1, cy2 + 1)
                    pdf.multi_cell(widths[j] - 2, 6, c, border=0, align='C')
                    cx2 += widths[j]
                    pdf.set_xy(cx2, cy2)
                pdf.set_y(cy2 + ch_h)
                pdf.set_font('Arial', '', 9)
            row_y = pdf.get_y()
            cx = pdf.l_margin
            for j, val in enumerate(vals):
                pdf.rect(cx, row_y, widths[j], rh)
                pdf.set_xy(cx + 1, row_y + 1)
                pdf.multi_cell(widths[j] - 2, 5, val, border=0,
                               align='L' if j <= 2 else 'C')
                cx += widths[j]
                pdf.set_xy(cx, row_y)
            pdf.set_y(row_y + rh)
        pdf.ln(5)

    # ── LAYOUT IMAGES ─────────────────────────────────────────────────────────
    def render_layout_images(pdf, layout_images):
        if not layout_images:
            return
        pdf.add_page()
        pdf.set_fill_color(26, 58, 92)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, '  LAYOUT', border=0, ln=1, align='L', fill=True)
        pdf.set_text_color(0, 0, 0)
        pdf.ln(6)

        usable_w = pdf.w - pdf.l_margin - pdf.r_margin
        n = len(layout_images)

        if n == 1:
            img_w = usable_w * 0.75
            img_h = img_w * 0.65
            img_x = pdf.l_margin + (usable_w - img_w) / 2
            _place_image(pdf, layout_images[0], img_x, pdf.get_y(), img_w, img_h)
            pdf.set_y(pdf.get_y() + img_h + 4)
        elif n == 2:
            img_w = (usable_w - 6) / 2
            img_h = img_w * 0.7
            y = pdf.get_y()
            _place_image(pdf, layout_images[0], pdf.l_margin, y, img_w, img_h)
            _place_image(pdf, layout_images[1], pdf.l_margin + img_w + 6, y, img_w, img_h)
            pdf.set_y(y + img_h + 4)
        else:
            img_w = (usable_w - 6) / 2
            img_h = img_w * 0.65
            for i in range(0, n, 2):
                y = pdf.get_y()
                if y + img_h > pdf.page_break_trigger:
                    pdf.add_page()
                    y = pdf.get_y()
                _place_image(pdf, layout_images[i], pdf.l_margin, y, img_w, img_h)
                if i + 1 < n:
                    _place_image(pdf, layout_images[i + 1], pdf.l_margin + img_w + 6, y, img_w, img_h)
                pdf.set_y(y + img_h + 6)
        pdf.ln(4)

    def _place_image(pdf, img_bytes, x, y, w, h):
        if not isinstance(img_bytes, bytes):
            return
        try:
            img = Image.open(io.BytesIO(img_bytes))
            iw, ih = img.size
            ratio = min(w / iw, h / ih)
            draw_w, draw_h = iw * ratio, ih * ratio
            cx = x + (w - draw_w) / 2
            cy = y + (h - draw_h) / 2
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                img.save(tmp.name, "PNG")
                pdf.image(tmp.name, x=cx, y=cy, w=draw_w, h=draw_h)
            os.remove(tmp.name)
            pdf.rect(x, y, w, h)
        except Exception:
            pass

    # ── BUILD PDF ─────────────────────────────────────────────────────────────
    pdf = PDF('P', 'mm', 'A4')
    pdf._data = data
    pdf.alias_nb_pages()

    create_cover_page(pdf)
    pdf.add_page()

    # 1. REQUIREMENT BACKGROUND
    pdf.section_title('REQUIREMENT BACKGROUND')
    pdf.set_font('Arial', '', 11)
    usable_w = pdf.w - pdf.l_margin - pdf.r_margin

    raw_purpose = data.get('purpose', '')
    paragraphs = _prepare_purpose_text(raw_purpose)
    if paragraphs:
        for para_text in paragraphs:
            if pdf.get_y() + 12 > pdf.page_break_trigger:
                pdf.add_page()
            pdf.set_x(pdf.l_margin)
            pdf.multi_cell(usable_w, 7, para_text, border=0, align='L')
            pdf.ln(3)
    else:
        pdf.set_x(pdf.l_margin)
        pdf.multi_cell(usable_w, 7, _safe_text(raw_purpose), border=0, align='L')
    pdf.ln(5)

    # 2. TECHNICAL SPECIFICATION
    pdf.section_title('TECHNICAL SPECIFICATION')
    rfq_category   = data.get('rfq_category', 'General')
    wh_sub         = data.get('wh_sub', '')
    use_custom_spec = data.get('use_custom_spec', False)

    # ── Custom table path (overrides standard spec tables) ───────────────────
    if use_custom_spec:
        render_custom_spec_table(pdf, data.get('custom_tables', []))
        render_layout_images(pdf, data.get('layout_images', []))

    elif rfq_category == "Warehouse Equipment":
        if wh_sub == "Storage Container":
            render_container_table(pdf, data.get('storage_containers_df', pd.DataFrame()),
                                   data.get('storage_containers_images', {}))
            render_layout_images(pdf, data.get('layout_images', []))

        elif wh_sub == "Automated Storage System":
            wh_items = data.get('wh_items_df', pd.DataFrame())
            if wh_items is not None and not wh_items.empty:
                valid_items = wh_items[wh_items.get("Item Name", pd.Series(dtype=str)).astype(str).str.strip() != ""]
                if not valid_items.empty:
                    pdf.set_font('Arial', 'B', 10)
                    pdf.cell(0, 7, 'Items Required:', 0, 1)
                    pdf.set_font('Arial', '', 10)
                    for _, r in valid_items.iterrows():
                        item_line = f"  - {_clean(r.get('Item Name', ''))}   Qty: {_clean(r.get('Quantity', ''))} {_clean(r.get('Unit', ''))}"
                        if _clean(r.get('Description / Specification', '')):
                            item_line += f"   -   {_clean(r.get('Description / Specification', ''))}"
                        pdf.cell(0, 6, item_line, 0, 1)
                    pdf.ln(4)

            render_model_details(pdf, data.get('carousel_model_df'), subtitle=data.get('model_detail_header', ''))
            render_navy_section(pdf, "Key Features", data.get('key_features_df'),
                                ["Sr.no", "Description", "Status", "Remarks"], [10, 102, 38, 40])
            render_navy_section(pdf, "Inbuilt features", data.get('inbuilt_features_df'),
                                ["Sr.no", "Description", "Vendor Scope (Yes/No)", "Remarks"], [10, 102, 38, 40])
            render_navy_section(pdf, "Installation Accountability", data.get('installation_df'),
                                ["Sr.no", "Category", "Vendor Scope (Yes/No)", "Customer Scope (Yes/No)", "Remarks"],
                                [10, 67, 36, 36, 41])
            render_layout_images(pdf, data.get('layout_images', []))

        else:
            pfx = {'Storage System': 'ss', 'Material Handling': 'mh', 'Dock Leveller': 'dl'}.get(wh_sub, 'ss')
            render_model_details(pdf, data.get(f'spec_{pfx}_Model Details'),
                                 subtitle=data.get('model_detail_header', ''))
            render_navy_section(pdf, "Key Features", data.get(f'spec_{pfx}_Key Features'),
                                ["Sr.no", "Description", "Status", "Remarks"], [10, 102, 38, 40])
            render_navy_section(pdf, "Inbuilt features", data.get(f'spec_{pfx}_Inbuilt features'),
                                ["Sr.no", "Description", "Vendor Scope (Yes/No)", "Remarks"], [10, 102, 38, 40])
            render_navy_section(pdf, "Installation Accountability",
                                data.get(f'spec_{pfx}_Installation Accountability'),
                                ["Sr.no", "Category", "Vendor Scope (Yes/No)", "Customer Scope (Yes/No)", "Remarks"],
                                [10, 67, 36, 36, 41])
            render_layout_images(pdf, data.get('layout_images', []))
    else:
        render_generic_items(pdf, data.get('items_df', pd.DataFrame()))

    # 3. QUOTATION SUBMISSION & DELIVERY
    if pdf.get_y() + 40 > pdf.page_break_trigger:
        pdf.add_page()
    pdf.section_title('QUOTATION SUBMISSION & DELIVERY')
    pdf.set_font('Arial', 'B', 11)
    pdf.cell(0, 7, f"Quotation to be submit to: {data.get('submit_to_name', '')}", 0, 1)
    if data.get('submit_to_registered_office'):
        pdf.set_font('Arial', '', 10)
        pdf.cell(0, 6, data.get('submit_to_registered_office', ''), 0, 1)
    pdf.ln(3)
    pdf.set_font('Arial', 'B', 11)
    pdf.cell(0, 7, 'Delivery Location:', 0, 1)
    pdf.set_font('Arial', '', 11)
    del_company = _safe_text(data.get('delivery_company', ''))
    del_gstin   = _safe_text(data.get('delivery_gstin', ''))
    del_address = _normalize_paragraph(_safe_text(data.get('delivery_address', '')))
    if del_company:
        pdf.set_font('Arial', 'B', 11)
        pdf.set_x(pdf.l_margin)
        pdf.multi_cell(usable_w, 7, del_company, border=0, align='L')
    if del_gstin:
        pdf.set_font('Arial', '', 11)
        pdf.set_x(pdf.l_margin)
        pdf.multi_cell(usable_w, 7, f'GSTIN No: {del_gstin}', border=0, align='L')
    if del_address:
        pdf.set_font('Arial', '', 11)
        pdf.set_x(pdf.l_margin)
        pdf.multi_cell(usable_w, 7, f'Address: {del_address}', border=0, align='L')

    annexures = data.get('annexures', '').strip()
    if annexures:
        pdf.ln(5)
        pdf.section_title('ANNEXURES')
        pdf.set_font('Arial', '', 11)
        for line in annexures.split('\n'):
            if line.strip():
                pdf.cell(0, 7, f'  - {_safe_text(line.strip())}', 0, 1)
    pdf.ln(5)

    # 4. TIMELINES
    if pdf.get_y() + 60 > pdf.page_break_trigger:
        pdf.add_page()
    pdf.section_title('TIMELINES')
    milestones = [
        ('RFQ to Vendor',                      data.get('date_release')),
        ('1st Level Discussion',               data.get('date_query')),
        ('Techno Commercial Offer',            data.get('date_meet')),
        ('2nd Level Discussion on Proposal',   data.get('date_quote')),
        ('Final Techno Commercial Offer',      data.get('date_selection')),
        ('PO to Vendor',                       data.get('date_review')),
        ('Delivery at Site',                   data.get('date_delivery')),
        ('Installation at Site',               data.get('date_install')),
    ]
    pdf.set_fill_color(220, 230, 241)
    pdf.set_font('Arial', 'B', 11)
    pdf.cell(90, 9, 'Milestone', 1, 0, 'C', fill=True)
    pdf.cell(100, 9, 'Date', 1, 1, 'C', fill=True)
    pdf.set_font('Arial', '', 11)
    for m, d in milestones:
        date_str = d.strftime('%B %d, %Y') if d else 'TBD'
        pdf.cell(90, 8, m, 1, 0, 'L')
        pdf.cell(100, 8, date_str, 1, 1, 'L')
    pdf.ln(5)

    # 5. SPOC
    if pdf.get_y() + 50 > pdf.page_break_trigger:
        pdf.add_page()
    pdf.section_title('SINGLE POINT OF CONTACT')
    pdf.ln(4)
    col_w = usable_w / 2
    lbl_w = 28
    has_spoc2 = bool(data.get('spoc2_name', '').strip())

    hy = pdf.get_y()
    pdf.set_xy(pdf.l_margin, hy)
    pdf.set_font('Arial', 'BU', 12)
    pdf.cell(col_w, 8, 'Primary Contact', 0, 0, 'L')
    if has_spoc2:
        pdf.set_xy(pdf.l_margin + col_w, hy)
        pdf.cell(col_w, 8, 'Secondary Contact', 0, 0, 'L')
    pdf.ln(11)

    spoc_rows = [
        ('Name:',     'spoc1_name',  'spoc2_name'),
        ('Phone No:', 'spoc1_phone', 'spoc2_phone'),
        ('Email ID:', 'spoc1_email', 'spoc2_email'),
    ]
    for label, k1, k2 in spoc_rows:
        ry = pdf.get_y()
        pdf.set_xy(pdf.l_margin, ry)
        pdf.set_font('Arial', 'B', 11)
        pdf.cell(lbl_w, 8, label, 0, 0, 'L')
        pdf.set_font('Arial', '', 11)
        pdf.cell(col_w - lbl_w, 8, _safe_text(data.get(k1, '')), 0, 0, 'L')
        if has_spoc2:
            pdf.set_xy(pdf.l_margin + col_w, ry)
            pdf.set_font('Arial', 'B', 11)
            pdf.cell(lbl_w, 8, label, 0, 0, 'L')
            pdf.set_font('Arial', '', 11)
            pdf.cell(col_w - lbl_w, 8, _safe_text(data.get(k2, '')), 0, 0, 'L')
        pdf.ln(9)
    pdf.ln(5)

    # LAST PAGE: Sign-off
    pdf.add_page()
    page_w = pdf.w - pdf.l_margin - pdf.r_margin

    def _field_line(label, line_w=110):
        pdf.set_font('Arial', '', 10)
        lbl_w2 = pdf.get_string_width(label + '  ')
        pdf.cell(lbl_w2, 7, label, 0, 0, 'L')
        x1 = pdf.get_x()
        y1 = pdf.get_y() + 6.2
        pdf.line(x1, y1, x1 + line_w, y1)
        pdf.ln(8)

    pdf.ln(4)
    pdf.set_font('Arial', 'B', 11)
    pdf.set_fill_color(240, 244, 248)
    pdf.set_draw_color(180, 180, 180)
    pdf.cell(page_w, 8, '  Buyer Information', border='B', ln=1, align='L', fill=True)
    pdf.set_draw_color(0, 0, 0)
    pdf.ln(4)
    _field_line('Company Name:         ')
    _field_line('Contact Information:  ')
    _field_line('Date of Issue:        ')
    _field_line('RFQ Reference Number:')
    pdf.ln(6)

    pdf.set_font('Arial', 'B', 11)
    pdf.set_fill_color(240, 244, 248)
    pdf.set_draw_color(180, 180, 180)
    pdf.cell(page_w, 8, '  Supplier Information', border='B', ln=1, align='L', fill=True)
    pdf.set_draw_color(0, 0, 0)
    pdf.ln(4)
    _field_line('Supplier Name:   ')
    _field_line('Contact Person:  ')
    _field_line('Contact Details: ')
    pdf.ln(6)

    pdf.set_font('Arial', 'BI', 12)
    pdf.set_text_color(26, 58, 92)
    pdf.cell(0, 8, 'Terms and Conditions', 0, 1, 'L')
    pdf.set_text_color(0, 0, 0)
    terms = [
        'Prices quoted should include all applicable taxes and duties.',
        'Delivery timeline must be clearly mentioned.',
        'Payment terms to be agreed upon before order confirmation.',
        'This RFQ does not constitute a commitment to purchase.',
        'The company reserves the right to accept or reject any quotation.',
    ]
    pdf.set_font('Arial', '', 10)
    num_col_w = 8
    txt_col_w = page_w - num_col_w
    for idx, term in enumerate(terms, 1):
        row_y = pdf.get_y()
        pdf.set_xy(pdf.l_margin, row_y)
        pdf.cell(num_col_w, 6, f'{idx}.', 0, 0, 'L')
        pdf.set_xy(pdf.l_margin + num_col_w, row_y)
        pdf.multi_cell(txt_col_w, 6, _safe_text(term), 0, 'L')
        pdf.ln(1)
    pdf.ln(10)

    _field_line('Authorized Signatory: ')
    _field_line('Designation:          ')
    _field_line('Date:                 ')

    return bytes(pdf.output())


# ==============================================================
# STREAMLIT UI
# ==============================================================
st.title("🏭 Request For Quotation Generator")
st.markdown("---")

# Step 1: Logo
with st.expander("Step 1: Upload Your Company Logo (Optional)", expanded=True):
    st.markdown("**Logo 1 — Your company logo (top-left of every page)**")
    logo1_file = st.file_uploader("Upload Logo 1", type=['png', 'jpg', 'jpeg'], key="logo1")
    lc1, lc2 = st.columns(2)
    logo1_w = lc1.number_input("Width (mm)", 5, 80, 35, 1, key="l1w")
    logo1_h = lc2.number_input("Height (mm)", 5, 50, 18, 1, key="l1h")
    if logo1_file:
        st.image(logo1_file, width=160)
    if LOGO2_BYTES:
        st.success("✅ Agilomatrix logo (Image.png) loaded — appears automatically on every page (top-right).")
    else:
        st.warning("⚠️ Image.png not found next to app.py. Place your Agilomatrix logo as **Image.png** in the same folder as app.py and restart.")

# Step 2: Cover page
with st.expander("Step 2: Add Cover Page Details", expanded=True):
    Type_of_items = st.text_input("Type of Items *", help="e.g., Plastic Blue Bins OR Line Side Racks")
    Storage = st.text_input("Storage Type *", help="e.g., Material Storage")
    company_name = st.text_input("Requester Company Name *", help="e.g., Pinnacle Mobility Solutions Pvt. Ltd")
    company_address = st.text_input("Requester Company Address *", help="e.g., Nanekarwadi, Chakan, Pune 410501")

# Step 3: Footer
with st.expander("Step 3: Add Footer Details (Optional)", expanded=False):
    footer_company_name = st.text_input("Footer Company Name")
    footer_company_address = st.text_input("Footer Company Address")

# Step 4: Technical Specifications
st.subheader("Step 4: Technical Specifications")
st.markdown("---")

with st.expander("📦 Technical Specifications", expanded=True):
    category_list = list(CATEGORY_HINTS.keys())

    def _on_category_change():
        for k in ['dynamic_items_df', 'storage_containers_df', 'storage_containers_images',
                  'wh_items_df', 'carousel_model_df', 'key_features_df',
                  'inbuilt_features_df', 'installation_df',
                  'spec_ss_Model Details', 'spec_ss_Key Features', 'spec_ss_Inbuilt features', 'spec_ss_Installation Accountability',
                  'spec_mh_Model Details', 'spec_mh_Key Features', 'spec_mh_Inbuilt features', 'spec_mh_Installation Accountability',
                  'spec_dl_Model Details', 'spec_dl_Key Features', 'spec_dl_Inbuilt features', 'spec_dl_Installation Accountability',
                  'custom_spec_df']:
            st.session_state.pop(k, None)
        st.session_state.pop('wh_sub_select', None)

    rfq_category = st.selectbox(
        "Select RFQ Category *",
        options=category_list,
        key="rfq_category_select",
        on_change=_on_category_change,
    )

    is_warehouse = (rfq_category == "Warehouse Equipment")
    WH_SUB_CATEGORIES = ["Storage System", "Material Handling", "Automated Storage System", "Dock Leveller", "Storage Container"]

    if is_warehouse:
        def _on_wh_sub_change():
            for k in ['wh_items_df', 'storage_containers_df', 'storage_containers_images',
                      'carousel_model_df', 'key_features_df', 'inbuilt_features_df', 'installation_df',
                      'spec_ss_Model Details', 'spec_ss_Key Features', 'spec_ss_Inbuilt features', 'spec_ss_Installation Accountability',
                      'spec_mh_Model Details', 'spec_mh_Key Features', 'spec_mh_Inbuilt features', 'spec_mh_Installation Accountability',
                      'spec_dl_Model Details', 'spec_dl_Key Features', 'spec_dl_Inbuilt features', 'spec_dl_Installation Accountability',
                      'custom_spec_df']:
                st.session_state.pop(k, None)

        wh_sub = st.selectbox("Select Warehouse Sub-Category *", options=WH_SUB_CATEGORIES,
                              key="wh_sub_select", on_change=_on_wh_sub_change)
    else:
        wh_sub = ""

    # ── _render_multisection_spec ─────────────────────────────────────────────
    def _render_multisection_spec(state_key_prefix):
        section_cfg = {
            "Model Details": {
                "cols": ["Sr.no", "Category", "Description", "UNIT", "Requirement"],
                "column_config": {
                    "Sr.no":       st.column_config.TextColumn("Sr.no", width="small"),
                    "Category":    st.column_config.TextColumn("Category", width="medium"),
                    "Description": st.column_config.TextColumn("Description", width="large"),
                    "UNIT":        st.column_config.TextColumn("UNIT", width="small"),
                    "Requirement": st.column_config.TextColumn("Requirement ✏️", width="medium"),
                },
            },
            "Key Features": {
                "cols": ["Sr.no", "Description", "Status", "Remarks"],
                "column_config": {
                    "Sr.no":       st.column_config.TextColumn("Sr.no", width="small"),
                    "Description": st.column_config.TextColumn("Description", width="large"),
                    "Status":      st.column_config.TextColumn("Status ✏️", width="small"),
                    "Remarks":     st.column_config.TextColumn("Remarks", width="large"),
                },
            },
            "Inbuilt features": {
                "cols": ["Sr.no", "Description", "Vendor Scope (Yes/No)", "Remarks"],
                "column_config": {
                    "Sr.no":                 st.column_config.TextColumn("Sr.no", width="small"),
                    "Description":           st.column_config.TextColumn("Description", width="large"),
                    "Vendor Scope (Yes/No)": st.column_config.SelectboxColumn("Vendor Scope ✏️", width="small", options=["", "Yes", "No"]),
                    "Remarks":               st.column_config.TextColumn("Remarks", width="large"),
                },
            },
            "Installation Accountability": {
                "cols": ["Sr.no", "Category", "Vendor Scope (Yes/No)", "Customer Scope (Yes/No)", "Remarks"],
                "column_config": {
                    "Sr.no":                   st.column_config.TextColumn("Sr.no", width="small"),
                    "Category":                st.column_config.TextColumn("Category", width="large"),
                    "Vendor Scope (Yes/No)":   st.column_config.SelectboxColumn("Vendor Scope ✏️", width="small", options=["", "Yes", "No"]),
                    "Customer Scope (Yes/No)": st.column_config.SelectboxColumn("Customer Scope ✏️", width="small", options=["", "Yes", "No"]),
                    "Remarks":                 st.column_config.TextColumn("Remarks", width="medium"),
                },
            },
        }

        for section_name, rows in SPEC_TEMPLATE.items():
            fkey = f"frozen_{state_key_prefix}_{section_name}"
            dkey = f"data_{state_key_prefix}_{section_name}"
            wkey = f"widget_{state_key_prefix}_{section_name}"
            cfg  = section_cfg[section_name]

            st.markdown(
                f"<div style='background:#1a3a5c;color:white;font-weight:bold;"
                f"padding:6px 10px;margin-top:14px;margin-bottom:2px;"
                f"font-size:14px;border-radius:3px;'>{section_name}</div>",
                unsafe_allow_html=True
            )

            if fkey not in st.session_state:
                init_df = pd.DataFrame(_copy.deepcopy(rows))
                for col in cfg["cols"]:
                    if col not in init_df.columns:
                        init_df[col] = ""
                    init_df[col] = init_df[col].astype(str).replace("nan", "")
                st.session_state[fkey] = init_df[cfg["cols"]].copy()
                st.session_state[dkey] = init_df[cfg["cols"]].copy()

            def _make_cb(_wkey, _fkey, _dkey):
                def _cb():
                    delta = st.session_state.get(_wkey)
                    frozen = st.session_state.get(_fkey)
                    if not isinstance(delta, dict) or frozen is None:
                        return
                    df = frozen.copy()
                    for row_idx_str, changes in delta.get('edited_rows', {}).items():
                        row_idx = int(row_idx_str)
                        if row_idx < len(df):
                            for col, val in changes.items():
                                if col in df.columns:
                                    df.at[row_idx, col] = val
                    for new_row in delta.get('added_rows', []):
                        row_data = {c: new_row.get(c, '') for c in df.columns}
                        df = pd.concat([df, pd.DataFrame([row_data])], ignore_index=True)
                    del_indices = sorted(delta.get('deleted_rows', []), reverse=True)
                    for i in del_indices:
                        if i < len(df):
                            df = df.drop(df.index[i]).reset_index(drop=True)
                    st.session_state[_dkey] = df
                return _cb

            st.data_editor(
                st.session_state[fkey],
                num_rows="dynamic",
                use_container_width=True,
                column_config=cfg["column_config"],
                key=wkey,
                on_change=_make_cb(wkey, fkey, dkey),
            )

    def _render_layout_uploader(prefix):
        sk = f"layout_images_{prefix}"
        if sk not in st.session_state:
            st.session_state[sk] = []
        st.markdown("---")
        st.markdown(
            "<div style='background:#1a3a5c;color:white;font-weight:bold;"
            "padding:8px 12px;margin-bottom:8px;font-size:15px;border-radius:3px;'>"
            "📐 Layout Images (1 to 5)</div>", unsafe_allow_html=True
        )
        st.caption("Upload layout drawings, front/side views, or 3D renders. Min 1, Max 5.")
        uploaded = []
        cols = st.columns(5)
        for i in range(5):
            with cols[i]:
                f = st.file_uploader(
                    f"Image {i+1}" + (" *" if i == 0 else " (optional)"),
                    type=["png", "jpg", "jpeg"], key=f"layout_img_{prefix}_{i}"
                )
                if f is not None:
                    uploaded.append(f.getvalue())
                    st.image(f.getvalue(), use_container_width=True)
                elif i < len(st.session_state[sk]):
                    uploaded.append(st.session_state[sk][i])
                    st.image(st.session_state[sk][i], use_container_width=True)
        st.session_state[sk] = [b for b in uploaded if b]
        n = len(st.session_state[sk])
        if n > 0:
            st.success(f"✅ {n} layout image(s) ready for PDF")
        else:
            st.info("Upload at least 1 layout image to include the Layout section in the PDF.")

    # ══════════════════════════════════════════════════════════════════════════
    # CUSTOM SPEC TABLE UI
    # User defines: table title, number of columns (2-5), column names, rows.
    # Multiple tables supported — each with its own title/columns/rows.
    # ══════════════════════════════════════════════════════════════════════════
    def _render_custom_spec_editor(prefix):
        """
        Lets the user build one or more fully custom spec tables.
        Each table has:
          - An editable title (shown as the dark navy bar in the PDF)
          - 2 to 5 user-named columns
          - Dynamic rows filled in a data_editor
        """
        st.markdown(
            "<div style='background:#2e7d32;color:white;font-weight:bold;"
            "padding:8px 12px;margin-bottom:8px;font-size:15px;border-radius:3px;'>"
            "✏️ Create Your Own Specification Table(s)</div>",
            unsafe_allow_html=True
        )

        # ── How many tables? ──────────────────────────────────────────────────
        n_tables_key = f"custom_n_tables_{prefix}"
        if n_tables_key not in st.session_state:
            st.session_state[n_tables_key] = 1

        col_add, col_remove, _ = st.columns([1, 1, 6])
        with col_add:
            if st.button("➕ Add Table", key=f"add_tbl_{prefix}"):
                st.session_state[n_tables_key] = min(10, st.session_state[n_tables_key] + 1)
                st.rerun()
        with col_remove:
            if st.session_state[n_tables_key] > 1:
                if st.button("➖ Remove Last", key=f"rem_tbl_{prefix}"):
                    st.session_state[n_tables_key] -= 1
                    st.rerun()

        n_tables = st.session_state[n_tables_key]

        for tbl_idx in range(n_tables):
            tbl_pfx = f"{prefix}_t{tbl_idx}"

            st.markdown(
                f"<div style='background:#37474f;color:white;font-weight:bold;"
                f"padding:5px 10px;margin-top:16px;margin-bottom:6px;"
                f"font-size:13px;border-radius:3px;'>Table {tbl_idx + 1}</div>",
                unsafe_allow_html=True
            )

            # ── Table Title ───────────────────────────────────────────────────
            title_key = f"custom_title_{tbl_pfx}"
            if title_key not in st.session_state:
                st.session_state[title_key] = "Technical Specification"
            st.session_state[title_key] = st.text_input(
                "Table Title (shown as header in PDF)",
                value=st.session_state[title_key],
                key=f"title_input_{tbl_pfx}",
                placeholder="e.g. Model Details / Key Features / Dimensions",
            )

            # ── Number of columns ─────────────────────────────────────────────
            ncols_key = f"custom_ncols_{tbl_pfx}"
            if ncols_key not in st.session_state:
                st.session_state[ncols_key] = 3

            n_cols = st.slider(
                "Number of columns",
                min_value=2, max_value=10,
                value=st.session_state[ncols_key],
                key=f"ncols_slider_{tbl_pfx}",
            )
            # Detect column-count change and wipe frozen/data so editor rebuilds
            if n_cols != st.session_state[ncols_key]:
                st.session_state[ncols_key] = n_cols
                for k in [f"custom_frozen_{tbl_pfx}", f"custom_data_{tbl_pfx}"]:
                    st.session_state.pop(k, None)
                st.rerun()

            # ── Column name inputs ────────────────────────────────────────────
            col_name_keys = [f"custom_colname_{tbl_pfx}_{i}" for i in range(n_cols)]
            default_names = ["Parameter", "Value", "Unit", "Remarks", "Notes", "Col 6", "Col 7", "Col 8", "Col 9", "Col 10"]
            label_cols = st.columns(n_cols)
            user_col_names = []
            for i, lc in enumerate(label_cols):
                ck = col_name_keys[i]
                if ck not in st.session_state:
                    st.session_state[ck] = default_names[i] if i < len(default_names) else f"Col {i+1}"
                prev_name = st.session_state[ck]
                new_name = lc.text_input(
                    f"Column {i+1} name",
                    value=prev_name,
                    key=f"colname_input_{tbl_pfx}_{i}",
                )
                # If column name changed, wipe frozen so editor rebuilds with new headers
                if new_name != prev_name:
                    st.session_state[ck] = new_name
                    for k in [f"custom_frozen_{tbl_pfx}", f"custom_data_{tbl_pfx}"]:
                        st.session_state.pop(k, None)
                    st.rerun()
                user_col_names.append(st.session_state[ck])

            # ── Data Editor ───────────────────────────────────────────────────
            fkey = f"custom_frozen_{tbl_pfx}"
            dkey = f"custom_data_{tbl_pfx}"
            wkey = f"custom_widget_{tbl_pfx}"

            if fkey not in st.session_state:
                init_df = pd.DataFrame([{c: "" for c in user_col_names}])
                st.session_state[fkey] = init_df.copy()
                st.session_state[dkey] = init_df.copy()

            def _make_custom_cb(_wkey, _fkey, _dkey, _cols):
                def _cb():
                    delta  = st.session_state.get(_wkey)
                    frozen = st.session_state.get(_fkey)
                    if not isinstance(delta, dict) or frozen is None:
                        return
                    df = frozen.copy()
                    for row_idx_str, changes in delta.get("edited_rows", {}).items():
                        row_idx = int(row_idx_str)
                        if row_idx < len(df):
                            for col, val in changes.items():
                                if col in df.columns:
                                    df.at[row_idx, col] = val
                    for new_row in delta.get("added_rows", []):
                        row_data = {c: new_row.get(c, "") for c in df.columns}
                        df = pd.concat([df, pd.DataFrame([row_data])], ignore_index=True)
                    del_indices = sorted(delta.get("deleted_rows", []), reverse=True)
                    for i in del_indices:
                        if i < len(df):
                            df = df.drop(df.index[i]).reset_index(drop=True)
                    st.session_state[_dkey] = df
                return _cb

            st.data_editor(
                st.session_state[fkey],
                num_rows="dynamic",
                use_container_width=True,
                column_config={c: st.column_config.TextColumn(c, width="medium")
                               for c in user_col_names},
                key=wkey,
                on_change=_make_custom_cb(wkey, fkey, dkey, user_col_names),
            )

            current_df = st.session_state.get(dkey, pd.DataFrame())
            if not current_df.empty:
                def _has_any(row):
                    return any(str(row.get(c, "")).strip() for c in user_col_names if c in row)
                filled = sum(1 for _, r in current_df.iterrows() if _has_any(r))
                if filled:
                    st.success(f"✅ {filled} row(s) defined in Table {tbl_idx + 1}")
            st.markdown("---")

    # ──────────────────────────────────────────────────────────────────────────
    # Render per sub-category / category
    # ──────────────────────────────────────────────────────────────────────────
    if is_warehouse:
        if wh_sub in ("Storage System", "Material Handling", "Dock Leveller"):
            pfx = {'Storage System': 'ss', 'Material Handling': 'mh', 'Dock Leveller': 'dl'}[wh_sub]
            st.markdown(f"#### 📋 {wh_sub} Specification")

            # ── TABLE MODE TOGGLE ─────────────────────────────────────────────
            st.markdown("**Specification Table Mode:**")
            table_mode = st.radio(
                "Choose how to define the Technical Specification:",
                options=["📋 Use Standard Spec Tables", "✏️ Create My Own Table"],
                index=0,
                key=f"table_mode_{pfx}",
                horizontal=True,
                label_visibility="collapsed",
            )
            use_custom = (table_mode == "✏️ Create My Own Table")

            if use_custom:
                _render_custom_spec_editor(pfx)
            else:
                st.caption("Pre-filled from standard template. Edit the **Requirement / Status / Vendor Scope** columns.")
                st.info("💡 Only rows where you fill in a value (Requirement / Status / Vendor Scope / Customer Scope) will appear in the PDF.")
                model_header_pfx = st.text_input(
                    "Model Header / Subtitle (shown under Model Details in PDF)",
                    placeholder="e.g.  3400 (L) x 3200 (W)  -  465 kgs/tray  -  28 m Height",
                    key=f"model_detail_header_{pfx}"
                )
                _render_multisection_spec(pfx)

            _render_layout_uploader(pfx)

        elif wh_sub == "Automated Storage System":
            st.markdown("#### 📋 Automated Storage System")

            # ── TABLE MODE TOGGLE ─────────────────────────────────────────────
            st.markdown("**Specification Table Mode:**")
            table_mode = st.radio(
                "Choose how to define the Technical Specification:",
                options=["📋 Use Standard Spec Tables", "✏️ Create My Own Table"],
                index=0,
                key="table_mode_carousel",
                horizontal=True,
                label_visibility="collapsed",
            )
            use_custom = (table_mode == "✏️ Create My Own Table")

            item_opts = [""] + ["Vertical Carousel System", "Horizontal Carousel System"]
            if 'wh_items_df' not in st.session_state:
                st.session_state['wh_items_df'] = pd.DataFrame([{
                    "Sr.no": 1, "Item Name": "",
                    "Description / Specification": "", "Quantity": 1, "Unit": "Nos", "Remarks": ""
                }])
            df_items = st.session_state['wh_items_df'].copy()
            df_items["Sr.no"] = range(1, len(df_items) + 1)
            df_items["Quantity"] = pd.to_numeric(df_items.get("Quantity", 1), errors='coerce').fillna(1).astype(int)
            for col in ["Item Name", "Description / Specification", "Unit", "Remarks"]:
                if col not in df_items.columns: df_items[col] = ""
                df_items[col] = df_items[col].astype(str).replace("nan", "")

            edited_items = st.data_editor(
                df_items[["Sr.no", "Item Name", "Description / Specification", "Quantity", "Unit", "Remarks"]],
                num_rows="dynamic", use_container_width=True,
                column_config={
                    "Sr.no": st.column_config.NumberColumn("Sr.no", width="small", disabled=True),
                    "Item Name": st.column_config.SelectboxColumn("Item Name ▼", width="medium", options=item_opts),
                    "Description / Specification": st.column_config.TextColumn("Description / Specification", width="large"),
                    "Quantity": st.column_config.NumberColumn("Qty", width="small", min_value=0, step=1, format="%d"),
                    "Unit": st.column_config.SelectboxColumn("Unit ▼", width="small", options=[""] + UNIT_OPTIONS),
                    "Remarks": st.column_config.TextColumn("Remarks", width="medium"),
                }, key="wh_items_editor")
            st.session_state['wh_items_df'] = edited_items

            st.markdown("---")
            if use_custom:
                _render_custom_spec_editor("carousel")
            else:
                st.info("💡 Only rows where you fill in a value (Requirement / Status / Vendor Scope / Customer Scope) will appear in the PDF.")
                model_header = st.text_input("Model Header / Subtitle",
                                             placeholder="e.g.  3400 (L) x 3200 (W)  -  465 kgs/tray  -  28 m Height",
                                             key="model_detail_header_carousel")
                st.markdown("#### 📐 Full Specification Tables")
                _render_multisection_spec("carousel")

            _render_layout_uploader("carousel")

        elif wh_sub == "Storage Container":
            st.caption("Fill in container details. Type any container name freely in the Description column.")

            SC_COLS = ["Sr.No", "Description", "OL (mm)", "OW (mm)", "OH (mm)",
                       "Base Type", "Color", "Weight Kg", "Load capacity", "LID", "Qty"]

            if "sc_frozen" not in st.session_state:
                init_df = pd.DataFrame([_empty_container_row(1)])
                for col in SC_COLS:
                    if col not in init_df.columns:
                        init_df[col] = ""
                    init_df[col] = init_df[col].astype(str).replace("nan", "")
                init_df["Qty"] = pd.to_numeric(init_df["Qty"], errors="coerce").fillna(1).astype(int)
                init_df["Sr.No"] = 1
                st.session_state["sc_frozen"] = init_df[SC_COLS].copy()
                st.session_state["sc_data"]   = init_df[SC_COLS].copy()

            if "storage_containers_images" not in st.session_state:
                st.session_state["storage_containers_images"] = {}

            SC_COLS_NO_SR = [c for c in SC_COLS if c != "Sr.No"]

            def _sc_current_row_count():
                frozen = st.session_state.get("sc_frozen", pd.DataFrame())
                delta  = st.session_state.get("sc_wkey")
                n = len(frozen)
                if isinstance(delta, dict):
                    n += len(delta.get("added_rows", []))
                    n -= len(delta.get("deleted_rows", []))
                return max(1, n)

            def _sc_on_change():
                delta  = st.session_state.get("sc_wkey")
                frozen = st.session_state.get("sc_frozen")
                if not isinstance(delta, dict) or frozen is None:
                    return
                df = frozen[SC_COLS_NO_SR].copy()
                for row_idx_str, changes in delta.get("edited_rows", {}).items():
                    row_idx = int(row_idx_str)
                    if row_idx < len(df):
                        for col, val in changes.items():
                            if col in df.columns:
                                df.at[row_idx, col] = val
                for new_row in delta.get("added_rows", []):
                    row_data = {c: new_row.get(c, "") for c in df.columns}
                    df = pd.concat([df, pd.DataFrame([row_data])], ignore_index=True)
                del_indices = sorted(delta.get("deleted_rows", []), reverse=True)
                for i in del_indices:
                    if i < len(df):
                        df = df.drop(df.index[i]).reset_index(drop=True)
                df.insert(0, "Sr.No", range(1, len(df) + 1))
                st.session_state["sc_data"] = df[SC_COLS]
                st.session_state["storage_containers_df"] = st.session_state["sc_data"]

            editor_col, img_col = st.columns([4, 1])
            with editor_col:
                st.data_editor(
                    st.session_state["sc_frozen"][SC_COLS_NO_SR],
                    num_rows="dynamic",
                    use_container_width=True,
                    column_config={
                        "Description":   st.column_config.TextColumn("Container / Item Name", width="large"),
                        "OL (mm)":       st.column_config.TextColumn("OL (mm)", width="small"),
                        "OW (mm)":       st.column_config.TextColumn("OW (mm)", width="small"),
                        "OH (mm)":       st.column_config.TextColumn("OH (mm)", width="small"),
                        "Base Type":     st.column_config.SelectboxColumn("Base Type ▼", width="small", options=["", "Flat", "Ribbed", "Louvred", "Grid", "Plain", "Other"]),
                        "Color":         st.column_config.TextColumn("Color", width="small"),
                        "Weight Kg":     st.column_config.TextColumn("Weight Kg", width="small"),
                        "Load capacity": st.column_config.TextColumn("Load Cap (Kg)", width="small"),
                        "LID":           st.column_config.SelectboxColumn("LID ▼", width="small", options=["", "Yes", "No", "N/A"]),
                        "Qty":           st.column_config.NumberColumn("Qty", width="small", min_value=0, step=1),
                    },
                    key="sc_wkey",
                    on_change=_sc_on_change,
                )

            with img_col:
                st.write("**Conceptual Images**")
                n_rows = _sc_current_row_count()
                current_sc = st.session_state["sc_data"]
                for i in range(n_rows):
                    desc = ""
                    if i < len(current_sc):
                        desc = str(current_sc.iloc[i].get("Description", "")).strip()
                    lbl = f"Row {i+1}: {desc}" if desc else f"Row {i+1}"
                    f_up = st.file_uploader(lbl, type=["png", "jpg", "jpeg"], key=f"sc_img_{i}")
                    if f_up is not None:
                        st.session_state["storage_containers_images"][i] = f_up.getvalue()
                    if i in st.session_state["storage_containers_images"]:
                        st.image(st.session_state["storage_containers_images"][i], width=80)

            st.session_state["storage_containers_df"] = st.session_state["sc_data"]

            valid_count = len(st.session_state["sc_data"][
                st.session_state["sc_data"]["Description"].astype(str).str.strip() != ""])
            if valid_count:
                st.success(f"✅ {valid_count} container type(s) defined")
            _render_layout_uploader("sc")

    else:
        # ── Non-warehouse categories ──────────────────────────────────────────
        hints = CATEGORY_HINTS.get(rfq_category, [])
        if hints:
            st.markdown(f"**💡 Common items in *{rfq_category}*:**")
            btn_cols = st.columns(min(len(hints), 4))
            for idx, hint in enumerate(hints):
                with btn_cols[idx % 4]:
                    if st.button(f"➕ {hint}", key=f"hint_{idx}_{rfq_category}"):
                        if 'dynamic_items_df' not in st.session_state:
                            st.session_state['dynamic_items_df'] = pd.DataFrame(
                                columns=["Item Name", "Description / Specification", "Quantity", "Unit", "Remarks"])
                        new_row = pd.DataFrame([{"Item Name": hint, "Description / Specification": "",
                                                 "Quantity": 1, "Unit": "Nos", "Remarks": ""}])
                        st.session_state['dynamic_items_df'] = pd.concat(
                            [st.session_state['dynamic_items_df'], new_row], ignore_index=True)
                        st.rerun()

        st.markdown("---")
        st.markdown("##### 📋 Item List")
        if 'dynamic_items_df' not in st.session_state:
            st.session_state['dynamic_items_df'] = pd.DataFrame(
                [{"Item Name": "", "Description / Specification": "", "Quantity": 1, "Unit": "Nos", "Remarks": ""}])
        df_items_edit = st.session_state['dynamic_items_df'].copy()
        df_items_edit["Quantity"] = pd.to_numeric(df_items_edit.get("Quantity", 1), errors='coerce').fillna(1).astype(int)
        for col in ["Item Name", "Description / Specification", "Unit", "Remarks"]:
            if col not in df_items_edit.columns: df_items_edit[col] = ""
            df_items_edit[col] = df_items_edit[col].astype(str)

        edited_dynamic = st.data_editor(
            df_items_edit, num_rows="dynamic", use_container_width=True,
            column_config={
                "Item Name":                   st.column_config.TextColumn("Item Name", width="medium", required=True),
                "Description / Specification": st.column_config.TextColumn("Description / Specification", width="large"),
                "Quantity":                    st.column_config.NumberColumn("Qty", width="small", min_value=0, step=1, format="%d"),
                "Unit":                        st.column_config.SelectboxColumn("Unit", width="small", options=UNIT_OPTIONS, required=True),
                "Remarks":                     st.column_config.TextColumn("Remarks", width="medium"),
            }, key="dynamic_items_editor")
        st.session_state['dynamic_items_df'] = edited_dynamic
        n_items = len(edited_dynamic[edited_dynamic["Item Name"].astype(str).str.strip() != ""])
        if n_items > 0:
            st.success(f"✅ {n_items} item(s) added")
        else:
            st.warning("⚠️ Add at least one item to generate the RFQ.")

# Steps 5+
with st.form(key="rfq_form"):
    st.subheader("Step 5: Requirement Background")
    purpose = st.text_area(
        "Purpose of Requirement *",
        max_chars=2000,
        height=160,
        placeholder=(
            "Paste or type your requirement background here.\n\n"
            "Example:\n"
            "Pinnacle Mobility Solutions Pvt. Ltd. (PMSPL) is expanding its production "
            "capability in the component business with a new facility at Pithampur, MP. "
            "PMSPL seeks qualified vendors to design, manufacture, supply, and install a "
            "Carousel Racking System within the defined premises."
        )
    )

    with st.expander("📅 Timelines", expanded=True):
        today = date.today()
        c1, c2, c3 = st.columns(3)
        date_release  = c1.date_input("RFQ to Vendor *",                          today)
        date_query    = c1.date_input("1st Level Discussion *",                    today + timedelta(days=4))
        date_meet     = c1.date_input("Techno Commercial Offer (optional)",        None)
        date_quote    = c2.date_input("2nd Level Discussion on Proposal (opt.)",   None)
        date_selection= c2.date_input("Final Techno Commercial Offer *",           today + timedelta(days=15))
        date_review   = c2.date_input("PO to Vendor (optional)",                   None)
        date_delivery = c3.date_input("Delivery at Site *",                        today + timedelta(days=40))
        date_install  = c3.date_input("Installation at Site *",                    today + timedelta(days=47))

    with st.expander("👤 Single Point of Contact (SPOC)", expanded=True):
        st.markdown("##### Primary Contact *")
        p1, p2 = st.columns(2)
        spoc1_name        = p1.text_input("Name *",      key="s1n")
        spoc1_designation = p1.text_input("Designation", key="s1d")
        spoc1_phone       = p2.text_input("Phone No *",  key="s1p")
        spoc1_email       = p2.text_input("Email ID *",  key="s1e")
        st.markdown("##### Secondary Contact (Optional)")
        s1, s2 = st.columns(2)
        spoc2_name        = s1.text_input("Name",        key="s2n")
        spoc2_designation = s1.text_input("Designation", key="s2d")
        spoc2_phone       = s2.text_input("Phone No",    key="s2p")
        spoc2_email       = s2.text_input("Email ID",    key="s2e")

    with st.expander("📦 Submission, Delivery & Annexures", expanded=True):
        submit_to_name = st.text_input("Submit To (Company Name) *", "Agilomatrix Pvt. Ltd.")
        submit_to_registered_office = st.text_input(
            "Submit To (Registered Office Address)",
            "Registered Office: F1403, 7 Plumeria Drive, 7PD Street, Tathawade, Pune - 411033")
        st.markdown("**Delivery Location**")
        delivery_company = st.text_input("Delivery Company Name *", key="del_company",
                                         placeholder="e.g. EKA Mobility (Pinnacle Mobility Solutions Pvt. Ltd.)")
        delivery_gstin   = st.text_input("GSTIN No. (optional)", key="del_gstin",
                                         placeholder="e.g. 23AAFCI3261B1Z6")
        delivery_address = st.text_area("Delivery Address *", height=80, key="del_addr",
                                        placeholder="e.g. Plot no- A-3, Smart Industrial Township, Pithampur...")
        annexures = st.text_area("Annexures (one item per line)", height=80)

    submitted = st.form_submit_button("🚀 Generate RFQ Document", use_container_width=True, type="primary")


# ==============================================================
# PDF GENERATION TRIGGER
# ==============================================================
def _get_spec_df(prefix, section_name):
    import copy as _copy2
    fkey = f"frozen_{prefix}_{section_name}"
    dkey = f"data_{prefix}_{section_name}"
    wkey = f"widget_{prefix}_{section_name}"
    fallback = pd.DataFrame(_copy2.deepcopy(SPEC_TEMPLATE[section_name]))

    frozen = st.session_state.get(fkey)
    delta  = st.session_state.get(wkey)
    if frozen is not None and isinstance(delta, dict):
        df = frozen.copy()
        for row_idx_str, changes in delta.get('edited_rows', {}).items():
            row_idx = int(row_idx_str)
            if row_idx < len(df):
                for col, val in changes.items():
                    if col in df.columns:
                        df.at[row_idx, col] = val
        for new_row in delta.get('added_rows', []):
            row_data = {c: new_row.get(c, '') for c in df.columns}
            df = pd.concat([df, pd.DataFrame([row_data])], ignore_index=True)
        del_indices = sorted(delta.get('deleted_rows', []), reverse=True)
        for i in del_indices:
            if i < len(df):
                df = df.drop(df.index[i]).reset_index(drop=True)
        dkey_df = st.session_state.get(dkey)
        def _count_filled(d):
            if d is None or not isinstance(d, pd.DataFrame): return 0
            return int(d.astype(str).apply(lambda c: c.str.strip()).ne('').sum().sum())
        if _count_filled(df) >= _count_filled(dkey_df):
            return df
        return dkey_df if isinstance(dkey_df, pd.DataFrame) else fallback

    saved = st.session_state.get(dkey)
    if saved is not None and isinstance(saved, pd.DataFrame):
        return saved

    return fallback


def _get_custom_tables(prefix):
    """
    Reconstruct the list of custom tables for PDF generation.
    For each table index 0..n_tables-1, reads:
      - title  : from session_state custom_title_{prefix}_t{i}
      - columns: from session_state custom_colname_{prefix}_t{i}_{j}  (j = 0..n_cols-1)
      - df     : by applying widget delta onto frozen df
    Returns a list of dicts: [{'title': str, 'columns': [...], 'df': DataFrame}, ...]
    """
    n_tables = st.session_state.get(f"custom_n_tables_{prefix}", 1)
    result   = []

    def _apply_delta(frozen, delta, cols):
        """Apply Streamlit EditingState delta onto a frozen DataFrame."""
        if frozen is None:
            return pd.DataFrame()
        df = frozen.copy()
        if not isinstance(delta, dict):
            return df
        for row_idx_str, changes in delta.get('edited_rows', {}).items():
            row_idx = int(row_idx_str)
            if row_idx < len(df):
                for col, val in changes.items():
                    if col in df.columns:
                        df.at[row_idx, col] = val
        for new_row in delta.get('added_rows', []):
            row_data = {c: new_row.get(c, '') for c in df.columns}
            df = pd.concat([df, pd.DataFrame([row_data])], ignore_index=True)
        del_indices = sorted(delta.get('deleted_rows', []), reverse=True)
        for i in del_indices:
            if i < len(df):
                df = df.drop(df.index[i]).reset_index(drop=True)
        return df

    def _count_filled(d):
        if d is None or not isinstance(d, pd.DataFrame): return 0
        return int(d.astype(str).apply(lambda c: c.str.strip()).ne('').sum().sum())

    for tbl_idx in range(n_tables):
        tbl_pfx = f"{prefix}_t{tbl_idx}"

        # Title
        title = st.session_state.get(f"custom_title_{tbl_pfx}", "Technical Specification")

        # Column names
        n_cols = st.session_state.get(f"custom_ncols_{tbl_pfx}", 3)
        default_names = ["Parameter", "Value", "Unit", "Remarks", "Notes",
                         "Col 6", "Col 7", "Col 8", "Col 9", "Col 10"]
        cols = []
        for j in range(n_cols):
            ck = f"custom_colname_{tbl_pfx}_{j}"
            cols.append(st.session_state.get(ck, default_names[j] if j < len(default_names) else f"Col {j+1}"))

        # DataFrame — apply delta over frozen, fall back to saved dkey
        fkey = f"custom_frozen_{tbl_pfx}"
        dkey = f"custom_data_{tbl_pfx}"
        wkey = f"custom_widget_{tbl_pfx}"

        frozen  = st.session_state.get(fkey)
        delta   = st.session_state.get(wkey)
        df_live = _apply_delta(frozen, delta, cols)
        df_saved = st.session_state.get(dkey)

        # Use whichever has more data (handles both on_change and direct-generate paths)
        if _count_filled(df_live) >= _count_filled(df_saved):
            df_final = df_live
        else:
            df_final = df_saved if isinstance(df_saved, pd.DataFrame) else pd.DataFrame()

        result.append({
            'title':   title,
            'columns': cols,
            'df':      df_final,
        })

    return result


if submitted:
    current_category = st.session_state.get('rfq_category_select', rfq_category)
    current_wh_sub   = st.session_state.get('wh_sub_select', '') if st.session_state.get('rfq_category_select') == 'Warehouse Equipment' else ''
    is_wh = (current_category == "Warehouse Equipment")

    # Determine if custom table mode is active
    pfx_map_mode = {
        "Storage System":           "ss",
        "Material Handling":        "mh",
        "Dock Leveller":            "dl",
        "Automated Storage System": "carousel",
    }
    _mode_pfx       = pfx_map_mode.get(current_wh_sub, "ss")
    _table_mode_key = f"table_mode_{_mode_pfx}" if current_wh_sub in pfx_map_mode else None
    use_custom_spec = False
    if _table_mode_key:
        use_custom_spec = (st.session_state.get(_table_mode_key, "") == "✏️ Create My Own Table")

    errors = []
    if not Type_of_items.strip():     errors.append("Type of Items")
    if not Storage.strip():           errors.append("Storage Type")
    if not company_name.strip():      errors.append("Company Name")
    if not company_address.strip():   errors.append("Company Address")
    if not purpose.strip():           errors.append("Purpose of Requirement")
    if not spoc1_name.strip():        errors.append("SPOC Primary Name")
    if not spoc1_phone.strip():       errors.append("SPOC Primary Phone")
    if not spoc1_email.strip():       errors.append("SPOC Primary Email")
    if not submit_to_name.strip():    errors.append("Submit To Company Name")
    if not delivery_company.strip():  errors.append("Delivery Company Name")
    if not delivery_address.strip():  errors.append("Delivery Address")

    if not is_wh:
        items_df_check = st.session_state.get('dynamic_items_df', pd.DataFrame())
        if items_df_check.empty or items_df_check[items_df_check["Item Name"].astype(str).str.strip() != ""].empty:
            errors.append("At least one Item in the Item List")

    if errors:
        st.error("⚠️ Please fill in the following mandatory fields:\n" + "\n".join(f"  • {e}" for e in errors))
        st.stop()

    pdf_data_dict = {
        'rfq_category':   current_category,
        'wh_sub':         current_wh_sub,
        'use_custom_spec': use_custom_spec,
        'Type_of_items':  Type_of_items,
        'Storage':        Storage,
        'company_name':   company_name,
        'company_address': company_address,
        'footer_company_name':    footer_company_name,
        'footer_company_address': footer_company_address,
        'logo1_data': logo1_file.getvalue() if logo1_file else None,
        'logo1_w': logo1_w, 'logo1_h': logo1_h,
        'purpose': purpose,
        'date_release': date_release, 'date_query': date_query,
        'date_meet': date_meet, 'date_quote': date_quote,
        'date_selection': date_selection, 'date_delivery': date_delivery,
        'date_install': date_install, 'date_review': date_review,
        'spoc1_name': spoc1_name, 'spoc1_designation': spoc1_designation,
        'spoc1_phone': spoc1_phone, 'spoc1_email': spoc1_email,
        'spoc2_name': spoc2_name, 'spoc2_designation': spoc2_designation,
        'spoc2_phone': spoc2_phone, 'spoc2_email': spoc2_email,
        'submit_to_name': submit_to_name,
        'submit_to_registered_office': submit_to_registered_office,
        'delivery_company': delivery_company,
        'delivery_gstin':   delivery_gstin,
        'delivery_address': delivery_address,
        'annexures': annexures,
        'model_detail_header': st.session_state.get('model_detail_header_carousel', ''),
    }

    if is_wh:
        pfx_map = {
            "Storage Container":        "sc",
            "Automated Storage System": "carousel",
            "Storage System":           "ss",
            "Material Handling":        "mh",
            "Dock Leveller":            "dl",
        }
        layout_key = f"layout_images_{pfx_map.get(current_wh_sub, 'ss')}"
        pdf_data_dict['layout_images'] = st.session_state.get(layout_key, [])

        # Custom spec tables (overrides standard tables)
        if use_custom_spec:
            pdf_data_dict['custom_tables'] = _get_custom_tables(_mode_pfx)

        if current_wh_sub == "Storage Container":
            sc_images = st.session_state.get('storage_containers_images', {})
            sc_frozen = st.session_state.get('sc_frozen', pd.DataFrame())
            delta     = st.session_state.get('sc_wkey')

            if sc_frozen is not None and not sc_frozen.empty:
                sc_df = sc_frozen.copy()
                if isinstance(delta, dict):
                    for row_idx_str, changes in delta.get('edited_rows', {}).items():
                        row_idx = int(row_idx_str)
                        if row_idx < len(sc_df):
                            for col, val in changes.items():
                                if col in sc_df.columns:
                                    sc_df.at[row_idx, col] = val
                    for new_row in delta.get('added_rows', []):
                        row_data = {c: new_row.get(c, '') for c in sc_df.columns}
                        sc_df = pd.concat([sc_df, pd.DataFrame([row_data])], ignore_index=True)
                    del_indices = sorted(delta.get('deleted_rows', []), reverse=True)
                    for i in del_indices:
                        if i < len(sc_df):
                            sc_df = sc_df.drop(sc_df.index[i]).reset_index(drop=True)
                sc_saved = st.session_state.get('sc_data')
                if sc_saved is not None and isinstance(sc_saved, pd.DataFrame) and not sc_saved.empty:
                    def _filled_cells(df):
                        return df.astype(str).apply(lambda col: col.str.strip()).ne('').sum().sum()
                    if _filled_cells(sc_saved) >= _filled_cells(sc_df):
                        sc_df = sc_saved.copy()
            else:
                sc_df = st.session_state.get('sc_data', pd.DataFrame())

            if sc_df is not None and not sc_df.empty:
                sc_df = sc_df.reset_index(drop=True).copy()
                sc_df['image_data_bytes'] = [sc_images.get(i) for i in range(len(sc_df))]
                pdf_data_dict['storage_containers_df'] = sc_df
            else:
                pdf_data_dict['storage_containers_df'] = pd.DataFrame()
            pdf_data_dict['storage_containers_images'] = sc_images

        elif current_wh_sub == "Automated Storage System":
            pdf_data_dict['wh_items_df'] = st.session_state.get('wh_items_df', pd.DataFrame())
            if not use_custom_spec:
                for section_name in SPEC_TEMPLATE.keys():
                    val = _get_spec_df("carousel", section_name)
                    pdf_data_dict[{
                        "Model Details":               'carousel_model_df',
                        "Key Features":                'key_features_df',
                        "Inbuilt features":            'inbuilt_features_df',
                        "Installation Accountability": 'installation_df',
                    }[section_name]] = val

        else:
            pfx = {'Storage System': 'ss', 'Material Handling': 'mh', 'Dock Leveller': 'dl'}.get(current_wh_sub, 'ss')
            pdf_data_dict['model_detail_header'] = st.session_state.get(f'model_detail_header_{pfx}', '')
            if not use_custom_spec:
                for section_name in SPEC_TEMPLATE.keys():
                    pdf_data_dict[f"spec_{pfx}_{section_name}"] = _get_spec_df(pfx, section_name)
    else:
        pdf_data_dict['layout_images'] = []
        items_df = st.session_state.get('dynamic_items_df', pd.DataFrame())
        pdf_data_dict['items_df'] = items_df[items_df["Item Name"].astype(str).str.strip() != ""].reset_index(drop=True)

    with st.spinner("⚙️ Generating your RFQ PDF..."):
        try:
            pdf_bytes = create_advanced_rfq_pdf(pdf_data_dict)
            st.success("✅ RFQ PDF Generated Successfully!")
            fname = f"RFQ_{Type_of_items.replace(' ', '_')}_{date.today().strftime('%Y%m%d')}.pdf"
            st.download_button(
                "📥 Download RFQ Document",
                data=pdf_bytes, file_name=fname,
                mime="application/pdf",
                use_container_width=True, type="primary"
            )
        except Exception as e:
            st.error(f"❌ PDF generation failed: {e}")
            st.exception(e)
