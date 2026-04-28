# # app.py
# # ==========================================================
# # SIMPLE FINAL VERSION
# # ONLY WORK:
# # Upload Grid Report + Master Stock File
# # Copy ONLY NUMBERS from Grid -> Master Stock
# # Match Correct Shape
# # Match Correct Size Group
# # Do NOT touch formulas / totals / formatting
# # Download Updated Master Stock File
# # ==========================================================

# import streamlit as st
# from openpyxl import load_workbook
# from io import BytesIO

# # --------------------------------------------------
# # PAGE
# # --------------------------------------------------
# st.set_page_config(page_title="Grid To Master Stock", layout="wide")
# st.title("📊 Grid Report → Master Stock")

# # --------------------------------------------------
# # SHAPES
# # --------------------------------------------------
# SHAPES = [
#     "ROUND",
#     "OVAL",
#     "EMERALD",
#     "RADIANT",
#     "PEAR",
#     "PRINCESS",
#     "MARQUISE",
#     "CUSHION MODIFIED",
#     "CUSHION BRILLIANT",
#     "ASSCHER",
#     "HEART"
# ]

# # --------------------------------------------------
# # HELPERS
# # --------------------------------------------------
# def clean(v):
#     if v is None:
#         return ""
#     return str(v).strip().upper()

# def is_number(v):
#     try:
#         float(v)
#         return True
#     except:
#         return False

# def find_shape_rows(ws):
#     found = {}

#     for r in range(1, ws.max_row + 1):
#         val = clean(ws.cell(r, 1).value)

#         if val in SHAPES and val not in found:
#             found[val] = r

#     return found

# def find_total_row(ws, start_row):
#     for r in range(start_row, ws.max_row + 1):
#         for c in range(1, 10):
#             if clean(ws.cell(r, c).value) == "TOTAL":
#                 return r
#     return start_row + 15

# # --------------------------------------------------
# # MAIN COPY LOGIC
# # --------------------------------------------------
# def copy_grid_to_master(ws_grid, ws_master):

#     grid_shapes = find_shape_rows(ws_grid)
#     master_shapes = find_shape_rows(ws_master)

#     updated = []

#     for shape in SHAPES:

#         if shape in grid_shapes and shape in master_shapes:

#             src_start = grid_shapes[shape]
#             src_end   = find_total_row(ws_grid, src_start)

#             tgt_start = master_shapes[shape]

#             rows = src_end - src_start

#             for i in range(rows):

#                 src_row = src_start + i
#                 tgt_row = tgt_start + i

#                 for col in range(1, ws_grid.max_column + 1):

#                     val = ws_grid.cell(src_row, col).value

#                     # only numeric copy
#                     if is_number(val):
#                         ws_master.cell(tgt_row, col).value = val

#             updated.append(shape)

#     return updated

# # --------------------------------------------------
# # PROCESS
# # --------------------------------------------------
# def process(grid_file, master_file):

#     wb_grid = load_workbook(grid_file, data_only=True)
#     wb_master = load_workbook(master_file)

#     ws_grid = wb_grid.active
#     ws_master = wb_master.active

#     updated = copy_grid_to_master(ws_grid, ws_master)

#     output = BytesIO()
#     wb_master.save(output)
#     output.seek(0)

#     return output, updated

# # --------------------------------------------------
# # UI
# # --------------------------------------------------
# grid_file = st.file_uploader("Upload Grid Report", type=["xlsx"])
# master_file = st.file_uploader("Upload Master Stock File", type=["xlsx"])

# if grid_file and master_file:

#     if st.button("🚀 Process"):

#         try:
#             with st.spinner("Working..."):

#                 output, updated = process(grid_file, master_file)

#             st.success("✅ Completed Successfully")

#             st.write("### Updated Shapes")
#             st.write(", ".join(updated))

#             st.download_button(
#                 "📥 Download Updated Master Stock File",
#                 data=output,
#                 file_name="Updated_Master_Stock.xlsx",
#                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#             )

#         except Exception as e:
#             st.error(str(e))


# app.py
# ==========================================================
# FIXED VERSION
# Copy/Paste VALUES ONLY
# No formulas inserted
# No #REF issue
# Download all 3 files
# ==========================================================

# import streamlit as st
# from openpyxl import load_workbook
# from io import BytesIO

# st.set_page_config(page_title="Daily Work Automation", layout="wide")
# st.title("📊 Full Daily Work Automation")

# # --------------------------------------------------
# # SESSION
# # --------------------------------------------------
# if "processed" not in st.session_state:
#     st.session_state.processed = False

# # --------------------------------------------------
# # SHAPES
# # --------------------------------------------------
# SHAPES = [
#     "ROUND","OVAL","EMERALD","RADIANT","PEAR",
#     "PRINCESS","MARQUISE","CUSHION MODIFIED",
#     "CUSHION BRILLIANT","ASSCHER","HEART"
# ]

# # --------------------------------------------------
# # HELPERS
# # --------------------------------------------------
# def clean(v):
#     if v is None:
#         return ""
#     return str(v).strip().upper()

# def is_number(v):
#     try:
#         float(v)
#         return True
#     except:
#         return False

# def find_header(ws, text):
#     for row in ws.iter_rows():
#         for cell in row:
#             if clean(cell.value) == clean(text):
#                 return cell.row, cell.column
#     return None, None

# def find_shape_rows(ws):
#     found = {}

#     for r in range(1, ws.max_row + 1):
#         val = clean(ws.cell(r,1).value)

#         if val in SHAPES and val not in found:
#             found[val] = r

#     return found

# def find_total_row(ws, start):
#     for r in range(start, ws.max_row + 1):
#         for c in range(1,10):
#             if clean(ws.cell(r,c).value) == "TOTAL":
#                 return r
#     return start + 15

# # --------------------------------------------------
# # STEP 1 GRID -> MASTER STOCK
# # --------------------------------------------------
# def copy_grid_to_stock(ws_grid, ws_stock):

#     g = find_shape_rows(ws_grid)
#     s = find_shape_rows(ws_stock)

#     for shape in SHAPES:

#         if shape in g and shape in s:

#             gs = g[shape]
#             ge = find_total_row(ws_grid, gs)

#             ss = s[shape]

#             rows = ge - gs

#             for i in range(rows):

#                 for col in range(1, ws_grid.max_column + 1):

#                     val = ws_grid.cell(gs+i, col).value

#                     if is_number(val):
#                         ws_stock.cell(ss+i, col).value = val

# # --------------------------------------------------
# # COPY VALUE ONLY
# # --------------------------------------------------
# def copy_value_column(ws_source, source_header, ws_target, target_header):

#     sr, sc = find_header(ws_source, source_header)
#     tr, tc = find_header(ws_target, target_header)

#     if not sr or not tr:
#         return

#     r = sr + 1
#     t = tr + 1

#     while r <= ws_source.max_row:

#         val = ws_source.cell(r, sc).value

#         if r > sr + 1000:
#             break

#         # IMPORTANT: copy only final values
#         if val is None or clean(val) == "":
#             ws_target.cell(t, tc).value = 0
#         else:
#             ws_target.cell(t, tc).value = val

#         r += 1
#         t += 1

# # --------------------------------------------------
# # PROCESS
# # --------------------------------------------------
# def process(grid_file, stock_file, master_file):

#     # normal files
#     wb_grid = load_workbook(grid_file)
#     wb_stock = load_workbook(stock_file)
#     wb_master = load_workbook(master_file)

#     ws_grid = wb_grid.active
#     ws_stock = wb_stock.active
#     ws_master = wb_master.active

#     # ------------------------------------------------
#     # STEP 1
#     # Grid -> Master Stock
#     # ------------------------------------------------
#     copy_grid_to_stock(ws_grid, ws_stock)

#     # Save updated stock first
#     temp = BytesIO()
#     wb_stock.save(temp)
#     temp.seek(0)

#     # ------------------------------------------------
#     # IMPORTANT FIX
#     # Load SAME updated file values
#     # ------------------------------------------------
#     wb_stock_values = load_workbook(temp, data_only=False)

#     # Sheet2
#     ws_sheet2 = wb_stock_values.worksheets[1]

#     # ------------------------------------------------
#     # STEP 2
#     # INHAND -> Available
#     # MAX PCS -> Grid
#     # ------------------------------------------------
#     copy_value_column(ws_sheet2, "INHAND", ws_master, "AVAILABLE")
#     copy_value_column(ws_sheet2, "MAX PCS", ws_grid, "GRID")

#     # ------------------------------------------------
#     # SAVE
#     # ------------------------------------------------
#     out1 = BytesIO()
#     out2 = BytesIO()
#     out3 = BytesIO()

#     wb_grid.save(out1)
#     wb_stock.save(out2)
#     wb_master.save(out3)

#     out1.seek(0)
#     out2.seek(0)
#     out3.seek(0)

#     return out1, out2, out3

# # --------------------------------------------------
# # UI
# # --------------------------------------------------
# grid_file = st.file_uploader("Upload Grid Report", type=["xlsx"])
# stock_file = st.file_uploader("Upload Master Stock File", type=["xlsx"])
# master_file = st.file_uploader("Upload Master File", type=["xlsx"])

# if grid_file and stock_file and master_file:

#     if st.button("🚀 Process All Files"):

#         try:
#             with st.spinner("Processing..."):

#                 g, s, m = process(grid_file, stock_file, master_file)

#                 st.session_state.grid = g.getvalue()
#                 st.session_state.stock = s.getvalue()
#                 st.session_state.master = m.getvalue()

#                 st.session_state.processed = True

#             st.success("✅ Completed Successfully")

#         except Exception as e:
#             st.error(str(e))

# # --------------------------------------------------
# # DOWNLOADS
# # --------------------------------------------------
# if st.session_state.processed:

#     st.download_button(
#         "📥 Download Grid Report",
#         data=st.session_state.grid,
#         file_name="Updated_Grid_Report.xlsx"
#     )

#     st.download_button(
#         "📥 Download Master Stock",
#         data=st.session_state.stock,
#         file_name="Updated_Master_Stock.xlsx"
#     )

#     st.download_button(
#         "📥 Download Master File",
#         data=st.session_state.master,
#         file_name="Updated_Master_File.xlsx"
#     )



import streamlit as st
from openpyxl import load_workbook
from io import BytesIO

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="DiamondFlow – Daily Work Automation",
    page_icon="💎",
    layout="wide",
)

# ─────────────────────────────────────────────
# CUSTOM CSS  (professional dark-blue theme)
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Syne:wght@700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

/* ── background ── */
.stApp {
    background: linear-gradient(135deg, #0f1b2d 0%, #142035 60%, #0d2444 100%);
    min-height: 100vh;
}

/* ── top hero ── */
.hero-wrap {
    background: linear-gradient(90deg, #0d2444 0%, #163569 50%, #0d2444 100%);
    border-bottom: 1px solid rgba(56,189,248,0.25);
    padding: 2rem 2.5rem 1.5rem;
    margin: -1rem -1rem 2rem -1rem;
}
.hero-title {
    font-family: 'Syne', sans-serif;
    font-size: 2.6rem;
    font-weight: 800;
    letter-spacing: -0.5px;
    color: #e0f2fe;
    margin: 0 0 0.25rem 0;
}
.hero-title span { color: #38bdf8; }
.hero-sub {
    font-size: 0.95rem;
    color: #7dd3fc;
    margin: 0;
    letter-spacing: 0.03em;
}
.badge {
    display: inline-block;
    background: rgba(56,189,248,0.15);
    border: 1px solid rgba(56,189,248,0.4);
    color: #38bdf8;
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    padding: 3px 10px;
    border-radius: 20px;
    margin-bottom: 0.75rem;
}

/* ── upload section ── */
.section-label {
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: #7dd3fc;
    margin-bottom: 0.4rem;
}
.upload-card {
    background: rgba(255,255,255,0.04);
    border: 1px solid rgba(56,189,248,0.2);
    border-radius: 14px;
    padding: 1.25rem 1.5rem;
    margin-bottom: 1rem;
    transition: border-color 0.2s;
}
.upload-card:hover { border-color: rgba(56,189,248,0.5); }

/* ── step flow ── */
.flow-wrap {
    display: flex;
    align-items: flex-start;
    gap: 0;
    margin: 2rem 0 1.5rem;
}
.flow-step {
    flex: 1;
    text-align: center;
    position: relative;
}
.flow-step:not(:last-child)::after {
    content: '';
    position: absolute;
    top: 20px;
    right: -1px;
    width: 50%;
    height: 2px;
    background: linear-gradient(90deg, #38bdf8, rgba(56,189,248,0.15));
}
.flow-step:not(:first-child)::before {
    content: '';
    position: absolute;
    top: 20px;
    left: -1px;
    width: 50%;
    height: 2px;
    background: linear-gradient(90deg, rgba(56,189,248,0.15), #38bdf8);
}
.step-dot {
    width: 40px; height: 40px;
    border-radius: 50%;
    background: rgba(56,189,248,0.1);
    border: 2px solid #38bdf8;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    font-weight: 700;
    font-size: 0.85rem;
    color: #38bdf8;
    margin-bottom: 0.5rem;
}
.step-label {
    font-size: 0.75rem;
    color: #94a3b8;
    line-height: 1.3;
}

/* ── process button ── */
.stButton > button {
    width: 100%;
    background: linear-gradient(135deg, #0ea5e9 0%, #2563eb 100%);
    color: white;
    border: none;
    border-radius: 12px;
    padding: 0.85rem 2rem;
    font-size: 1rem;
    font-weight: 600;
    letter-spacing: 0.02em;
    cursor: pointer;
    transition: opacity 0.2s, transform 0.1s;
    box-shadow: 0 4px 24px rgba(14,165,233,0.35);
}
.stButton > button:hover { opacity: 0.92; transform: translateY(-1px); }
.stButton > button:active { transform: scale(0.98); }

/* ── file uploader ── */
[data-testid="stFileUploader"] {
    background: rgba(255,255,255,0.03) !important;
    border: 1.5px dashed rgba(56,189,248,0.3) !important;
    border-radius: 10px !important;
}
[data-testid="stFileUploader"] label { color: #cbd5e1 !important; }

/* ── success / error ── */
[data-testid="stAlert"] {
    border-radius: 10px !important;
    font-weight: 500;
}

/* ── download buttons ── */
[data-testid="stDownloadButton"] > button {
    background: rgba(56,189,248,0.08) !important;
    border: 1px solid rgba(56,189,248,0.4) !important;
    color: #7dd3fc !important;
    border-radius: 10px !important;
    font-weight: 500;
    width: 100%;
    transition: background 0.2s;
}
[data-testid="stDownloadButton"] > button:hover {
    background: rgba(56,189,248,0.2) !important;
}

/* ── metrics ── */
[data-testid="metric-container"] {
    background: rgba(255,255,255,0.04);
    border: 1px solid rgba(56,189,248,0.15);
    border-radius: 12px;
    padding: 1rem;
}
[data-testid="stMetricLabel"] { color: #94a3b8 !important; }
[data-testid="stMetricValue"] { color: #e0f2fe !important; }

/* ── misc text colors ── */
h1,h2,h3,h4 { color: #e0f2fe !important; }
p, li { color: #cbd5e1; }
.stMarkdown p { color: #cbd5e1; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────
for key in ["processed", "grid", "stock", "master", "log"]:
    if key not in st.session_state:
        st.session_state[key] = False if key == "processed" else None

# ─────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────
SHAPES = [
    "ROUND", "OVAL", "EMERALD", "RADIANT", "PEAR",
    "PRINCESS", "MARQUISE", "CUSHION MODIFIED",
    "CUSHION BRILLIANT", "ASSCHER", "HEART",
]

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def clean(v):
    if v is None:
        return ""
    return str(v).strip().upper()

def is_number(v):
    try:
        float(v)
        return True
    except Exception:
        return False

def find_header(ws, text, search_rows=10):
    """Find header cell by scanning first `search_rows` rows across ALL columns."""
    target = clean(text)
    for row in ws.iter_rows(max_row=search_rows):
        for cell in row:
            if clean(cell.value) == target:
                return cell.row, cell.column
    return None, None

def find_shape_rows(ws):
    """Return {SHAPE: first_row} scanning column A."""
    found = {}
    for r in range(1, ws.max_row + 1):
        val = clean(ws.cell(r, 1).value)
        if val in SHAPES and val not in found:
            found[val] = r
    return found

def find_total_row(ws, start):
    """Row index of the TOTAL row after `start`."""
    for r in range(start, min(ws.max_row + 1, start + 300)):
        for c in range(1, 10):
            if clean(ws.cell(r, c).value) == "TOTAL":
                return r
    return start + 50

# ─────────────────────────────────────────────
# STEP 1 – Grid Report → Master Stock (Sheet1)
# Copy only numeric cells, shape by shape
# ─────────────────────────────────────────────
def copy_grid_to_stock(ws_grid, ws_stock, log):
    g = find_shape_rows(ws_grid)
    s = find_shape_rows(ws_stock)
    copied_shapes = []

    for shape in SHAPES:
        if shape not in g:
            log.append(f"⚠️  Shape '{shape}' not found in Grid Report – skipped.")
            continue
        if shape not in s:
            log.append(f"⚠️  Shape '{shape}' not found in Master Stock – skipped.")
            continue

        gs = g[shape]
        ge = find_total_row(ws_grid, gs)
        ss = s[shape]
        rows = ge - gs

        for i in range(rows):
            for col in range(1, ws_grid.max_column + 1):
                val = ws_grid.cell(gs + i, col).value
                if is_number(val):
                    ws_stock.cell(ss + i, col).value = float(val)

        copied_shapes.append(shape)

    log.append(f"✅ Grid → Master Stock: copied {len(copied_shapes)} shapes "
               f"({', '.join(copied_shapes) if copied_shapes else 'none'}).")

# ─────────────────────────────────────────────
# KEY-MATCH COPY: Sheet2 → Master File
#
# Master Stock Sheet2 columns:
#   SHAPE | SIZE | FROM SIZE | TO SIZE | COLOR | CLARITY | MAX PCS | INHAND | TO ORDER
#
# Master File Sheet1 columns:
#   Shape | From Size | To Size | Clarity | Color | Grid | Available | ...
#
# Strategy: build a lookup dict from Sheet2 keyed by
#   (SHAPE, FROM_SIZE, TO_SIZE, CLARITY, COLOR)
# then walk every data row in the Master File and fill
#   Grid ← MAX PCS
#   Available ← INHAND
# ─────────────────────────────────────────────

def _safe_num(v):
    """Return int value of v, or 0 if missing/formula."""
    if v is None:
        return 0
    if isinstance(v, str):
        v = v.strip()
        if v == "" or v.startswith("="):
            return 0
        try:
            return int(float(v))
        except Exception:
            return 0
    try:
        return int(float(v))
    except Exception:
        return 0

def _key_str(v):
    """Normalise a key component to uppercase string, rounding floats."""
    if v is None:
        return ""
    if isinstance(v, float):
        return f"{v:.2f}"
    return str(v).strip().upper()

def build_sheet2_lookup(ws_sheet2, log):
    """
    Returns a dict: (shape, from_size, to_size, clarity, color) -> {MAX PCS, INHAND}
    Column positions are discovered from the header row.
    """
    # find header row (scan first 5 rows)
    header_row = None
    for r in range(1, 6):
        row_vals = [clean(ws_sheet2.cell(r, c).value) for c in range(1, 15)]
        if any(v in ("SHAPE", "MAX PCS", "INHAND") for v in row_vals):
            header_row = r
            break

    if header_row is None:
        log.append("❌ Sheet2: could not find header row – lookup aborted.")
        return {}

    # map column name → column index
    col_map = {}
    for c in range(1, ws_sheet2.max_column + 1):
        h = clean(ws_sheet2.cell(header_row, c).value)
        if h:
            col_map[h] = c

    log.append(f"  Sheet2 headers found: {col_map}")

    needed = ["SHAPE", "FROM SIZE", "TO SIZE", "CLARITY", "COLOR", "MAX PCS", "INHAND"]
    missing = [n for n in needed if n not in col_map]
    if missing:
        log.append(f"⚠️  Sheet2 missing columns: {missing} — will try partial match.")

    lookup = {}
    for r in range(header_row + 1, ws_sheet2.max_row + 1):
        shape    = _key_str(ws_sheet2.cell(r, col_map.get("SHAPE",    0)).value) if "SHAPE"     in col_map else ""
        from_sz  = _key_str(ws_sheet2.cell(r, col_map.get("FROM SIZE",0)).value) if "FROM SIZE" in col_map else ""
        to_sz    = _key_str(ws_sheet2.cell(r, col_map.get("TO SIZE",  0)).value) if "TO SIZE"   in col_map else ""
        clarity  = _key_str(ws_sheet2.cell(r, col_map.get("CLARITY",  0)).value) if "CLARITY"   in col_map else ""
        color    = _key_str(ws_sheet2.cell(r, col_map.get("COLOR",    0)).value) if "COLOR"     in col_map else ""
        max_pcs  = _safe_num(ws_sheet2.cell(r, col_map.get("MAX PCS", 0)).value) if "MAX PCS"   in col_map else 0
        inhand   = _safe_num(ws_sheet2.cell(r, col_map.get("INHAND",  0)).value) if "INHAND"    in col_map else 0

        if not shape:   # empty row — skip
            continue

        key = (shape, from_sz, to_sz, clarity, color)
        lookup[key] = {"MAX PCS": max_pcs, "INHAND": inhand}

    log.append(f"  Sheet2 lookup built: {len(lookup)} rows indexed.")
    return lookup


def fill_master_from_lookup(ws_master, lookup, log):
    """
    Walk every data row in Master File Sheet1.
    Match by (Shape, From Size, To Size, Clarity, Color).
    Write MAX PCS → Grid column, INHAND → Available column.
    Clears existing formula before writing (prevents #REF!).
    """
    # discover header row and column positions in Master File
    header_row = None
    for r in range(1, 6):
        row_vals = [clean(ws_master.cell(r, c).value) for c in range(1, 15)]
        if any(v in ("SHAPE", "GRID", "AVAILABLE") for v in row_vals):
            header_row = r
            break

    if header_row is None:
        log.append("❌ Master File: could not find header row.")
        return

    col_map = {}
    for c in range(1, ws_master.max_column + 1):
        h = clean(ws_master.cell(header_row, c).value)
        if h:
            col_map[h] = c

    log.append(f"  Master File headers: {col_map}")

    # column indices we need
    c_shape    = col_map.get("SHAPE")
    c_from     = col_map.get("FROM SIZE")
    c_to       = col_map.get("TO SIZE")
    c_clarity  = col_map.get("CLARITY")
    c_color    = col_map.get("COLOR")
    c_grid     = col_map.get("GRID")
    c_avail    = col_map.get("AVAILABLE")

    missing = [n for n, c in [("SHAPE",c_shape),("FROM SIZE",c_from),("TO SIZE",c_to),
                                ("CLARITY",c_clarity),("COLOR",c_color),
                                ("GRID",c_grid),("AVAILABLE",c_avail)] if c is None]
    if missing:
        log.append(f"⚠️  Master File missing columns: {missing}")

    hit = 0; miss = 0
    for r in range(header_row + 1, ws_master.max_row + 1):
        shape   = _key_str(ws_master.cell(r, c_shape).value)   if c_shape   else ""
        from_sz = _key_str(ws_master.cell(r, c_from).value)    if c_from    else ""
        to_sz   = _key_str(ws_master.cell(r, c_to).value)      if c_to      else ""
        clarity = _key_str(ws_master.cell(r, c_clarity).value) if c_clarity else ""
        color   = _key_str(ws_master.cell(r, c_color).value)   if c_color   else ""

        if not shape:
            continue

        key = (shape, from_sz, to_sz, clarity, color)
        data = lookup.get(key)

        def write_cell(col, val):
            if col is None:
                return
            cell = ws_master.cell(r, col)
            cell.value = None        # wipe formula → no #REF!
            cell.data_type = "n"
            cell.value = val

        if data:
            write_cell(c_grid,  data["MAX PCS"])
            write_cell(c_avail, data["INHAND"])
            hit += 1
        else:
            # key not in Sheet2 — write 0 (and still clear formula)
            write_cell(c_grid,  0)
            write_cell(c_avail, 0)
            miss += 1

    log.append(f"✅ Master File updated: {hit} rows matched, {miss} rows set to 0 (no Sheet2 match).")

# ─────────────────────────────────────────────
# MAIN PROCESS
# ─────────────────────────────────────────────
def process(grid_file, stock_file, master_file):
    log = []

    # ── Read all file bytes upfront (BytesIO objects get consumed on first read) ──
    grid_bytes   = grid_file.read()
    stock_bytes  = stock_file.read()
    master_bytes = master_file.read()

    # ── Load for READING computed values (data_only=True) ──
    wb_grid_read   = load_workbook(BytesIO(grid_bytes),   data_only=True)
    wb_stock_read  = load_workbook(BytesIO(stock_bytes),  data_only=True)
    wb_master_read = load_workbook(BytesIO(master_bytes), data_only=True)

    # ── Load for WRITING (data_only=False keeps formulas editable) ──
    # We will write INTO these, then save them as outputs
    wb_stock_write  = load_workbook(BytesIO(stock_bytes),  data_only=False)
    wb_master_write = load_workbook(BytesIO(master_bytes), data_only=False)
    wb_grid_write   = load_workbook(BytesIO(grid_bytes),   data_only=False)

    ws_grid_read    = wb_grid_read.active
    ws_stock_read   = wb_stock_read.worksheets[0]    # Sheet1 – source for Step 1
    ws_stock_write  = wb_stock_write.worksheets[0]   # Sheet1 – target for Step 1

    # ─────────────────────────────────────
    # STEP 1 – Grid Report → Master Stock Sheet1
    # Read numbers from Grid (data_only), write into Stock Sheet1 (write wb)
    # ─────────────────────────────────────
    copy_grid_to_stock(ws_grid_read, ws_stock_write, log)

    # ─────────────────────────────────────
    # STEP 2 & 3 – Master Stock Sheet2 → Master File
    # MAX PCS → Grid column
    # INHAND  → Available column
    # Matched by (Shape, From Size, To Size, Clarity, Color) key
    # ─────────────────────────────────────
    if len(wb_stock_read.worksheets) < 2:
        log.append("❌ Master Stock file does not have a Sheet2 – steps 2 & 3 skipped.")
    else:
        ws_sheet2       = wb_stock_read.worksheets[1]    # data_only=True → real values
        ws_master_write = wb_master_write.worksheets[0]

        log.append("─── Building Sheet2 lookup ───")
        lookup = build_sheet2_lookup(ws_sheet2, log)

        log.append("─── Filling Master File (Grid + Available) ───")
        fill_master_from_lookup(ws_master_write, lookup, log)

    # ── Save all three outputs ──
    out_grid   = BytesIO()
    out_stock  = BytesIO()
    out_master = BytesIO()

    wb_grid_write.save(out_grid)      # Grid Report (Step 1 numbers written)
    wb_stock_write.save(out_stock)    # Master Stock Sheet1 (Step 1 target)
    wb_master_write.save(out_master)  # Master File (Grid + Available filled)

    out_grid.seek(0);  out_stock.seek(0);  out_master.seek(0)

    return out_grid, out_stock, out_master, log


# ─────────────────────────────────────────────
# UI – HERO
# ─────────────────────────────────────────────
st.markdown("""
<div class="hero-wrap">
    <div class="badge">💎 Diamond Inventory Suite</div>
    <h1 class="hero-title">Diamond<span>Flow</span></h1>
    <p class="hero-sub">Daily Work Automation · Grid → Stock → Master · Built for speed &amp; accuracy</p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# UI – STEP FLOW DIAGRAM
# ─────────────────────────────────────────────
st.markdown("""
<div class="flow-wrap">
  <div class="flow-step">
    <div class="step-dot">1</div>
    <div class="step-label">Upload<br/>Grid Report</div>
  </div>
  <div class="flow-step">
    <div class="step-dot">2</div>
    <div class="step-label">Upload<br/>Master Stock</div>
  </div>
  <div class="flow-step">
    <div class="step-dot">3</div>
    <div class="step-label">Upload<br/>Master File</div>
  </div>
  <div class="flow-step">
    <div class="step-dot">4</div>
    <div class="step-label">Process<br/>All Files</div>
  </div>
  <div class="flow-step">
    <div class="step-dot">5</div>
    <div class="step-label">Download<br/>Results</div>
  </div>
</div>
""", unsafe_allow_html=True)

st.divider()

# ─────────────────────────────────────────────
# UI – UPLOAD COLUMNS
# ─────────────────────────────────────────────
col1, col2, col3 = st.columns(3)

with col1:
    st.markdown('<div class="section-label">📋 Step 1 – Grid Report</div>', unsafe_allow_html=True)
    grid_file = st.file_uploader(
        "Grid Report (.xlsx)",
        type=["xlsx"],
        key="grid_upload",
        label_visibility="collapsed",
    )
    if grid_file:
        st.success(f"✓ {grid_file.name}")

with col2:
    st.markdown('<div class="section-label">📦 Step 2 – Master Stock File</div>', unsafe_allow_html=True)
    stock_file = st.file_uploader(
        "Master Stock (.xlsx)",
        type=["xlsx"],
        key="stock_upload",
        label_visibility="collapsed",
    )
    if stock_file:
        st.success(f"✓ {stock_file.name}")

with col3:
    st.markdown('<div class="section-label">📁 Step 3 – Master File</div>', unsafe_allow_html=True)
    master_file = st.file_uploader(
        "Master File (.xlsx)",
        type=["xlsx"],
        key="master_upload",
        label_visibility="collapsed",
    )
    if master_file:
        st.success(f"✓ {master_file.name}")

st.markdown("<br/>", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# UI – PROCESS BUTTON
# ─────────────────────────────────────────────
all_ready = grid_file and stock_file and master_file

if all_ready:
    if st.button("🚀 Process All Files", use_container_width=True):
        try:
            with st.spinner("Processing files… please wait"):
                g, s, m, log = process(grid_file, stock_file, master_file)

                st.session_state.grid      = g.getvalue()
                st.session_state.stock     = s.getvalue()
                st.session_state.master    = m.getvalue()
                st.session_state.log       = log
                st.session_state.processed = True

            st.success("✅ All files processed successfully!")

        except Exception as e:
            st.error(f"❌ Error: {e}")
            import traceback
            st.code(traceback.format_exc())
else:
    st.info("ℹ️ Please upload all three files above to enable processing.")

# ─────────────────────────────────────────────
# UI – RESULTS
# ─────────────────────────────────────────────
if st.session_state.processed:
    st.divider()
    st.markdown("### 📥 Download Updated Files")

    d1, d2, d3 = st.columns(3)

    with d1:
        st.download_button(
            "📋 Grid Report (Updated)",
            data=st.session_state.grid,
            file_name="Updated_Grid_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with d2:
        st.download_button(
            "📦 Master Stock (Updated)",
            data=st.session_state.stock,
            file_name="Updated_Master_Stock.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with d3:
        st.download_button(
            "📁 Master File (Updated)",
            data=st.session_state.master,
            file_name="Updated_Master_File.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    # ── Processing log ──
    if st.session_state.log:
        with st.expander("🔍 View Processing Log", expanded=False):
            for line in st.session_state.log:
                if line.startswith("✅"):
                    st.success(line)
                elif line.startswith("⚠️"):
                    st.warning(line)
                elif line.startswith("❌"):
                    st.error(line)
                else:
                    st.text(line)

# ─────────────────────────────────────────────
# UI – FOOTER
# ─────────────────────────────────────────────
st.divider()
st.markdown("""
<div style="text-align:center; color: #475569; font-size: 0.8rem; padding: 0.5rem 0 1rem;">
    DiamondFlow · Daily Automation Tool · Shapes: ROUND · OVAL · EMERALD · RADIANT · PEAR · PRINCESS · MARQUISE · CUSHION · ASSCHER · HEART
</div>
""", unsafe_allow_html=True)