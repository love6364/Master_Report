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


# # app.py
# # ==========================================================
# # FIXED VERSION
# # Copy/Paste VALUES ONLY
# # No formulas inserted
# # No #REF issue
# # Download all 3 files
# # ==========================================================

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


# app.py - Daily Work Automation
import streamlit as st
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="Diamond Work Automation", layout="wide", page_icon="💎")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono&display=swap');
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.stApp { background: #0f1117; color: #e8eaf0; }
header[data-testid="stHeader"] { background: transparent; }
.upload-label { font-size: 0.78rem; font-weight: 500; letter-spacing: 1.2px; text-transform: uppercase; color: #4f6ef7; margin-bottom: 0.5rem; }
.steps-panel { background: #1a1d27; border: 1px solid #2a2d3a; border-radius: 14px; padding: 1.4rem 1.6rem; }
.steps-panel h4 { color: #7c8394; font-size: 0.72rem; letter-spacing: 1.4px; text-transform: uppercase; margin-bottom: 1rem; font-weight: 500; }
.step-row { display: flex; align-items: flex-start; gap: 0.9rem; margin-bottom: 0.85rem; }
.step-num { background: #4f6ef7; color: white; border-radius: 50%; width: 22px; height: 22px; font-size: 0.7rem; font-weight: 600; display: flex; align-items: center; justify-content: center; flex-shrink: 0; margin-top: 1px; }
.step-text { font-size: 0.88rem; color: #c0c4d0; line-height: 1.5; }
.step-text span { color: #ffffff; font-weight: 500; }
.step-arrow { color: #4f6ef7; font-size: 0.8rem; }
div[data-testid="stButton"] > button { background: linear-gradient(135deg, #4f6ef7 0%, #7c4ff7 100%); color: white; border: none; border-radius: 10px; padding: 0.75rem 2rem; font-family: 'DM Sans', sans-serif; font-size: 0.95rem; font-weight: 500; letter-spacing: 0.3px; width: 100%; cursor: pointer; box-shadow: 0 4px 20px rgba(79, 110, 247, 0.3); transition: opacity 0.2s, transform 0.1s; }
div[data-testid="stButton"] > button:hover { opacity: 0.9; transform: translateY(-1px); }
div[data-testid="stDownloadButton"] > button { background: #1a1d27; color: #e8eaf0; border: 1px solid #2a2d3a; border-radius: 10px; padding: 0.65rem 1.2rem; font-family: 'DM Sans', sans-serif; font-size: 0.88rem; font-weight: 400; width: 100%; transition: border-color 0.2s, background 0.2s; }
div[data-testid="stDownloadButton"] > button:hover { border-color: #4f6ef7; background: #1f2235; color: #ffffff; }
div[data-testid="stFileUploader"] section { border: 1.5px dashed #2a2d3a; border-radius: 10px; background: #12151e; padding: 0.8rem; transition: border-color 0.2s; }
div[data-testid="stFileUploader"] section:hover { border-color: #4f6ef7; }
hr { border-color: #2a2d3a; margin: 1.5rem 0; }
.dl-header { font-size: 0.75rem; font-weight: 500; letter-spacing: 1.3px; text-transform: uppercase; color: #7c8394; margin-bottom: 0.8rem; margin-top: 1.2rem; }
.hero { text-align: center; padding: 2.5rem 0 1.5rem 0; }
.hero h1 { font-size: 2.4rem; font-weight: 600; letter-spacing: -0.5px; color: #ffffff; margin-bottom: 0.3rem; }
.hero p { color: #7c8394; font-size: 1rem; font-weight: 300; }
</style>
""", unsafe_allow_html=True)

# --------------------------------------------------
# SESSION
# --------------------------------------------------
if "processed" not in st.session_state:
    st.session_state.processed = False

# --------------------------------------------------
# SHAPES
# --------------------------------------------------
SHAPES = [
    "ROUND","OVAL","EMERALD","RADIANT","PEAR",
    "PRINCESS","MARQUISE","CUSHION MODIFIED",
    "CUSHION BRILLIANT","ASSCHER","HEART","SQ RADIANT"
]

# --------------------------------------------------
# HELPERS
# --------------------------------------------------
def clean(v):
    if v is None:
        return ""
    return str(v).strip().upper()

def is_number(v):
    try:
        float(v)
        return True
    except:
        return False

def find_header(ws, text):
    for row in ws.iter_rows():
        for cell in row:
            if clean(cell.value) == clean(text):
                return cell.row, cell.column
    return None, None

def find_shape_rows(ws):
    found = {}
    for r in range(1, ws.max_row + 1):
        val = clean(ws.cell(r, 1).value)
        if val in SHAPES and val not in found:
            found[val] = r
    return found

def find_total_row(ws, start):
    for r in range(start, ws.max_row + 1):
        for c in range(1, 10):
            if clean(ws.cell(r, c).value) == "TOTAL":
                return r
    return start + 15

# --------------------------------------------------
# STEP 1 + MISSING SHAPE BLANK:
# Grid -> Master Stock shape-wise numeric copy
# If shape is IN Master Stock but NOT in Grid -> blank ALL numeric cols for that shape
# --------------------------------------------------
def copy_grid_to_stock(ws_grid, ws_stock):
    g = find_shape_rows(ws_grid)
    s = find_shape_rows(ws_stock)

    for shape in SHAPES:
        if shape not in s:
            continue  # shape not in master stock at all, skip

        ss = s[shape]
        se = find_total_row(ws_stock, ss)
        rows = se - ss

        if shape in g:
            # Shape exists in Grid -> copy numeric values
            gs = g[shape]
            ge = find_total_row(ws_grid, gs)
            grid_rows = ge - gs

            for i in range(rows):
                for col in range(1, ws_grid.max_column + 1):
                    if i < grid_rows:
                        val = ws_grid.cell(gs + i, col).value
                        if is_number(val):
                            ws_stock.cell(ss + i, col).value = val
                    # if grid has fewer rows, don't touch remaining stock rows
        else:
            # ✅ Shape MISSING from Grid -> blank ALL numeric columns for this shape in Master Stock
            for i in range(rows):
                for col in range(4, ws_stock.max_column + 1):  # col 1-3 = Shape/Color/Clarity labels, keep those
                    val = ws_stock.cell(ss + i, col).value
                    if is_number(val):
                        ws_stock.cell(ss + i, col).value = None

# --------------------------------------------------
# STEP 2: MAX PCS from Grid -> Master Stock (shape-wise, blank if missing)
# --------------------------------------------------
def copy_max_pcs_grid_to_stock(ws_grid, ws_stock):
    grid_hr, grid_hc = find_header(ws_grid, "MAX PCS")
    stock_hr, stock_hc = find_header(ws_stock, "MAX PCS")
    if not grid_hr or not stock_hr:
        return

    g_shapes = find_shape_rows(ws_grid)
    s_shapes = find_shape_rows(ws_stock)

    for shape in SHAPES:
        if shape not in s_shapes:
            continue

        ss = s_shapes[shape]
        se = find_total_row(ws_stock, ss)
        rows = se - ss

        if shape in g_shapes:
            gs = g_shapes[shape]
            ge = find_total_row(ws_grid, gs)
            grid_rows = ge - gs
            for i in range(rows):
                if i < grid_rows:
                    val = ws_grid.cell(gs + i, grid_hc).value
                    ws_stock.cell(ss + i, stock_hc).value = 0 if (val is None or clean(val) == "") else val
                else:
                    ws_stock.cell(ss + i, stock_hc).value = None
        else:
            # Shape missing from Grid -> blank MAX PCS in Master Stock
            for i in range(rows):
                ws_stock.cell(ss + i, stock_hc).value = None

# --------------------------------------------------
# COPY VALUE COLUMN
# --------------------------------------------------
def copy_value_column(ws_source, source_header, ws_target, target_header):
    sr, sc = find_header(ws_source, source_header)
    tr, tc = find_header(ws_target, target_header)
    if not sr or not tr:
        return
    r, t = sr + 1, tr + 1
    while r <= ws_source.max_row:
        if r > sr + 1000:
            break
        val = ws_source.cell(r, sc).value
        ws_target.cell(t, tc).value = 0 if (val is None or clean(val) == "") else val
        r += 1
        t += 1

# --------------------------------------------------
# PROCESS
# --------------------------------------------------
def process(grid_file, stock_file, master_file):
    wb_grid   = load_workbook(grid_file)
    wb_stock  = load_workbook(stock_file)
    wb_master = load_workbook(master_file)

    ws_grid   = wb_grid.active
    ws_stock  = wb_stock.active
    ws_master = wb_master.active

    # STEP 1: Grid -> Master Stock + blank missing shapes
    copy_grid_to_stock(ws_grid, ws_stock)

    # STEP 2: MAX PCS Grid -> Master Stock (shape-wise, blanks missing)
    copy_max_pcs_grid_to_stock(ws_grid, ws_stock)

    # Save stock, reload to get updated Sheet2
    temp = BytesIO()
    wb_stock.save(temp)
    temp.seek(0)
    wb_stock_values = load_workbook(temp, data_only=False)
    ws_sheet2 = wb_stock_values.worksheets[1]

    # STEP 3: Sheet2 INHAND -> Master File AVAILABLE
    copy_value_column(ws_sheet2, "INHAND", ws_master, "AVAILABLE")

    # STEP 4: Sheet2 MAX PCS -> Master File GRID
    copy_value_column(ws_sheet2, "MAX PCS", ws_master, "GRID")

    out1, out2, out3 = BytesIO(), BytesIO(), BytesIO()
    wb_grid.save(out1)
    wb_stock.save(out2)
    wb_master.save(out3)
    out1.seek(0); out2.seek(0); out3.seek(0)
    return out1, out2, out3

# --------------------------------------------------
# UI — HERO
# --------------------------------------------------
st.markdown("""
<div class="hero">
    <h1>💎 Diamond Work Automation</h1>
    <p>Upload your three files, process in one click, download updated reports.</p>
</div>
""", unsafe_allow_html=True)

col_left, col_right = st.columns([3, 2], gap="large")

with col_left:
    st.markdown('<div class="upload-label">📂 Grid Report</div>', unsafe_allow_html=True)
    grid_file = st.file_uploader("Grid Report", type=["xlsx"], label_visibility="collapsed")

    st.markdown('<div class="upload-label" style="margin-top:1rem;">📦 Master Stock File</div>', unsafe_allow_html=True)
    stock_file = st.file_uploader("Master Stock", type=["xlsx"], label_visibility="collapsed")

    st.markdown('<div class="upload-label" style="margin-top:1rem;">📋 Master File</div>', unsafe_allow_html=True)
    master_file = st.file_uploader("Master File", type=["xlsx"], label_visibility="collapsed")

    st.markdown("<div style='margin-top:1.5rem;'></div>", unsafe_allow_html=True)

    all_uploaded = grid_file and stock_file and master_file

    if all_uploaded:
        if st.button("🚀 Process All Files"):
            try:
                with st.spinner("Processing your files..."):
                    g, s, m = process(grid_file, stock_file, master_file)
                    st.session_state.grid   = g.getvalue()
                    st.session_state.stock  = s.getvalue()
                    st.session_state.master = m.getvalue()
                    st.session_state.processed = True
                st.success("✅ All files processed successfully!")
            except Exception as e:
                st.error(f"❌ Error: {e}")
    else:
        st.button("🚀 Process All Files", disabled=True)
        st.markdown("<p style='color:#4a4f63; font-size:0.82rem; margin-top:0.4rem;'>Upload all 3 files to enable processing.</p>", unsafe_allow_html=True)

with col_right:
    st.markdown("""
    <div class="steps-panel">
        <h4>What this does</h4>
        <div class="step-row">
            <div class="step-num">1</div>
            <div class="step-text">
                <span>Grid Report</span> → <span>Master Stock</span><br>
                <span class="step-arrow">↳</span> Shape-wise numeric copy · <span style="color:#f87171;">blanks entire shape if missing from Grid</span>
            </div>
        </div>
        <div class="step-row">
            <div class="step-num">2</div>
            <div class="step-text">
                <span>Grid Report MAX PCS</span> → <span>Master Stock MAX PCS</span><br>
                <span class="step-arrow">↳</span> Shape-wise · blanks if shape missing in Grid
            </div>
        </div>
        <div class="step-row">
            <div class="step-num">3</div>
            <div class="step-text">
                <span>Master Stock Sheet2 INHAND</span> → <span>Master File AVAILABLE</span>
            </div>
        </div>
        <div class="step-row">
            <div class="step-num">4</div>
            <div class="step-text">
                <span>Master Stock Sheet2 MAX PCS</span> → <span>Master File GRID</span>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

# --------------------------------------------------
# DOWNLOADS
# --------------------------------------------------
if st.session_state.processed:
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown('<div class="dl-header">📥 Download Updated Files</div>', unsafe_allow_html=True)

    d1, d2, d3 = st.columns(3, gap="small")
    with d1:
        st.download_button("📊 Grid Report", data=st.session_state.grid, file_name="Updated_Grid_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with d2:
        st.download_button("📦 Master Stock", data=st.session_state.stock, file_name="Updated_Master_Stock.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with d3:
        st.download_button("📋 Master File", data=st.session_state.master, file_name="Updated_Master_File.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")