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

import streamlit as st
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="Daily Work Automation", layout="wide")
st.title("📊 Full Daily Work Automation")

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
    "CUSHION BRILLIANT","ASSCHER","HEART"
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
        val = clean(ws.cell(r,1).value)

        if val in SHAPES and val not in found:
            found[val] = r

    return found

def find_total_row(ws, start):
    for r in range(start, ws.max_row + 1):
        for c in range(1,10):
            if clean(ws.cell(r,c).value) == "TOTAL":
                return r
    return start + 15

# --------------------------------------------------
# STEP 1 GRID -> MASTER STOCK
# --------------------------------------------------
def copy_grid_to_stock(ws_grid, ws_stock):

    g = find_shape_rows(ws_grid)
    s = find_shape_rows(ws_stock)

    for shape in SHAPES:

        if shape in g and shape in s:

            gs = g[shape]
            ge = find_total_row(ws_grid, gs)

            ss = s[shape]

            rows = ge - gs

            for i in range(rows):

                for col in range(1, ws_grid.max_column + 1):

                    val = ws_grid.cell(gs+i, col).value

                    if is_number(val):
                        ws_stock.cell(ss+i, col).value = val

# --------------------------------------------------
# COPY VALUE ONLY
# --------------------------------------------------
def copy_value_column(ws_source, source_header, ws_target, target_header):

    sr, sc = find_header(ws_source, source_header)
    tr, tc = find_header(ws_target, target_header)

    if not sr or not tr:
        return

    r = sr + 1
    t = tr + 1

    while r <= ws_source.max_row:

        val = ws_source.cell(r, sc).value

        if r > sr + 1000:
            break

        # IMPORTANT: copy only final values
        if val is None or clean(val) == "":
            ws_target.cell(t, tc).value = 0
        else:
            ws_target.cell(t, tc).value = val

        r += 1
        t += 1

# --------------------------------------------------
# PROCESS
# --------------------------------------------------
def process(grid_file, stock_file, master_file):

    # normal files
    wb_grid = load_workbook(grid_file)
    wb_stock = load_workbook(stock_file)
    wb_master = load_workbook(master_file)

    ws_grid = wb_grid.active
    ws_stock = wb_stock.active
    ws_master = wb_master.active

    # ------------------------------------------------
    # STEP 1
    # Grid -> Master Stock
    # ------------------------------------------------
    copy_grid_to_stock(ws_grid, ws_stock)

    # Save updated stock first
    temp = BytesIO()
    wb_stock.save(temp)
    temp.seek(0)

    # ------------------------------------------------
    # IMPORTANT FIX
    # Load SAME updated file values
    # ------------------------------------------------
    wb_stock_values = load_workbook(temp, data_only=False)

    # Sheet2
    ws_sheet2 = wb_stock_values.worksheets[1]

    # ------------------------------------------------
    # STEP 2
    # INHAND -> Available
    # MAX PCS -> Grid
    # ------------------------------------------------
    copy_value_column(ws_sheet2, "INHAND", ws_master, "AVAILABLE")
    copy_value_column(ws_sheet2, "MAX PCS", ws_grid, "GRID")

    # ------------------------------------------------
    # SAVE
    # ------------------------------------------------
    out1 = BytesIO()
    out2 = BytesIO()
    out3 = BytesIO()

    wb_grid.save(out1)
    wb_stock.save(out2)
    wb_master.save(out3)

    out1.seek(0)
    out2.seek(0)
    out3.seek(0)

    return out1, out2, out3

# --------------------------------------------------
# UI
# --------------------------------------------------
grid_file = st.file_uploader("Upload Grid Report", type=["xlsx"])
stock_file = st.file_uploader("Upload Master Stock File", type=["xlsx"])
master_file = st.file_uploader("Upload Master File", type=["xlsx"])

if grid_file and stock_file and master_file:

    if st.button("🚀 Process All Files"):

        try:
            with st.spinner("Processing..."):

                g, s, m = process(grid_file, stock_file, master_file)

                st.session_state.grid = g.getvalue()
                st.session_state.stock = s.getvalue()
                st.session_state.master = m.getvalue()

                st.session_state.processed = True

            st.success("✅ Completed Successfully")

        except Exception as e:
            st.error(str(e))

# --------------------------------------------------
# DOWNLOADS
# --------------------------------------------------
if st.session_state.processed:

    st.download_button(
        "📥 Download Grid Report",
        data=st.session_state.grid,
        file_name="Updated_Grid_Report.xlsx"
    )

    st.download_button(
        "📥 Download Master Stock",
        data=st.session_state.stock,
        file_name="Updated_Master_Stock.xlsx"
    )

    st.download_button(
        "📥 Download Master File",
        data=st.session_state.master,
        file_name="Updated_Master_File.xlsx"
    )