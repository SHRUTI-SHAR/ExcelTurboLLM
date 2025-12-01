# excel_ai_turbo_engine_cloud.py
"""
Excel AI Turbo Engine — CLOUD Version

✔ Works on Streamlit Cloud / Render (Linux)
✔ No Excel COM required (removed)
✔ Supports Excel formulas using xlcalculator (pure Python)
✔ Auto-detects inputs + outputs
✔ Fully rewritten clean version
"""

import os
import tempfile
import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

# Excel formula engine (Python)
from xlcalculator import ModelCompiler, Evaluator


# --------------------------------------------------
# Streamlit UI
# --------------------------------------------------

st.set_page_config(page_title="Excel AI Turbo Engine — Cloud", layout="wide")
st.title("🚀 Excel AI Turbo Engine — Cloud Version")
st.markdown("Upload your Excel file, edit inputs, and auto-recalculate using a Python formula engine.")


# --------------------------------------------------
# Load workbook
# --------------------------------------------------

def load_workbook(uploaded_file):
    try:
        wb = openpyxl.load_workbook(uploaded_file, data_only=False)
        return wb
    except Exception as e:
        st.error(f"Could not open workbook: {e}")
        st.stop()


def list_sheets_with_state(wb):
    return [(ws.title, getattr(ws, "sheet_state", "visible")) for ws in wb.worksheets]


def sheet_to_dataframe(ws):
    rows = list(ws.values)
    if not rows:
        return pd.DataFrame(), [], {}

    headers = list(rows[0])
    for i, h in enumerate(headers):
        if h is None or str(h).strip() == "":
            headers[i] = f"Column_{get_column_letter(i+1)}"

    df = pd.DataFrame(rows[1:], columns=headers)

    # mapping Excel column letter → header
    colmap = {}
    for col_idx in range(1, ws.max_column + 1):
        letter = get_column_letter(col_idx)
        header = headers[col_idx - 1] if col_idx - 1 < len(headers) else f"Column_{letter}"
        colmap[letter] = str(header)

    return df, headers, colmap


# --------------------------------------------------
# Detect formulas & input cells
# --------------------------------------------------

def detect_formulas_and_inputs(ws, df, colmap):
    formulas_map = {}
    formula_cells = []

    for row in ws.iter_rows(values_only=False):
        for cell in row:
            if cell.row == 1:
                continue
            if cell.data_type == "f" or (isinstance(cell.value, str) and str(cell.value).startswith("=")):
                formulas_map[cell.column_letter] = cell.value
                formula_cells.append({"cell": cell.coordinate, "formula": cell.value})

    output_headers = [colmap[c] for c in formulas_map.keys() if c in colmap]
    input_headers = [h for h in df.columns if h not in output_headers]

    return formulas_map, formula_cells, input_headers, output_headers


# --------------------------------------------------
# Detect orientation of sheet (table or key/value)
# --------------------------------------------------

def detect_orientation(ws, df):
    first_col_vals = [ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)]

    text_count = sum(1 for v in first_col_vals if isinstance(v, str))
    non_empty = sum(1 for v in first_col_vals if v is not None)

    if non_empty == 0:
        return "table"

    return "row" if text_count / non_empty > 0.6 else "table"


# --------------------------------------------------
# Python-based Excel Formula Calculation
# --------------------------------------------------

def auto_cast(v):
    try:
        if v is None:
            return None
        if isinstance(v, (float, int, str, bool)):
            if isinstance(v, float) and np.isnan(v):
                return None
            return v
        if hasattr(v, "item"):
            return v.item()
        if isinstance(v, pd.Timestamp):
            return v.to_pydatetime()
        return v
    except Exception:
        return v


def python_excel_calculate(file_path, sheet_name, updates, colmap):
    try:
        mc = ModelCompiler()
        model = mc.read_and_parse_archive(file_path)
        evaluator = Evaluator(model)

        # apply user input updates
        for (r, c), val in updates.items():
            cell_name = f"{sheet_name}!{get_column_letter(c)}{r}"
            evaluator.set_cell_value(cell_name, auto_cast(val))

        rows = []
        max_rows = 2000
        max_cols = len(colmap)

        for r in range(1, max_rows):
            row_vals = []
            empty = True
            for c in range(1, max_cols + 1):
                cell = f"{sheet_name}!{get_column_letter(c)}{r}"
                try:
                    val = evaluator.evaluate(cell)
                    if val not in [None, ""]:
                        empty = False
                except Exception:
                    val = None
                row_vals.append(val)
            if empty:
                break
            rows.append(row_vals)

        if not rows:
            return pd.DataFrame()

        headers = rows[0]
        clean_headers = []
        used = {}

        for i, h in enumerate(headers):
            if h is None or str(h).strip() == "":
                h = f"Column_{i+1}"
            if h in used:
                used[h] += 1
                h = f"{h}_{used[h]}"
            else:
                used[h] = 1
            clean_headers.append(str(h))

        df = pd.DataFrame(rows[1:], columns=clean_headers)
        return df

    except Exception as e:
        st.error(f"Python Excel evaluator failed: {e}")
        return pd.DataFrame()


# --------------------------------------------------
# MAIN APP
# --------------------------------------------------

uploaded = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])
if not uploaded:
    st.stop()

tmp_fd, tmp_path = tempfile.mkstemp(suffix=".xlsx")
os.close(tmp_fd)
with open(tmp_path, "wb") as f:
    f.write(uploaded.getbuffer())

wb = load_workbook(tmp_path)
sheets = list_sheets_with_state(wb)
sheet_names = [name for name, _ in sheets]

st.sidebar.header("Sheets")
for name, state in sheets:
    st.sidebar.write(f"- {name} — {state}")

selected_sheet = st.selectbox("Select sheet", sheet_names)
ws = wb[selected_sheet]

df, headers, colmap = sheet_to_dataframe(ws)
formulas_map, formula_cells, input_headers, output_headers = detect_formulas_and_inputs(ws, df, colmap)
orientation = detect_orientation(ws, df)

st.write(f"Detected orientation: **{orientation}**")
st.write(f"Detected outputs: {output_headers}")
st.write(f"Detected inputs: {input_headers}")


# --------------------------------------------------
# AGGRID Editable Table
# --------------------------------------------------

if orientation == "table":
    editable_df = df[input_headers].copy().reset_index(drop=True)
else:
    keys, vals = [], []
    for r in range(2, ws.max_row + 1):
        k = ws.cell(row=r, column=1).value
        v = ws.cell(row=r, column=2).value
        if k or v:
            keys.append(k if k else f"Row_{r}")
            vals.append(v)
    editable_df = pd.DataFrame({"Key": keys, "Value": vals})

st.markdown("### ✏️ Editable Inputs")
gb = GridOptionsBuilder.from_dataframe(editable_df)
gb.configure_default_column(editable=True)
grid_options = gb.build()

grid_response = AgGrid(
    editable_df,
    gridOptions=grid_options,
    update_mode=GridUpdateMode.VALUE_CHANGED,
    fit_columns_on_grid_load=True,
    theme="alpine",
    allow_unsafe_jscode=True,
)

edited = pd.DataFrame(grid_response["data"]).reset_index(drop=True)

# detect changes
common_cols = [c for c in edited.columns if c in editable_df.columns]
edited_aligned = edited[common_cols].reset_index(drop=True)
orig_aligned = editable_df[common_cols].reset_index(drop=True)

max_len = max(len(edited_aligned), len(orig_aligned))
edited_aligned = edited_aligned.reindex(range(max_len)).fillna("")
orig_aligned = orig_aligned.reindex(range(max_len)).fillna("")

changed_mask = (edited_aligned != orig_aligned).any(axis=1)
edited["_changed"] = changed_mask

st.markdown("### ⚡ Live Computed Results")
st.info("Using Python Excel formula engine (xlcalculator).")


# --------------------------------------------------
# APPLY UPDATES + RECALCULATE
# --------------------------------------------------

updates = {}
header_to_letter = {v: k for k, v in colmap.items()}

if orientation == "table":
    for r in range(len(edited)):
        for col_header in input_headers:
            letter = header_to_letter.get(col_header)
            if letter:
                col_idx = column_index_from_string(letter)
                updates[(r + 2, col_idx)] = edited.at[r, col_header]
else:
    for r in range(len(edited)):
        updates[(2 + r, 2)] = edited.at[r, "Value"]

calc_df = python_excel_calculate(tmp_path, selected_sheet, updates, colmap)
calc_df["_changed"] = changed_mask.reindex(range(len(calc_df))).fillna(False).tolist()

AgGrid(calc_df, fit_columns_on_grid_load=True, theme="alpine", allow_unsafe_jscode=True)

st.success("✅ Live update complete. Edit values above to auto-recalculate.")
