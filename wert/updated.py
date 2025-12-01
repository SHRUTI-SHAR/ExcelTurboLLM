# excel_ai_turbo_engine_cloud_final.py
"""
Excel AI Turbo Engine — Cloud Final

- Works in Linux cloud (Render / Streamlit Cloud)
- Uses xlcalculator if available to evaluate formulas (optional)
- Falls back to reading stored values (data_only=True) when xlcalculator isn't present
- Uses streamlit-aggrid for editable grid UI
- Defensive: will not crash on missing libs or unexpected files
"""

import os
import tempfile
import traceback

import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string

# Grid
try:
    from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
except Exception as e:
    st.error("Required UI package `streamlit-aggrid` (pip package name: streamlit-aggrid) is missing.")
    st.stop()

# Try optional formula engine
XL_CALCULATOR_AVAILABLE = False
try:
    from xlcalculator import ModelCompiler, Evaluator
    XL_CALCULATOR_AVAILABLE = True
except Exception:
    XL_CALCULATOR_AVAILABLE = False

st.set_page_config(page_title="Excel AI Turbo Engine — Cloud", layout="wide")
st.title("🚀 Excel AI Turbo Engine — Cloud Final")
if XL_CALCULATOR_AVAILABLE:
    st.info("Formula engine available: using xlcalculator for Excel formula evaluation.")
else:
    st.warning("Formula engine `xlcalculator` not available. App will show stored values from the uploaded workbook (formulas won't be recalculated).")

# --------------------- Helpers --------------------- #

def save_uploaded_tmp(uploaded_file):
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".xlsx")
    os.close(tmp_fd)
    with open(tmp_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return tmp_path

def load_workbook(file_path, data_only=False):
    try:
        wb = openpyxl.load_workbook(file_path, data_only=data_only)
        return wb
    except Exception as e:
        st.error(f"Could not open workbook: {e}")
        st.stop()

def build_df_from_ws_values(ws):
    rows = list(ws.values)
    if not rows:
        return pd.DataFrame()
    headers = list(rows[0])
    # make headers safe
    for i, h in enumerate(headers):
        if h is None or str(h).strip() == "":
            headers[i] = f"Column_{get_column_letter(i+1)}"
    df = pd.DataFrame(rows[1:], columns=headers)
    return df

def detect_formulas(ws):
    formulas = {}
    for row in ws.iter_rows(values_only=False):
        for cell in row:
            # skip header row
            if cell.row == 1:
                continue
            try:
                has_formula = (cell.data_type == "f") or (isinstance(cell.value, str) and str(cell.value).startswith("="))
            except Exception:
                has_formula = False
            if has_formula:
                formulas[cell.coordinate] = cell.value
    return formulas

def detect_orientation(ws):
    # If first column mostly text → row orientation (key/value), else table
    first_col_vals = [ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)]
    non_empty = [v for v in first_col_vals if v is not None]
    if len(non_empty) == 0:
        return "table"
    text_count = sum(1 for v in non_empty if isinstance(v, str))
    return "row" if (text_count / len(non_empty)) > 0.6 else "table"

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

# --------------------- Formula evaluation --------------------- #

def evaluate_with_xlcalculator(path, sheet_name, updates, max_rows=2000):
    """
    Use xlcalculator ModelCompiler + Evaluator to evaluate workbook.
    updates: dict keyed by (row, col_index) -> value (1-based col index)
    Returns a DataFrame of evaluated values (headers from row 1).
    """
    try:
        mc = ModelCompiler()
        model = mc.read_and_parse_archive(path)
        evaluator = Evaluator(model)

        # apply updates
        for (r, c), val in updates.items():
            addr = f"{sheet_name}!{get_column_letter(c)}{r}"
            try:
                evaluator.set_cell_value(addr, auto_cast(val))
            except Exception:
                # fail quietly per cell
                continue

        # discover sheet dimensions: we'll attempt up to max_rows and columns from model if possible
        # Fallback: use a conservative column count (50) if we can't detect
        # We'll try to read header row first
        cols_to_try = 80

        rows = []
        for r in range(1, max_rows + 1):
            row_vals = []
            empty = True
            for c in range(1, cols_to_try + 1):
                cell_addr = f"{sheet_name}!{get_column_letter(c)}{r}"
                try:
                    val = evaluator.evaluate(cell_addr)
                except Exception:
                    val = None
                if val not in (None, ""):
                    empty = False
                row_vals.append(val)
            if r == 1:
                # always include header row
                rows.append(row_vals)
                continue
            if empty:
                # stop when we find a fully empty row after header
                break
            rows.append(row_vals)

        if len(rows) <= 1:
            return pd.DataFrame()

        headers = rows[0]
        # normalize headers
        clean_headers = []
        seen = {}
        for i, h in enumerate(headers):
            if h is None or str(h).strip() == "":
                h = f"Column_{i+1}"
            h = str(h)
            if h in seen:
                seen[h] += 1
                h = f"{h}_{seen[h]}"
            else:
                seen[h] = 1
            clean_headers.append(h)

        df = pd.DataFrame(rows[1:], columns=clean_headers)
        return df

    except Exception as e:
        st.error("xlcalculator evaluation failed. Falling back to stored values.")
        st.error(str(e))
        return pd.DataFrame()

# --------------------- Main App --------------------- #

uploaded = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])
if not uploaded:
    st.info("Upload an .xlsx file to begin.")
    st.stop()

tmp_path = save_uploaded_tmp(uploaded)

# load two workbooks:
# - values_wb: data_only=True to read last-saved computed values
# - raw_wb: data_only=False to inspect formulas
values_wb = load_workbook(tmp_path, data_only=True)
raw_wb = load_workbook(tmp_path, data_only=False)

sheet_names = values_wb.sheetnames
if not sheet_names:
    st.error("No sheets found in workbook.")
    st.stop()

st.sidebar.header("Workbook")
for s in sheet_names:
    st.sidebar.write(f"- {s}")

selected_sheet = st.selectbox("Select sheet", sheet_names)
if not selected_sheet:
    st.stop()

values_ws = values_wb[selected_sheet]
raw_ws = raw_wb[selected_sheet]

# build dataframe from stored values (safe)
df_values = build_df_from_ws_values(values_ws)

# detect formulas and orientation
formulas = detect_formulas(raw_ws)
orientation = detect_orientation(raw_ws)

st.write(f"Detected orientation: **{orientation}**")
st.write(f"Found {len(formulas)} formula cells in this sheet (formulas may not recalc on cloud if xlcalculator is missing).")

# Prepare editable DataFrame (only input columns if possible)
# Determine output headers (columns that contain formulas) by scanning header row for columns that have formulas below
output_cols = set()
for coord in formulas.keys():
    # coord like 'C5' → column letter = letters from start
    col_letter = ''.join([ch for ch in coord if ch.isalpha()])
    # map letter to header if header exists
    try:
        idx = column_index_from_string(col_letter) - 1
        if idx < len(df_values.columns):
            output_cols.add(df_values.columns[idx])
    except Exception:
        continue

input_headers = [c for c in df_values.columns if c not in output_cols]
if len(input_headers) == 0:
    # fallback: allow editing all columns
    input_headers = list(df_values.columns)

# If orientation = row (key/value), convert to two-column editable DF
if orientation == "row":
    keys = []
    vals = []
    for r in range(2, raw_ws.max_row + 1):
        k = raw_ws.cell(row=r, column=1).value
        v = values_ws.cell(row=r, column=2).value
        if k is not None or v is not None:
            keys.append(k if k is not None else f"Row_{r}")
            vals.append(v)
    editable_df = pd.DataFrame({"Key": keys, "Value": vals})
else:
    editable_df = df_values[input_headers].copy().reset_index(drop=True)

st.markdown("### ✏️ Editable inputs")
gb = GridOptionsBuilder.from_dataframe(editable_df)
gb.configure_default_column(editable=True)
grid_opts = gb.build()

grid_response = AgGrid(
    editable_df,
    gridOptions=grid_opts,
    update_mode=GridUpdateMode.VALUE_CHANGED,
    fit_columns_on_grid_load=True,
    theme="alpine",
    allow_unsafe_jscode=True,
)

edited = pd.DataFrame(grid_response["data"]).reset_index(drop=True)

# Compare to find which rows changed
common_cols = [c for c in edited.columns if c in editable_df.columns]
edited_aligned = edited[common_cols].reset_index(drop=True)
orig_aligned = editable_df[common_cols].reset_index(drop=True)

max_len = max(len(edited_aligned), len(orig_aligned))
edited_aligned = edited_aligned.reindex(range(max_len)).fillna("")
orig_aligned = orig_aligned.reindex(range(max_len)).fillna("")

changed_mask = (edited_aligned != orig_aligned).any(axis=1)
edited["_changed"] = changed_mask

st.markdown("### ⚡ Computed / Displayed Results")

# Build updates mapping for evaluator: (row, col_index) -> value
updates = {}
header_to_letter = {}
# create mapping column letter -> header using raw_ws header row
try:
    header_row = [raw_ws.cell(row=1, column=c).value for c in range(1, raw_ws.max_column + 1)]
    for c in range(1, raw_ws.max_column + 1):
        letter = get_column_letter(c)
        header = header_row[c - 1] if c - 1 < len(header_row) else f"Column_{letter}"
        header_to_letter[str(header)] = letter
except Exception:
    # fallback mapping: use df_values.columns sequentially
    for idx, colname in enumerate(df_values.columns):
        header_to_letter[str(colname)] = get_column_letter(idx + 1)

if orientation == "table":
    # for each edited row, set updates for input headers only
    for r in range(len(edited)):
        for col_header in input_headers:
            if col_header not in header_to_letter:
                continue
            letter = header_to_letter[col_header]
            try:
                col_idx = column_index_from_string(letter)
            except Exception:
                continue
            # row in excel = r + 2 (since header is row 1)
            updates[(r + 2, col_idx)] = edited.at[r, col_header]
else:
    # key/value sheet: map to column 2 values starting row 2
    for r in range(len(edited)):
        updates[(2 + r, 2)] = edited.at[r, "Value"]

# Try to compute using xlcalculator if available
if XL_CALCULATOR_AVAILABLE:
    try:
        calc_df = evaluate_with_xlcalculator(tmp_path, selected_sheet, updates)
        if calc_df.empty:
            st.info("xlcalculator produced empty results; showing stored values instead.")
            calc_df = df_values
    except Exception as e:
        st.warning("xlcalculator failed during evaluation; showing stored values. See error below.")
        st.text(traceback.format_exc())
        calc_df = df_values
else:
    # fallback: apply updates into a copy of df_values without recalculation
    calc_df = df_values.copy().reset_index(drop=True)
    if orientation == "table":
        for (r, c), val in updates.items():
            # r is excel row (1-based). our df index corresponds to excel row - 2
            idx = r - 2
            if idx < 0 or idx >= len(calc_df):
                continue
            # find column name from c
            col_letter = get_column_letter(c)
            # map letter to header
            try:
                col_idx = c - 1
                if col_idx < len(calc_df.columns):
                    calc_df.iat[idx, col_idx] = val
            except Exception:
                continue
    else:
        # key/value sheet: put values in second column if present
        if calc_df.shape[1] >= 2:
            for (r, c), val in updates.items():
                idx = r - 2
                if 0 <= idx < len(calc_df):
                    calc_df.iat[idx, 1] = val

# Mark changed rows in calc_df if length matches; otherwise add column
try:
    calc_df["_changed"] = changed_mask.reindex(range(len(calc_df))).fillna(False).tolist()
except Exception:
    calc_df["_changed"] = False

# Display results grid
AgGrid(calc_df, fit_columns_on_grid_load=True, theme="alpine", allow_unsafe_jscode=True)

st.success("✅ Done — edits applied. If formulas didn't recalc, install xlcalculator in your environment for formula evaluation.")
