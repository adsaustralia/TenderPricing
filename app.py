import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
from openpyxl import load_workbook

st.set_page_config(page_title="Tender Pricing App", layout="wide")


# ---------- Helpers ----------

def num_to_col_letters(n: int) -> str:
    """1 -> A, 2 -> B, ... 27 -> AA, etc."""
    result = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def parse_dimension_to_sqm(dim_str: str) -> float:
    """
    Parse strings like '841mm x 1189mm', '594 x 841mm', '1.2m x 2m' to sqm.
    Assumptions:
    - mm, cm, m supported
    - if no unit, assume mm
    """
    if pd.isna(dim_str):
        return np.nan

    s = str(dim_str).lower()
    s = s.replace("×", "x")

    # Find up to two numbers with optional units
    matches = re.findall(r'(\d+(\.\d+)?)\s*(mm|cm|m)?', s)
    if len(matches) < 2:
        return np.nan

    (v1, _, u1) = matches[0]
    (v2, _, u2) = matches[1]

    v1 = float(v1)
    v2 = float(v2)

    def to_m(v, u):
        if u == "cm":
            return v / 100.0
        if u == "m":
            return v
        # default or mm
        return v / 1000.0

    w = to_m(v1, u1)
    h = to_m(v2, u2)
    return w * h


def detect_side(text, ds_synonyms, ss_synonyms, default="SS"):
    """Return 'DS' or 'SS' based on synonyms found in text."""
    if pd.isna(text):
        return default

    s = str(text).strip().lower()
    if any(tok in s for tok in ds_synonyms):
        return "DS"
    if any(tok in s for tok in ss_synonyms):
        return "SS"
    return default


def build_items_from_rows(
    df,
    col_letters_map,
    size_col_letter,
    material_col_letter,
    qty_annum_col_letter,
    qty_run_col_letter,
    side_mode,
    side_col_letter,
    side_source_letter,
    ds_synonyms,
    ss_synonyms,
    double_sided_loading_percent,
):
    """
    Items are in rows (BP-style).
    """
    letter_to_header = col_letters_map
    result_rows = []

    size_col = letter_to_header.get(size_col_letter)
    mat_col = letter_to_header.get(material_col_letter) if material_col_letter else None
    qty_annum_col = (
        letter_to_header.get(qty_annum_col_letter) if qty_annum_col_letter else None
    )
    qty_run_col = (
        letter_to_header.get(qty_run_col_letter) if qty_run_col_letter else None
    )

    side_col = (
        letter_to_header.get(side_col_letter)
        if side_mode == "Separate column"
        and side_col_letter
        else None
    )
    side_src_col = (
        letter_to_header.get(side_source_letter)
        if side_mode == "Embedded in another column"
        and side_source_letter
        else None
    )

    ds_load_factor = 1.0 + double_sided_loading_percent / 100.0

    for idx, row in df.iterrows():
        size_val = row[size_col] if size_col else None
        material_val = row[mat_col] if mat_col else None

        qty_annum = (
            pd.to_numeric(row[qty_annum_col], errors="coerce")
            if qty_annum_col
            else np.nan
        )
        qty_run = (
            pd.to_numeric(row[qty_run_col], errors="coerce")
            if qty_run_col
            else np.nan
        )

        # Side detection
        if side_mode == "Separate column" and side_col:
            side_raw = row[side_col]
        elif side_mode == "Embedded in another column" and side_src_col:
            side_raw = row[side_src_col]
        else:
            side_raw = None

        side = detect_side(side_raw, ds_synonyms, ss_synonyms, default="SS")

        sqm_per_unit = parse_dimension_to_sqm(size_val)

        sqm_per_annum = (
            sqm_per_unit * qty_annum if (not np.isnan(sqm_per_unit) and not np.isnan(qty_annum)) else np.nan
        )
        sqm_per_run = (
            sqm_per_unit * qty_run if (not np.isnan(sqm_per_unit) and not np.isnan(qty_run)) else np.nan
        )

        result_rows.append(
            {
                "Source Row": idx + 1,  # Excel-style row (data row)
                "Size": size_val,
                "Material": material_val,
                "Qty per annum": qty_annum,
                "Qty per run": qty_run,
                "Side": side,
                "SQM per unit": sqm_per_unit,
                "SQM per annum": sqm_per_annum,
                "SQM per run": sqm_per_run,
            }
        )

    result_df = pd.DataFrame(result_rows)
    return result_df


def build_items_from_columns(
    df,
    size_row,
    material_row,
    qty_annum_row,
    qty_run_row,
    side_mode,
    side_row,
    side_source_row,
    ds_synonyms,
    ss_synonyms,
    double_sided_loading_percent,
):
    """
    Items are in columns (Foot Locker-style).
    size_row etc are 1-based row numbers.
    """
    max_row, max_col = df.shape
    result_rows = []

    ds_load_factor = 1.0 + double_sided_loading_percent / 100.0

    # Convert to 0-based indices (if provided)
    size_r = size_row - 1 if size_row else None
    mat_r = material_row - 1 if material_row else None
    qty_annum_r = qty_annum_row - 1 if qty_annum_row else None
    qty_run_r = qty_run_row - 1 if qty_run_row else None

    side_r = side_row - 1 if (side_mode == "Separate row" and side_row) else None
    side_src_r = (
        side_source_row - 1
        if (side_mode == "Embedded in another row" and side_source_row)
        else None
    )

    for col_idx in range(max_col):
        col_letter = num_to_col_letters(col_idx + 1)

        size_val = df.iloc[size_r, col_idx] if size_r is not None else None
        material_val = df.iloc[mat_r, col_idx] if mat_r is not None else None

        qty_annum = (
            pd.to_numeric(df.iloc[qty_annum_r, col_idx], errors="coerce")
            if qty_annum_r is not None
            else np.nan
        )
        qty_run = (
            pd.to_numeric(df.iloc[qty_run_r, col_idx], errors="coerce")
            if qty_run_r is not None
            else np.nan
        )

        # Skip totally empty items
        if (
            pd.isna(size_val)
            and pd.isna(material_val)
            and np.isnan(qty_annum)
            and np.isnan(qty_run)
        ):
            continue

        # Side detection
        if side_mode == "Separate row" and side_r is not None:
            side_raw = df.iloc[side_r, col_idx]
        elif side_mode == "Embedded in another row" and side_src_r is not None:
            side_raw = df.iloc[side_src_r, col_idx]
        else:
            side_raw = None

        side = detect_side(side_raw, ds_synonyms, ss_synonyms, default="SS")

        sqm_per_unit = parse_dimension_to_sqm(size_val)

        sqm_per_annum = (
            sqm_per_unit * qty_annum if (not np.isnan(sqm_per_unit) and not np.isnan(qty_annum)) else np.nan
        )
        sqm_per_run = (
            sqm_per_unit * qty_run if (not np.isnan(sqm_per_unit) and not np.isnan(qty_run)) else np.nan
        )

        result_rows.append(
            {
                "Source Column": col_letter,
                "Size": size_val,
                "Material": material_val,
                "Qty per annum": qty_annum,
                "Qty per run": qty_run,
                "Side": side,
                "SQM per unit": sqm_per_unit,
                "SQM per annum": sqm_per_annum,
                "SQM per run": sqm_per_run,
            }
        )

    result_df = pd.DataFrame(result_rows)
    return result_df


# ---------- UI ----------

st.title("Tender Pricing App (Steps 1–3)")

st.markdown(
    """
**Step 1:** Upload Excel and view all rows/columns  
**Step 2:** Hide/Unhide rows & columns (without deleting)  
**Step 3:** Map fields (Size, Material, Qty, DS/SS) and calculate SQM + Prices
"""
)

uploaded_file = st.file_uploader(
    "Upload Excel file", type=["xlsx", "xls"], accept_multiple_files=False
)

if uploaded_file is None:
    st.info("Please upload an Excel file to begin.")
    st.stop()

# Read file bytes once for reuse
file_bytes = uploaded_file.read()

# --- Load sheet list ---
excel_file = pd.ExcelFile(BytesIO(file_bytes))
sheet_name = st.selectbox("Select sheet", options=excel_file.sheet_names)

# --- Read selected sheet into DataFrame ---
df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name)

# --- Build Excel-style column letter mapping ---
col_letters = {
    num_to_col_letters(i + 1): col_name for i, col_name in enumerate(df.columns)
}

with st.expander("Show column mapping (Excel letters → headers)"):
    mapping_df = pd.DataFrame(
        {
            "Excel Column": list(col_letters.keys()),
            "Header": [str(v) for v in col_letters.values()],
        }
    )
    st.table(mapping_df)

# ======================================================
# STEP 2: HIDE / UNHIDE COLUMNS & ROWS (for preview + export)
# ======================================================

st.header("Step 2 – Hide / Unhide Rows & Columns")

# Columns to hide
cols_to_hide_letters = st.multiselect(
    "Select columns to HIDE (by Excel letter):",
    options=list(col_letters.keys()),
    default=[],
)
cols_to_hide_headers = [col_letters[letter] for letter in cols_to_hide_letters]

# Rows to hide (1-based rows in DataFrame)
max_row = len(df)
row_numbers = list(range(1, max_row + 1))
rows_to_hide_display = st.multiselect(
    "Select rows to HIDE (by row number in this table):",
    options=row_numbers,
    default=[],
)

# Preview with hidden rows/cols
preview_df = df.copy()
if cols_to_hide_headers:
    preview_df = preview_df.drop(columns=cols_to_hide_headers)
if rows_to_hide_display:
    indices_to_drop = [r - 1 for r in rows_to_hide_display]
    preview_df = preview_df.drop(index=indices_to_drop)

st.subheader(f"Preview: {sheet_name}")
st.caption(
    "Preview hides selected rows/columns. Original workbook remains intact; "
    "exported file will mark them as hidden in Excel."
)
st.dataframe(preview_df)

# Export with hidden rows/columns
st.subheader("Export with Hidden Rows / Columns")
if st.button("Prepare file with hidden rows/columns"):
    wb = load_workbook(BytesIO(file_bytes))
    ws = wb[sheet_name]

    # Hide selected columns
    for letter in cols_to_hide_letters:
        ws.column_dimensions[letter].hidden = True

    # Hide selected rows (data rows: +1 because header is row 1)
    for r in rows_to_hide_display:
        excel_row = r + 1
        ws.row_dimensions[excel_row].hidden = True

    out_buf = BytesIO()
    wb.save(out_buf)
    out_buf.seek(0)

    st.download_button(
        "Download workbook (with hidden rows/columns)",
        data=out_buf,
        file_name=f"{sheet_name}_hidden.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )

# ======================================================
# STEP 3: SQM & PRICE CALCULATION
# ======================================================

st.header("Step 3 – SQM & Price Calculation")

st.markdown(
    """
Here you tell the app **where** the data lives (columns vs rows) and how DS/SS is encoded,
so it can calculate **square meters** and **pricing by material**.
"""
)

layout_type = st.radio(
    "How are items laid out in this sheet?",
    ["Items are in rows (BP-style)", "Items are in columns (Foot Locker-style)"],
)

# DS/SS synonyms + loading
st.subheader("Double-sided / Single-sided configuration")

ds_syn_input = st.text_input(
    "Values meaning DOUBLE-SIDED (comma-separated)",
    value="ds,double sided,double-sided,2s,2 sided,2sided,double",
)
ss_syn_input = st.text_input(
    "Values meaning SINGLE-SIDED (comma-separated)",
    value="ss,single sided,single-sided,1s,1 sided,1sided,single",
)

ds_synonyms = [s.strip().lower() for s in ds_syn_input.split(",") if s.strip()]
ss_synonyms = [s.strip().lower() for s in ss_syn_input.split(",") if s.strip()]

double_sided_loading_percent = st.number_input(
    "Double-sided loading % (e.g. 25 for 25% extra over single-sided)",
    min_value=0.0,
    max_value=500.0,
    value=25.0,
    step=1.0,
)

calc_df = None  # will hold result if we calculate

if layout_type == "Items are in rows (BP-style)":
    st.subheader("Mapping (items in rows)")

    letters = list(col_letters.keys())

    # Try to auto-guess some defaults by header name
    headers_lower = {ltr: str(h).lower() for ltr, h in col_letters.items()}

    def guess_letter(substrings, fallback):
        for ltr, h in headers_lower.items():
            if any(sub in h for sub in substrings):
                return ltr
        return fallback

    size_default = guess_letter(["dim", "size"], letters[0] if letters else None)
    material_default = guess_letter(
        ["material", "stock", "substrate"], letters[0] if letters else None
    )
    qty_annum_default = guess_letter(
        ["annual", "per annum", "pa"], letters[0] if letters else None
    )
    qty_run_default = guess_letter(
        ["per run", "run qty", "run quantity"], letters[0] if letters else None
    )

    size_col_letter = st.selectbox(
        "Size / Dimensions column (Excel letter)", options=letters, index=letters.index(size_default) if size_default in letters else 0
    )
    material_col_letter = st.selectbox(
        "Material name column (Excel letter)",
        options=["(none)"] + letters,
        index=(letters.index(material_default) + 1) if material_default in letters else 0,
    )
    qty_annum_col_letter = st.selectbox(
        "Quantity PER ANNUM column (Excel letter)",
        options=["(none)"] + letters,
        index=(letters.index(qty_annum_default) + 1) if qty_annum_default in letters else 0,
    )
    qty_run_col_letter = st.selectbox(
        "Quantity PER RUN column (Excel letter)",
        options=["(none)"] + letters,
        index=(letters.index(qty_run_default) + 1) if qty_run_default in letters else 0,
    )

    # Convert "(none)" to None
    material_col_letter = None if material_col_letter == "(none)" else material_col_letter
    qty_annum_col_letter = None if qty_annum_col_letter == "(none)" else qty_annum_col_letter
    qty_run_col_letter = None if qty_run_col_letter == "(none)" else qty_run_col_letter

    st.markdown("**Where is Single / Double-sided information?**")
    side_mode = st.selectbox(
        "Choose how DS/SS is stored:",
        ["Separate column", "Embedded in another column", "Not available (assume SS)"],
    )

    side_col_letter = None
    side_source_letter = None

    if side_mode == "Separate column":
        side_col_letter = st.selectbox(
            "Column that contains DS/SS values",
            options=letters,
        )
    elif side_mode == "Embedded in another column":
        side_source_letter = st.selectbox(
            "Column where DS/SS text appears (e.g. Size or Description)",
            options=letters,
            index=letters.index(size_col_letter) if size_col_letter in letters else 0,
        )

    if st.button("Calculate SQM & build item table", key="calc_rows"):
        calc_df = build_items_from_rows(
            df=df,
            col_letters_map=col_letters,
            size_col_letter=size_col_letter,
            material_col_letter=material_col_letter,
            qty_annum_col_letter=qty_annum_col_letter,
            qty_run_col_letter=qty_run_col_letter,
            side_mode=side_mode,
            side_col_letter=side_col_letter,
            side_source_letter=side_source_letter,
            ds_synonyms=ds_synonyms,
            ss_synonyms=ss_synonyms,
            double_sided_loading_percent=double_sided_loading_percent,
        )

elif layout_type == "Items are in columns (Foot Locker-style)":
    st.subheader("Mapping (items in columns)")

    max_row, max_col = df.shape
    row_options = list(range(1, max_row + 1))

    size_row = st.selectbox(
        "Row number that contains Size / Dimensions (across columns)",
        options=row_options,
        index=0,
    )
    material_row = st.selectbox(
        "Row number that contains Material name (across columns)",
        options=["(none)"] + row_options,
        index=0,
    )
    qty_annum_row = st.selectbox(
        "Row number that contains Quantity PER ANNUM (across columns)",
        options=["(none)"] + row_options,
        index=0,
    )
    qty_run_row = st.selectbox(
        "Row number that contains Quantity PER RUN (across columns)",
        options=["(none)"] + row_options,
        index=0,
    )

    # Convert "(none)" to None
    material_row = None if material_row == "(none)" else material_row
    qty_annum_row = None if qty_annum_row == "(none)" else qty_annum_row
    qty_run_row = None if qty_run_row == "(none)" else qty_run_row

    st.markdown("**Where is Single / Double-sided information?**")
    side_mode = st.selectbox(
        "Choose how DS/SS is stored:",
        ["Separate row", "Embedded in another row", "Not available (assume SS)"],
    )

    side_row = None
    side_source_row = None

    if side_mode == "Separate row":
        side_row = st.selectbox(
            "Row that contains DS/SS values (across columns)",
            options=row_options,
        )
    elif side_mode == "Embedded in another row":
        side_source_row = st.selectbox(
            "Row where DS/SS text appears (e.g. in Size or Description row)",
            options=row_options,
            index=row_options.index(size_row) if size_row in row_options else 0,
        )

    if st.button("Calculate SQM & build item table", key="calc_cols"):
        calc_df = build_items_from_columns(
            df=df,
            size_row=size_row,
            material_row=material_row,
            qty_annum_row=qty_annum_row,
            qty_run_row=qty_run_row,
            side_mode=side_mode,
            side_row=side_row,
            side_source_row=side_source_row,
            ds_synonyms=ds_synonyms,
            ss_synonyms=ss_synonyms,
            double_sided_loading_percent=double_sided_loading_percent,
        )

# ---------- Show calculation results + Material price mapping ----------

if calc_df is not None:
    st.subheader("Calculated SQM table (before pricing)")
    st.dataframe(calc_df)

    st.subheader("Material Pricing (per sqm)")

    # Build unique material list for pricing
    materials = sorted(
        {m for m in calc_df["Material"].dropna().unique()} if "Material" in calc_df.columns else []
    )
    price_df = pd.DataFrame(
        {"Material": materials, "Price per SQM": [np.nan] * len(materials)}
    )

    edited_price_df = st.data_editor(
        price_df,
        num_rows="dynamic",
        key="material_price_editor",
        use_container_width=True,
    )

    # Merge prices back
    calc_with_price = calc_df.merge(
        edited_price_df, how="left", on="Material"
    )

    # Apply DS loading
    ds_factor = 1.0 + double_sided_loading_percent / 100.0
    calc_with_price["Effective Price per SQM"] = calc_with_price.apply(
        lambda r: r["Price per SQM"] * ds_factor if r.get("Side") == "DS" else r["Price per SQM"],
        axis=1,
    )

    # Price calculations
    calc_with_price["Price per unit"] = (
        calc_with_price["SQM per unit"] * calc_with_price["Effective Price per SQM"]
    )
    calc_with_price["Price per annum"] = (
        calc_with_price["SQM per annum"] * calc_with_price["Effective Price per SQM"]
    )
    calc_with_price["Price per run"] = (
        calc_with_price["SQM per run"] * calc_with_price["Effective Price per SQM"]
    )

    st.subheader("Final calculation table (including pricing)")
    st.dataframe(calc_with_price)

    # Download calculated table
    out_calc = BytesIO()
    calc_with_price.to_excel(out_calc, index=False, sheet_name="CALC")
    out_calc.seek(0)

    st.download_button(
        "Download SQM & pricing table (CALC.xlsx)",
        data=out_calc,
        file_name="sqm_pricing_calc.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )
