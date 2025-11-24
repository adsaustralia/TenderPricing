import streamlit as st
import pandas as pd
import numpy as np
import re
import json
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


def to_excel_view(df: pd.DataFrame) -> pd.DataFrame:
    """
    Build a view that looks like Excel:
    - Column headers: A, B, C...
    - Row numbers: 1, 2, 3...
    - Row 1 contains the original header names.
    """
    letters = [num_to_col_letters(i + 1) for i in range(len(df.columns))]
    data = [list(df.columns)] + df.astype(object).values.tolist()
    excel_df = pd.DataFrame(data, columns=letters)
    excel_df.index = range(1, len(excel_df) + 1)
    excel_df.index.name = ""
    return excel_df


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
    runs_pa_col_letter,
    side_mode,
    side_col_letter,
    side_source_letter,
    ds_synonyms,
    ss_synonyms,
    double_sided_loading_percent,
):
    """
    Items are in rows (BP-style).
    Supports either:
    - explicit Qty per run column, OR
    - Qty per annum + Runs p.a. column -> Qty per run = Qty per annum / Runs p.a.
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
    runs_pa_col = (
        letter_to_header.get(runs_pa_col_letter) if runs_pa_col_letter else None
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

    for idx, row in df.iterrows():
        size_val = row[size_col] if size_col else None
        material_val = row[mat_col] if mat_col else None

        qty_annum = (
            pd.to_numeric(row[qty_annum_col], errors="coerce")
            if qty_annum_col
            else np.nan
        )

        if qty_run_col:
            qty_run = pd.to_numeric(row[qty_run_col], errors="coerce")
        elif runs_pa_col and qty_annum_col:
            runs_pa = pd.to_numeric(row[runs_pa_col], errors="coerce")
            if not np.isnan(qty_annum) and not np.isnan(runs_pa) and runs_pa > 0:
                qty_run = qty_annum / runs_pa
            else:
                qty_run = np.nan
        else:
            qty_run = np.nan

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
                "Source Row": idx + 2,  # Excel-style row: +2 because header is row 1
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
    runs_pa_row,
    side_mode,
    side_row,
    side_source_row,
    ds_synonyms,
    ss_synonyms,
    double_sided_loading_percent,
):
    """
    Items are in columns (Foot Locker-style).
    Excel rows are used for mapping. Supports:
    - explicit Qty per run row, OR
    - Qty per annum row + Runs p.a. row -> Qty per run = Qty per annum / Runs p.a.
    """
    max_row, max_col = df.shape
    result_rows = []

    # Convert Excel row to df index (Excel row 2 -> df index 0)
    def excel_to_df_row(excel_row):
        return excel_row - 2

    size_r = excel_to_df_row(size_row) if size_row else None
    mat_r = excel_to_df_row(material_row) if material_row else None
    qty_annum_r = excel_to_df_row(qty_annum_row) if qty_annum_row else None
    qty_run_r = excel_to_df_row(qty_run_row) if qty_run_row else None
    runs_pa_r = excel_to_df_row(runs_pa_row) if runs_pa_row else None

    side_r = excel_to_df_row(side_row) if (side_mode == "Separate row" and side_row) else None
    side_src_r = (
        excel_to_df_row(side_source_row)
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

        if qty_run_r is not None:
            qty_run = pd.to_numeric(df.iloc[qty_run_r, col_idx], errors="coerce")
        elif runs_pa_r is not None and qty_annum_r is not None:
            runs_pa = pd.to_numeric(df.iloc[runs_pa_r, col_idx], errors="coerce")
            if not np.isnan(qty_annum) and not np.isnan(runs_pa) and runs_pa > 0:
                qty_run = qty_annum / runs_pa
            else:
                qty_run = np.nan
        else:
            qty_run = np.nan

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

st.title("Tender Pricing App (Excel-style, Grouped Pricing)")

st.markdown(
    """
**Step 1:** Upload Excel and view all rows/columns (Excel-style A,B,C + 1,2,3)  
**Step 2:** Hide/Unhide rows & columns (without deleting)  
**Step 3:** Map fields (Size, Material, Qty, DS/SS) and calculate SQM + **Grouped Prices**
"""
)

# ---------- Load default preset at startup (from repo) ----------
if "group_assignments" not in st.session_state:
    try:
        with open("material_groups_default.json", "r", encoding="utf-8") as f:
            preset = json.load(f)
        st.session_state["group_assignments"] = preset.get("group_assignments", {})
        st.session_state["group_prices"] = preset.get("group_prices", {})
        st.session_state["material_overrides"] = preset.get("material_overrides", {})
    except Exception:
        st.session_state["group_assignments"] = {}
        st.session_state["group_prices"] = {}
        st.session_state["material_overrides"] = {}

uploaded_file = st.file_uploader(
    "Upload Excel file", type=["xlsx", "xls"], accept_multiple_files=False
)

if uploaded_file is None:
    st.info("Please upload an Excel file to begin.")
    st.stop()

# IMPORTANT FIX: use getvalue() instead of read() so the file content
# is available on every rerun and typing in data_editor doesn't wipe out the data.
file_bytes = uploaded_file.getvalue()

# --- Load sheet list ---
excel_file = pd.ExcelFile(BytesIO(file_bytes))
sheet_name = st.selectbox("Select sheet", options=excel_file.sheet_names)

# --- Read selected sheet into DataFrame ---
df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name)

# Show the sheet in Excel-like view (A,B,C... and rows 1,2,3...)
st.subheader(f"Sheet preview (Excel-style): {sheet_name}")
excel_view = to_excel_view(df)
st.dataframe(excel_view)

# --- Build Excel-style column letter mapping (internal) ---
# And also build friendly labels like "A - Lot ID"
col_letters = {}
col_labels = {}
for i, col_name in enumerate(df.columns):
    letter = num_to_col_letters(i + 1)
    col_letters[letter] = col_name
    label = f"{letter} - {col_name}"
    col_labels[letter] = label

# Helper to show dropdowns with friendly labels but keep track of the letter
def select_letter(label, options_letters, default_letter=None, key=None, allow_none=False, none_label="(none)"):
    options = []
    mapping = {}
    if allow_none:
        options.append(none_label)
        mapping[none_label] = None
    for ltr in options_letters:
        lab = col_labels[ltr]
        options.append(lab)
        mapping[lab] = ltr

    if default_letter is not None and default_letter in options_letters:
        default_label = col_labels[default_letter]
        default_index = options.index(default_label)
    else:
        default_index = 0

    choice = st.selectbox(label, options=options, index=default_index, key=key)
    return mapping[choice]

# ======================================================
# STEP 2: HIDE / UNHIDE COLUMNS & ROWS (for preview + export)
# ======================================================

st.header("Step 2 – Hide / Unhide Rows & Columns")

all_letters = list(col_letters.keys())

# Columns to hide
cols_to_hide_labels = st.multiselect(
    "Select columns to HIDE (by Excel letter):",
    options=[col_labels[ltr] for ltr in all_letters],
    default=[],
)
cols_to_hide_letters = [opt.split(" - ")[0] for opt in cols_to_hide_labels]

# Rows to hide (Excel-style rows: include header row 1)
max_row = len(df) + 1  # +1 for header row
row_numbers = list(range(1, max_row + 1))
rows_to_hide_display = st.multiselect(
    "Select rows to HIDE (by Excel row number):",
    options=row_numbers,
    default=[],
)

# Preview with hidden rows/cols (Excel-like view)
preview_excel_view = excel_view.copy()
if cols_to_hide_letters:
    preview_excel_view = preview_excel_view.drop(columns=cols_to_hide_letters)
if rows_to_hide_display:
    preview_excel_view = preview_excel_view.drop(index=rows_to_hide_display)

st.subheader(f"Preview with hidden rows/columns (Excel-style): {sheet_name}")
st.caption(
    "Preview hides selected rows/columns. Original workbook remains intact; "
    "exported file will mark them as hidden in Excel."
)
st.dataframe(preview_excel_view)

# Export with hidden rows/columns
st.subheader("Export with Hidden Rows / Columns")
if st.button("Prepare file with hidden rows/columns"):
    wb = load_workbook(BytesIO(file_bytes))
    ws = wb[sheet_name]

    # Hide selected columns
    for letter in cols_to_hide_letters:
        ws.column_dimensions[letter].hidden = True

    # Hide selected rows (Excel rows directly)
    for r in rows_to_hide_display:
        ws.row_dimensions[r].hidden = True

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

In this setup, **Qty per run** can either come from:
- An explicit "Qty per run" column/row, OR
- `Qty per annum ÷ Runs per annum`.

In the **Material Pricing** area you can:
- Assign each material to a **Group name** (e.g. "3mm ACM", "Posters", "Window Vinyl").  
- Enter one **Group Price per SQM** to apply to all materials in that group.  
- Optionally override a single material with its own price.  
- Save/Load **group presets** so you can reuse them next campaign/tender.
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

# Currency configuration
st.subheader("Currency configuration")

st.caption("Base calculations are in AUD. You can also display a converted currency.")
display_currency = st.text_input("Display currency code", value="AUD")
conversion_rate = st.number_input(
    f"Conversion rate (1 AUD = X {display_currency})",
    value=1.0,
    min_value=0.0001,
    step=0.01,
)

calc_df = None  # will hold numeric result if we calculate

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
        ["annual", "per annum", "total annual volume", "pa"], letters[0] if letters else None
    )
    qty_run_default = guess_letter(
        ["per run", "run qty", "run quantity"], letters[0] if letters else None
    )
    runs_pa_default = guess_letter(
        ["runs p.a", "approx runs", "runs pa"], letters[0] if letters else None
    )

    size_col_letter = select_letter(
        "Size / Dimensions column",
        options_letters=letters,
        default_letter=size_default,
        key="size_col_letter_rows"
    )
    material_col_letter = select_letter(
        "Material name column",
        options_letters=letters,
        default_letter=material_default,
        key="material_col_letter_rows",
        allow_none=True,
    )
    qty_annum_col_letter = select_letter(
        "Quantity PER ANNUM column",
        options_letters=letters,
        default_letter=qty_annum_default,
        key="qty_annum_col_letter_rows",
        allow_none=True,
    )
    qty_run_col_letter = select_letter(
        "Quantity PER RUN column (optional — leave None to derive from runs p.a)",
        options_letters=letters,
        default_letter=qty_run_default,
        key="qty_run_col_letter_rows",
        allow_none=True,
    )
    runs_pa_col_letter = select_letter(
        "Runs PER ANNUM column (e.g. Approx runs p.a, optional)",
        options_letters=letters,
        default_letter=runs_pa_default,
        key="runs_pa_col_letter_rows",
        allow_none=True,
    )

    st.markdown("**Where is Single / Double-sided information?**")
    side_mode = st.selectbox(
        "Choose how DS/SS is stored:",
        ["Separate column", "Embedded in another column", "Not available (assume SS)"],
    )

    side_col_letter = None
    side_source_letter = None

    if side_mode == "Separate column":
        side_col_letter = select_letter(
            "Column that contains DS/SS values",
            options_letters=letters,
            key="side_col_letter_rows"
        )
    elif side_mode == "Embedded in another column":
        side_source_letter = select_letter(
            "Column where DS/SS text appears (e.g. Size or Description)",
            options_letters=letters,
            default_letter=size_col_letter,
            key="side_source_letter_rows"
        )

    if st.button("Calculate SQM & build item table", key="calc_rows"):
        calc_df = build_items_from_rows(
            df=df,
            col_letters_map=col_letters,
            size_col_letter=size_col_letter,
            material_col_letter=material_col_letter,
            qty_annum_col_letter=qty_annum_col_letter,
            qty_run_col_letter=qty_run_col_letter,
            runs_pa_col_letter=runs_pa_col_letter,
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
    # Excel rows: header row is 1, df data starts at Excel row 2
    row_options = list(range(2, max_row + 2))

    size_row = st.selectbox(
        "Excel row that contains Size / Dimensions (across columns)",
        options=row_options,
        index=0,
    )
    material_row = st.selectbox(
        "Excel row that contains Material name (across columns)",
        options=["(none)"] + row_options,
        index=0,
    )
    qty_annum_row = st.selectbox(
        "Excel row that contains Quantity PER ANNUM (across columns)",
        options=["(none)"] + row_options,
        index=0,
    )
    qty_run_row = st.selectbox(
        "Excel row that contains Quantity PER RUN (across columns)",
        options=["(none)"] + row_options,
        index=0,
    )
    runs_pa_row = st.selectbox(
        "Excel row that contains Runs PER ANNUM (across columns, optional)",
        options=["(none)"] + row_options,
        index=0,
    )

    # Convert "(none)" to None
    material_row = None if material_row == "(none)" else material_row
    qty_annum_row = None if qty_annum_row == "(none)" else qty_annum_row
    qty_run_row = None if qty_run_row == "(none)" else qty_run_row
    runs_pa_row = None if runs_pa_row == "(none)" else runs_pa_row

    st.markdown("**Where is Single / Double-sided information?**")
    side_mode = st.selectbox(
        "Choose how DS/SS is stored:",
        ["Separate row", "Embedded in another row", "Not available (assume SS)"],
    )

    side_row = None
    side_source_row = None

    if side_mode == "Separate row":
        side_row = st.selectbox(
            "Excel row that contains DS/SS values (across columns)",
            options=row_options,
        )
    elif side_mode == "Embedded in another row":
        side_source_row = st.selectbox(
            "Excel row where DS/SS text appears (e.g. in Size or Description row)",
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
            runs_pa_row=runs_pa_row,
            side_mode=side_mode,
            side_row=side_row,
            side_source_row=side_source_row,
            ds_synonyms=ds_synonyms,
            ss_synonyms=ss_synonyms,
            double_sided_loading_percent=double_sided_loading_percent,
        )

# ---------- Show calculation results + Material group pricing ----------

if calc_df is not None:
    st.subheader("Calculated SQM table (before pricing, numeric)")
    # Round sqm and qty to 2 decimals for clarity
    for col in ["Qty per annum", "Qty per run", "SQM per unit", "SQM per annum", "SQM per run"]:
        if col in calc_df.columns:
            calc_df[col] = calc_df[col].round(2)
    st.dataframe(calc_df)

    # ---------- Material groups + pricing presets ----------
    st.subheader("Material Groups & Pricing Presets")

    # Load saved preset (optional override of default)
    preset_file = st.file_uploader(
        "Load saved material groups preset (JSON, optional, overrides default)",
        type=["json"],
        key="group_preset_uploader",
    )
    if preset_file is not None:
        try:
            preset = json.load(preset_file)
            st.session_state["group_assignments"] = preset.get("group_assignments", {})
            st.session_state["group_prices"] = preset.get("group_prices", {})
            st.session_state["material_overrides"] = preset.get("material_overrides", {})
            st.success("Loaded group preset from uploaded file.")
        except Exception as e:
            st.error(f"Failed to load preset: {e}")

    # Materials list
    materials = sorted(
        {m for m in calc_df["Material"].dropna().unique()} if "Material" in calc_df.columns else []
    )

    # Existing assignments/prices from session
    existing_assignments = st.session_state.get("group_assignments", {})
    existing_group_prices = st.session_state.get("group_prices", {})
    existing_overrides = st.session_state.get("material_overrides", {})

    # Build group assignment table
    group_assign_rows = []
    for m in materials:
        group_assign_rows.append(
            {
                "Material": m,
                "Group": existing_assignments.get(m, ""),
            }
        )
    group_assign_df = pd.DataFrame(group_assign_rows)

    st.markdown("**Step 1 – Assign materials to groups**")
    edited_group_assign_df = st.data_editor(
        group_assign_df,
        num_rows="fixed",
        use_container_width=True,
        key="group_assign_editor",
    )

    # Derive list of groups
    groups = sorted(
        {
            str(g).strip()
            for g in edited_group_assign_df["Group"].dropna().unique()
            if str(g).strip() != ""
        }
    )

    # Build group price table
    group_price_rows = []
    for g in groups:
        group_price_rows.append(
            {
                "Group": g,
                "Group Price per SQM (AUD)": existing_group_prices.get(g, np.nan),
            }
        )
    group_price_df = pd.DataFrame(group_price_rows)

    st.markdown("**Step 2 – Set group prices (per SQM, AUD)**")
    edited_group_price_df = st.data_editor(
        group_price_df,
        num_rows="dynamic",
        use_container_width=True,
        key="group_price_editor",
    )

    # Individual material override prices
    st.markdown("**Step 3 – Optional material overrides (per SQM, AUD)**")
    override_rows = []
    for m in materials:
        override_rows.append(
            {
                "Material": m,
                "Override Price per SQM (AUD)": existing_overrides.get(m, np.nan),
            }
        )
    override_df = pd.DataFrame(override_rows)
    edited_override_df = st.data_editor(
        override_df,
        num_rows="fixed",
        use_container_width=True,
        key="material_override_editor",
    )

    # Save latest presets into session_state
    st.session_state["group_assignments"] = dict(
        zip(edited_group_assign_df["Material"], edited_group_assign_df["Group"])
    )
    st.session_state["group_prices"] = dict(
        zip(edited_group_price_df["Group"], edited_group_price_df["Group Price per SQM (AUD)"])
    )
    st.session_state["material_overrides"] = dict(
        zip(edited_override_df["Material"], edited_override_df["Override Price per SQM (AUD)"])
    )

    # Build preset data from current state
    preset_data = {
        "group_assignments": st.session_state["group_assignments"],
        "group_prices": st.session_state["group_prices"],
        "material_overrides": st.session_state["material_overrides"],
    }

    st.markdown("**Preset saving options**")
    col_left, col_right = st.columns(2)

    with col_left:
        save_mode = st.radio(
            "How should these pricing changes be saved?",
            [
                "Only for this session (do not update JSON on server)",
                "Update default JSON on server (material_groups_default.json)",
            ],
            index=0,
            help=(
                "If you choose to update the default JSON, the app will try to overwrite "
                "'material_groups_default.json' in the current environment. On Streamlit Cloud "
                "this may not persist across deployments; use the download option to commit to GitHub."
            ),
        )

        if st.button("Apply save option now"):
            if save_mode.startswith("Update default JSON"):
                try:
                    with open("material_groups_default.json", "w", encoding="utf-8") as f:
                        json.dump(preset_data, f, indent=2)
                    st.success("Updated material_groups_default.json on server.")
                except Exception as e:
                    st.error(f"Could not update default JSON on server: {e}")
            else:
                st.info("Changes kept only in this session (JSON file not modified).")

    with col_right:
        preset_bytes = BytesIO(json.dumps(preset_data, indent=2).encode("utf-8"))
        st.download_button(
            "Download current material group preset (JSON)",
            data=preset_bytes,
            file_name="material_groups_preset.json",
            mime="application/json",
            help=(
                "Download the latest preset and commit it to your Git repo if you want it "
                "to be the new default next time you deploy."
            ),
        )

    # ---------- Apply pricing ----------

    st.subheader("Apply grouped pricing")

    group_assignment_map = st.session_state["group_assignments"]
    group_price_map = st.session_state["group_prices"]
    material_override_map = st.session_state["material_overrides"]

    calc_with_price = calc_df.copy()

    # Resolve base price per SQM (AUD) for each row:
    # 1) If material override exists, use that
    # 2) Else if group has a price, use group price
    # 3) Else NaN
    def resolve_base_price(material):
        override = material_override_map.get(material)
        if override is not None and not pd.isna(override):
            return override
        group = group_assignment_map.get(material)
        if group:
            gp = group_price_map.get(group)
            if gp is not None and not pd.isna(gp):
                return gp
        return np.nan

    calc_with_price["Base Price per SQM (AUD)"] = calc_with_price["Material"].apply(
        resolve_base_price
    )

    # Apply DS loading to get effective price (AUD)
    ds_factor = 1.0 + double_sided_loading_percent / 100.0
    calc_with_price["Effective Price per SQM (AUD)"] = calc_with_price.apply(
        lambda r: r["Base Price per SQM (AUD)"] * ds_factor if r.get("Side") == "DS" else r["Base Price per SQM (AUD)"],
        axis=1,
    )

    # Price calculations in AUD
    calc_with_price["Price per unit (AUD)"] = (
        calc_with_price["SQM per unit"] * calc_with_price["Effective Price per SQM (AUD)"]
    )
    calc_with_price["Price per annum (AUD)"] = (
        calc_with_price["SQM per annum"] * calc_with_price["Effective Price per SQM (AUD)"]
    )
    calc_with_price["Price per run (AUD)"] = (
        calc_with_price["SQM per run"] * calc_with_price["Effective Price per SQM (AUD)"]
    )

    # Optional converted currency
    if display_currency.upper() != "AUD" or abs(conversion_rate - 1.0) > 1e-9:
        calc_with_price[f"Base Price per SQM ({display_currency})"] = (
            calc_with_price["Base Price per SQM (AUD)"] * conversion_rate
        )
        calc_with_price[f"Effective Price per SQM ({display_currency})"] = (
            calc_with_price["Effective Price per SQM (AUD)"] * conversion_rate
        )
        calc_with_price[f"Price per unit ({display_currency})"] = (
            calc_with_price["Price per unit (AUD)"] * conversion_rate
        )
        calc_with_price[f"Price per annum ({display_currency})"] = (
            calc_with_price["Price per annum (AUD)"] * conversion_rate
        )
        calc_with_price[f"Price per run ({display_currency})"] = (
            calc_with_price["Price per run (AUD)"] * conversion_rate
        )

    # Round all numeric price-related columns to 2 decimals
    price_cols = [c for c in calc_with_price.columns if "Price" in c]
    for col in price_cols:
        calc_with_price[col] = calc_with_price[col].round(2)

    # Build a display copy with $ sign for price columns
    display_df = calc_with_price.copy()
    for col in price_cols:
        display_df[col] = display_df[col].apply(
            lambda x: "" if pd.isna(x) else f"${x:,.2f}"
        )

    st.subheader("Final calculation table (with grouped pricing, formatted)")
    st.dataframe(display_df)

    # Download calculated table (numeric, rounded) as Excel
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
