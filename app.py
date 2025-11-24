
import re
import numpy as np
import pandas as pd
import streamlit as st


# =========================================================
# Helper: Parse dimensions → m2
# =========================================================
def parse_area_m2(dimensions: str):
    if pd.isna(dimensions):
        return np.nan

    s = str(dimensions).lower().replace("×", "x")

    # Multi-panel pattern: "2 x 2547mm x 755mm"
    panel_pattern = r"(\d+)\s*x\s*(\d+)\s*mm\s*x\s*(\d+)\s*mm"
    matches = list(re.finditer(panel_pattern, s))

    if matches:
        total_mm2 = 0.0
        for m in matches:
            qty = float(m.group(1))
            w = float(m.group(2))
            h = float(m.group(3))
            total_mm2 += qty * w * h
        return total_mm2 / 1_000_000.0

    # Simple "841mm x 1189mm"
    simple = s.replace(" ", "").replace("mm", "")
    parts = simple.split("x")
    if len(parts) != 2:
        return np.nan
    try:
        w = float(parts[0])
        h = float(parts[1])
        return (w * h) / 1_000_000.0
    except:
        return np.nan


# =========================================================
# Helper: Sidedness detection
# =========================================================
def detect_sides_from_text(text):
    if pd.isna(text):
        return "Single Sided"

    s = str(text).lower()
    if "double" in s or "ds" in s:
        return "Double Sided"
    if "single" in s or "ss" in s:
        return "Single Sided"
    return "Single Sided"


# =========================================================
# Helper: Medium grouping
# =========================================================
def material_group_key(stock):
    if not isinstance(stock, str):
        return ""

    raw = stock.strip()
    s = raw.lower()

    # Thickness grouping
    m_thick = re.search(r"(\d+)\s*mm", s)
    if m_thick:
        t = m_thick.group(1)
        if "screenboard" in s:
            return f"{t}mm Screenboard"
        if "corflute" in s or "coreflute" in s:
            return f"{t}mm Corflute"
        if "acrylic" in s:
            return f"{t}mm Acrylic"
        if "pvc" in s:
            return f"{t}mm PVC"
        if "hips" in s:
            return f"{t}mm HIPS"
        if "acm" in s:
            return f"{t}mm ACM"

    # GSM grouping
    m_gsm = re.search(r"(\d{3})\s*gsm", s)
    if m_gsm:
        gsm = m_gsm.group(1)
        if "silk" in s or "satin" in s:
            return f"{gsm}gsm Silk/Satin"
        if "matt" in s:
            return f"{gsm}gsm Matt"
        if "gloss" in s:
            return f"{gsm}gsm Gloss"
        if "synthetic" in s or "plasnet" in s:
            return f"{gsm}gsm Synthetic"
        return f"{gsm}gsm Paper/Card"

    # SAV
    if "sav" in s or "vinyl" in s:
        return "SAV / Vinyl"

    # Fallback
    cleaned = re.sub(r"[^a-z0-9]+", " ", s)
    tokens = cleaned.split()
    if len(tokens) >= 2:
        return " ".join(tokens[:2])
    if tokens:
        return tokens[0]
    return raw


# =========================================================
# Helper: Excel column letters (A,B…)
# =========================================================
def num_to_col(n):
    result = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        result = chr(ord("A") + r) + result
    return result


# =========================================================
# Helper: Money format
# =========================================================
def fmt_money(val):
    try:
        return f"${float(val):,.2f}"
    except:
        return ""


# =========================================================
# Helper: Tier pricing
# =========================================================
def get_tiered_rate(qty, tiers):
    if pd.isna(qty):
        return 0.0
    for t in tiers:
        lo, hi, price = t["min"], t["max"], t["price"]
        if lo is None and qty <= hi:
            return price
        if hi is None and qty >= lo:
            return price
        if lo is not None and hi is not None and lo <= qty <= hi:
            return price
    return 0.0


# =========================================================
# STREAMLIT APP
# =========================================================
st.set_page_config(page_title="Tender SQM Mapping Wizard v13.5", layout="wide")
st.title("Tender SQM Mapping Wizard – v13.5 (Excel-style view)")


# ---------------------------------------------------------
# 1. Upload
# ---------------------------------------------------------
file = st.file_uploader("Upload Excel file", ["xlsx", "xls"])
if not file:
    st.stop()

xls = pd.ExcelFile(file)

# ---------------------------------------------------------
# 2. Sheet selection
# ---------------------------------------------------------
sheet = st.selectbox("Select sheet", xls.sheet_names)
df_raw = xls.parse(sheet)

# ---------------------------------------------------------
# 3. Raw Preview (Excel-style)
# ---------------------------------------------------------
st.subheader("Raw Sheet Preview (Excel-style)")

raw = df_raw.copy()
raw.index = range(1, len(raw) + 1)
raw.columns = [num_to_col(i + 1) for i in range(len(raw.columns))]

with st.expander("Raw sheet view options", expanded=True):
    rs = st.number_input("Start row", 1, len(raw), 1)
    re_ = st.number_input("End row", rs, len(raw), min(len(raw), rs + 50))
    show_cols = st.multiselect("Columns to display", list(raw.columns), list(raw.columns))

st.dataframe(raw.loc[rs:re_, show_cols], use_container_width=True)

# ---------------------------------------------------------
# 4. Auto-detect orientation
# ---------------------------------------------------------
r, c = df_raw.shape
auto = "Row Based" if r >= 2 * c else "Column Based"

layout = st.radio(
    "Is each item a row or a column?",
    ["Row Based", "Column Based"],
    index=0 if auto == "Row Based" else 1,
)

df = df_raw.copy() if layout == "Row Based" else df_raw.T.copy()

# ---------------------------------------------------------
# 5. Normalised preview
# ---------------------------------------------------------
st.subheader("Normalised Preview (Excel-style)")

norm = df.copy()
norm.index = range(1, len(norm) + 1)
norm_letters = [num_to_col(i + 1) for i in range(len(norm.columns))]
norm.columns = norm_letters

with st.expander("Normalised view options", expanded=True):
    ns = st.number_input("Normalised start row", 1, len(norm), 1, key="ns")
    ne = st.number_input("Normalised end row", ns, len(norm), min(len(norm), ns + 50), key="ne")
    ncols = st.multiselect("Normalised columns to display", list(norm.columns), list(norm.columns), key="ncols")

st.dataframe(norm.loc[ns:ne, ncols], use_container_width=True)

# ---------------------------------------------------------
# 6. Mapping
# ---------------------------------------------------------
st.subheader("Column Mapping")

df_cols = list(df.columns)
excel_to_real = dict(zip(norm_letters, df_cols))
labelled = [f"{ltr} – {col}" for ltr, col in zip(norm_letters, df_cols)]

def pick(col_label):
    letter = col_label.split("–")[0].strip()
    return excel_to_real.get(letter)

c1, c2 = st.columns(2)
with c1:
    mcol = st.selectbox("Material column", labelled)
    scol = st.selectbox("Size column", labelled)
    qcol = st.selectbox("Quantity column", labelled)
with c2:
    side_col = st.selectbox("Side column (optional)", ["<none>"] + labelled)
    run_col = st.selectbox("Runs per annum (optional)", ["<none>"] + labelled)
    perrun_col = st.selectbox("Per-run Qty (optional)", ["<none>"] + labelled)

c3, c4 = st.columns(2)
with c3:
    lot_col = st.selectbox("Lot ID (optional)", ["<none>"] + labelled)
with c4:
    desc_col = st.selectbox("Description (optional)", ["<none>"] + labelled)

if not st.button("Apply Mapping"):
    st.stop()

material = pick(mcol)
size = pick(scol)
qty = pick(qcol)
side = pick(side_col) if side_col != "<none>" else None
runs = pick(run_col) if run_col != "<none>" else None
per_run = pick(perrun_col) if perrun_col != "<none>" else None
lot = pick(lot_col) if lot_col != "<none>" else None
desc = pick(desc_col) if desc_col != "<none>" else None

# ---------------------------------------------------------
# 7. Build cleaned dataset
# ---------------------------------------------------------
data = pd.DataFrame()
data["Material"] = df[material]
data["Size"] = df[size]
data["Qty"] = pd.to_numeric(df[qty], errors="coerce")

if lot:
    data["Lot"] = df[lot]
if desc:
    data["Description"] = df[desc]

if runs:
    data["Runs"] = pd.to_numeric(df[runs], errors="coerce")
else:
    data["Runs"] = np.nan

if per_run:
    data["Qty per Run"] = pd.to_numeric(df[per_run], errors="coerce")
else:
    data["Qty per Run"] = np.nan

if side:
    data["Side (auto)"] = df[side].apply(detect_sides_from_text)
else:
    data["Side (auto)"] = data["Material"].apply(detect_sides_from_text)

data["DoubleSided"] = data["Side (auto)"].eq("Double Sided")

data["Area_each"] = data["Size"].apply(parse_area_m2)
data["Area_total"] = data["Area_each"] * data["Qty"]

if runs:
    safe_runs = data["Runs"].replace(0, np.nan)
    data["Area_per_Run"] = data["Area_total"] / safe_runs
else:
    data["Area_per_Run"] = np.nan

# ---------------------------------------------------------
# 8. Grouping
# ---------------------------------------------------------
st.subheader("Material Grouping")

materials = sorted(set(data["Material"].dropna()))

if "groups" not in st.session_state:
    st.session_state.groups = pd.DataFrame({
        "Material": materials,
        "InitialGroup": [material_group_key(s) for s in materials]
    })
    st.session_state.groups["Group"] = st.session_state.groups["InitialGroup"]

else:
    g = st.session_state.groups
    existing = set(g["Material"])
    new_items = [m for m in materials if m not in existing]
    if new_items:
        new_df = pd.DataFrame({
            "Material": new_items,
            "InitialGroup": [material_group_key(s) for s in new_items]
        })
        new_df["Group"] = new_df["InitialGroup"]
        g = pd.concat([g, new_df], ignore_index=True)
    g = g[g["Material"].isin(materials)].reset_index(drop=True)
    st.session_state.groups = g

groups = st.session_state.groups

st.markdown(
    "- **InitialGroup** is auto-derived from text\n"
    "- **Group** controls pricing\n"
    "- Edit **Group** to merge/split materials"
)

groups = st.data_editor(
    groups,
    num_rows="fixed",
    use_container_width=True,
    column_config={
        "Material": st.column_config.TextColumn(disabled=True),
        "InitialGroup": st.column_config.TextColumn(disabled=True)
    },
    key="group_editor"
)

mapping = dict(zip(groups["Material"], groups["Group"]))
data["Group"] = data["Material"].map(mapping)

# ---------------------------------------------------------
# 9. Edit Double-Sided
# ---------------------------------------------------------
st.subheader("Double-Sided Overrides")

edit_cols = ["Material", "Size", "Qty", "Group", "DoubleSided"]
if "Lot" in data.columns:
    edit_cols.insert(0, "Lot")
if "Description" in data.columns:
    edit_cols.insert(1, "Description")

edited = st.data_editor(
    data[edit_cols],
    use_container_width=True,
    column_config={"DoubleSided": st.column_config.CheckboxColumn()},
    key="ds_editor"
)

data["DoubleSided"] = edited["DoubleSided"]

# ---------------------------------------------------------
# 10. Pricing
# ---------------------------------------------------------
st.subheader("Pricing")

st.sidebar.header("Pricing Controls")
ds_pct = st.sidebar.number_input("Double-sided loading %", 0.0, 200.0, 25.0)

t1_max = st.sidebar.number_input("Tier1 max Qty", 1, 999999, 100)
t1_price = st.sidebar.number_input("Tier1 price per m²", 0.0, 999.0, 10.0)

t2_max = st.sidebar.number_input("Tier2 max Qty", t1_max, 999999, 1000)
t2_price = st.sidebar.number_input("Tier2 price per m²", 0.0, 999.0, 8.0)

t3_price = st.sidebar.number_input("Tier3 price per m²", 0.0, 999.0, 6.0)

tiers = [
    {"min": None, "max": t1_max, "price": t1_price},
    {"min": t1_max + 1, "max": t2_max, "price": t2_price},
    {"min": t2_max + 1, "max": None, "price": t3_price},
]

data["Rate"] = data["Qty"].apply(lambda x: get_tiered_rate(x, tiers))
mult = 1 + ds_pct / 100
data["Multiplier"] = np.where(data["DoubleSided"], mult, 1.0)

data["Value"] = data["Area_total"] * data["Rate"] * data["Multiplier"]

if "Runs" in data.columns:
    safe_runs = data["Runs"].replace(0, np.nan)
    data["Value_per_Run"] = data["Value"] / safe_runs
else:
    data["Value_per_Run"] = np.nan

# ---------------------------------------------------------
# 11. Final preview
# ---------------------------------------------------------
st.subheader("Final Output")

out_cols = [
    "Material",
    "Size",
    "Qty",
    "Group",
    "DoubleSided",
    "Area_each",
    "Area_total",
    "Rate",
    "Multiplier",
    "Value"
]

if "Runs" in data.columns:
    out_cols.insert(5, "Runs")
    out_cols.insert(6, "Area_per_Run")
    out_cols.append("Value_per_Run")

if "Lot" in data.columns:
    out_cols.insert(0, "Lot")
if "Description" in data.columns:
    out_cols.insert(1, "Description")

out = data[out_cols].copy()
out["Rate"] = out["Rate"].apply(fmt_money)
out["Value"] = out["Value"].apply(fmt_money)
if "Value_per_Run" in out.columns:
    out["Value_per_Run"] = out["Value_per_Run"].apply(fmt_money)

st.dataframe(out, use_container_width=True)

# Totals
st.metric("Total m² per annum", f"{data['Area_total'].sum():,.2f}")
st.metric("Total Value (ex GST)", fmt_money(data["Value"].sum()))
