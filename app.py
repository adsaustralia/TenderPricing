import re
import io
import numpy as np
import pandas as pd
import streamlit as st


# ---------------------------------------------------------
# Helper functions
# ---------------------------------------------------------


def parse_area_m2(dimensions: str):
    """Convert '841mm x 1189mm' → m² (assumes mm × mm)."""
    if pd.isna(dimensions):
        return np.nan

    s = (
        str(dimensions)
        .lower()
        .replace(" ", "")
        .replace("mm", "")
        .replace("×", "x")
    )
    parts = s.split("x")
    if len(parts) != 2:
        return np.nan

    try:
        w = float(parts[0])
        h = float(parts[1])
    except Exception:
        return np.nan

    return (w * h) / 1_000_000.0


def extract_stock_name(spec: str):
    """Take text before first comma as stock name."""
    if pd.isna(spec):
        return ""
    return str(spec).split(",")[0].strip()


def detect_sides(spec: str):
    """Detect single/double sided from text."""
    if pd.isna(spec):
        return "Single Sided"

    s = str(spec).lower()
    if "double" in s:
        return "Double Sided"
    return "Single Sided"


def _first_match(pattern: str, text: str):
    """Safe regex helper that works with or without capture groups."""
    m = re.search(pattern, text)
    if not m:
        return None
    try:
        return m.group(1)
    except IndexError:
        return m.group(0)


# ---------------------------------------------------------
# Option B – material grouping logic
# ---------------------------------------------------------


def material_group_key_medium(stock: str) -> str:
    """Derive a medium-detail material group key from a stock name (Option B).

    - Boards: thickness + substrate (e.g. '2mm Screenboard', '3mm Corflute')
    - Papers: gsm + finish (e.g. '300gsm Silk/Satin')
    - Vinyls/SAV: brand + code (e.g. 'SAV – Avery MPI 2126')
    - Special synthetics/films: named buckets (Duratran, Jellyfish, etc.)
    """
    if not isinstance(stock, str):
        return ""

    s_raw = stock
    s = stock.lower()

    # --------- Rigid boards (mm + substrate) ---------
    thickness = _first_match(r"(\d+)\s*mm", s)
    if thickness:
        if "screenboard" in s or "screen board" in s:
            return f"{thickness}mm Screenboard"
        if "corflute" in s or "coreflute" in s:
            return f"{thickness}mm Corflute"
        if "acrylic" in s:
            return f"{thickness}mm Acrylic"
        if "pvc" in s:
            return f"{thickness}mm PVC"
        if "hips" in s:
            return f"{thickness}mm HIPS"
        if "acm" in s:
            return f"{thickness}mm ACM"
        if "aluminium" in s or "aluminum" in s:
            return f"{thickness}mm Aluminium"
        if "maxi t" in s or "maxi-t" in s:
            return f"{thickness}mm Maxi-T"

    # --------- Special rigid / panels without mm ---------
    if "braille acrylic" in s:
        return "Braille Acrylic Panel"
    if "anodised aluminium" in s or "anodized aluminum" in s:
        return "Aluminium Panel"

    # --------- Backlit / lightbox films ---------
    if "duratran" in s or "backlit" in s:
        return "Backlit Film – Duratran"

    # --------- Jellyfish / Yuppo / synthetic papers ---------
    if "jellyfish" in s and "supercling" in s:
        return "Synthetic – Jellyfish Supercling"
    if "yuppo" in s:
        return "Synthetic – Yuppo"
    if "synthetic" in s and "plasnet" in s:
        return "283gsm Synthetic – Plasnet"

    # --------- Paper / card (gsm + finish) ---------
    gsm = _first_match(r"(\d{3})\s*gsm", s)
    if gsm:
        if "silk" in s or "satin" in s:
            return f"{gsm}gsm Silk/Satin"
        if "ecomatt" in s or "matt" in s:
            return f"{gsm}gsm Matt"
        if "gloss" in s:
            return f"{gsm}gsm Gloss"
        if "synthetic" in s or "plasnet" in s:
            return f"{gsm}gsm Synthetic"
        return f"{gsm}gsm Paper/Card"

    # --------- Vinyls / SAV by brand and code ---------
    # Avery / MPI
    if "avery" in s or "mpi" in s:
        code = _first_match(r"\b(11\d{2}|21\d{2}|29\d{2}|33\d{2})\b", s)
        if not code:
            code = _first_match(r"\b\d{3,4}\b", s)
        brand = "Avery MPI" if "mpi" in s else "Avery"
        if code:
            return f"SAV – {brand} {code}"
        return f"SAV – {brand}"

    # Arlon
    if "arlon" in s:
        code = _first_match(r"\b\d{3,4}\b", s)
        if code:
            return f"SAV – Arlon {code}"
        return "SAV – Arlon"

    # Mactac
    if "mactac" in s and "glass decor" in s:
        return "Glass Decor – Mactac"
    if "mactac" in s:
        return "SAV – Mactac"

    # 3M / Metamark / Hexis / generic SAV
    if "3m" in s:
        code = _first_match(r"\b\d{3,4}\b", s)
        if code:
            return f"SAV – 3M {code}"
        return "SAV – 3M"

    if "metamark" in s:
        return "SAV – Metamark"
    if "hexis" in s:
        return "SAV – Hexis"

    # Numeric SAV families (2126, 2903, 2904, 3302, 2105, etc.)
    if "sav" in s:
        code = _first_match(r"\b(2126|2903|2904|3302|2105)\b", s)
        if code:
            return f"SAV – {code} Family"
        return "SAV – Other"

    # Glass / frosted / clear
    if "glass decor" in s or "frosted" in s or "dusted" in s:
        return "Glass Decor / Frosted Film"
    if "ultra clear" in s:
        return "Clear Window Film – Ultra Clear"

    # Black CCV
    if "ccv" in s and "black" in s:
        return "SAV – Black CCV"

    # ---------- Fallback: cleaned first two tokens ----------
    cleaned = re.sub(r"\(.*?\)", "", s)
    cleaned = re.sub(r"[^a-z0-9]+", " ", cleaned)
    cleaned = cleaned.strip()
    tokens = cleaned.split()
    if len(tokens) >= 2:
        return " ".join(tokens[:2])
    elif tokens:
        return tokens[0]
    return s_raw.strip()


def friendly_group_name(group_key: str) -> str:
    """Nicer label for a group key."""
    if not isinstance(group_key, str):
        return ""

    g = group_key
    if g.startswith("SAV – "):
        core = g.replace("SAV – ", "").strip()
        return f"{core} Vinyl"
    if g.startswith("Synthetic – "):
        return g.replace("Synthetic – ", "Synthetic ")
    if "Backlit Film" in g:
        return g.replace("Backlit Film –", "Backlit Film ").strip()
    return g


# ---------------------------------------------------------
# Streamlit app
# ---------------------------------------------------------


st.set_page_config(layout="wide", page_title="BP Tender SQM Calculator v7")
st.title("BP Tender – Square Metre Calculator (v7)")
st.caption("Option B grouping + search, merge groups, group preview, double-sided control")

uploaded = st.file_uploader("Upload tender Excel", type=["xlsx", "xls"])
if not uploaded:
    st.info("Please upload an Excel file with at least: Dimensions, Print/Stock Specifications, Total Annual Volume.")
    st.stop()

df = pd.read_excel(uploaded)

required_cols = ["Dimensions", "Print/Stock Specifications", "Total Annual Volume"]
missing = [c for c in required_cols if c not in df.columns]
if missing:
    st.error(f"Missing required columns: {missing}")
    st.stop()

data = df.copy()

# Base calculations
data["Area m² (each)"] = data["Dimensions"].apply(parse_area_m2)
data["Stock Name"] = data["Print/Stock Specifications"].apply(extract_stock_name)
data["Sided (auto)"] = data["Print/Stock Specifications"].apply(detect_sides)
data["Double Sided?"] = data["Sided (auto)"] == "Double Sided"
data["Quantity"] = data["Total Annual Volume"]
data["Total Area m²"] = data["Area m² (each)"] * data["Quantity"]

# ---------------------------------------------------------
# Step 1 – Double-sided overrides
# ---------------------------------------------------------

st.subheader("Step 1 – Review Double-Sided Flags")

ds_cols = [
    "Dimensions",
    "Print/Stock Specifications",
    "Quantity",
    "Area m² (each)",
    "Total Area m²",
    "Double Sided?",
]
if "Lot ID" in data.columns:
    ds_cols.insert(0, "Lot ID")
if "Item Description" in data.columns:
    ds_cols.insert(1, "Item Description")

edited_ds = st.data_editor(
    data[ds_cols],
    use_container_width=True,
    num_rows="dynamic",
    column_config={
        "Double Sided?": st.column_config.CheckboxColumn(
            "Double Sided?", help="Tick if the item is double-sided."
        )
    },
    key="double_sided_editor",
)

data["Double Sided?"] = edited_ds["Double Sided?"].fillna(False)

# ---------------------------------------------------------
# Step 2 – Material grouping (Option B)
# ---------------------------------------------------------

st.subheader("Step 2 – Group Materials for Pricing (Option B)")

unique_stocks = sorted(s for s in data["Stock Name"].dropna().unique() if str(s).strip())

if "groups_df" not in st.session_state:
    st.session_state["groups_df"] = pd.DataFrame(
        {
            "Stock Name": unique_stocks,
            "Initial Group": [material_group_key_medium(s) for s in unique_stocks],
        }
    )
    st.session_state["groups_df"]["Assigned Group"] = st.session_state["groups_df"]["Initial Group"]
else:
    gdf = st.session_state["groups_df"]
    existing = set(gdf["Stock Name"])
    new_stocks = [s for s in unique_stocks if s not in existing]
    if new_stocks:
        new_rows = pd.DataFrame(
            {
                "Stock Name": new_stocks,
                "Initial Group": [material_group_key_medium(s) for s in new_stocks],
            }
        )
        new_rows["Assigned Group"] = new_rows["Initial Group"]
        gdf = pd.concat([gdf, new_rows], ignore_index=True)
    gdf = gdf[gdf["Stock Name"].isin(unique_stocks)].reset_index(drop=True)
    st.session_state["groups_df"] = gdf

groups_df = st.session_state["groups_df"]

st.markdown(
    """- **Initial Group** is auto-generated based on thickness, substrate, GSM, or SAV brand/code.

- **Assigned Group** is what actually controls pricing.

- Give multiple stocks the same Assigned Group to price them together."""
)

# Search bar (read-only view)
search_term = st.text_input("Search stock/groups (read-only view)", value="").lower().strip()
if search_term:
    filtered_view = groups_df[
        groups_df.apply(
            lambda r: search_term in r["Stock Name"].lower()
            or search_term in r["Initial Group"].lower()
            or search_term in r["Assigned Group"].lower(),
            axis=1,
        )
    ]
    st.dataframe(filtered_view, use_container_width=True, height=250)
else:
    st.dataframe(groups_df, use_container_width=True, height=250)

st.markdown("#### Edit Assigned Groups")

assigned_options = sorted(groups_df["Assigned Group"].unique())
editable_groups = st.data_editor(
    groups_df,
    use_container_width=True,
    num_rows="fixed",
    column_config={
        "Stock Name": st.column_config.TextColumn(disabled=True),
        "Initial Group": st.column_config.TextColumn(disabled=True),
        "Assigned Group": st.column_config.SelectboxColumn(
            "Assigned Group",
            options=assigned_options,
            help="Choose which group this stock belongs to.",
        ),
    },
    key="groups_editor",
)

st.session_state["groups_df"] = editable_groups
groups_df = editable_groups

# Merge groups
st.markdown("### Merge Groups")
all_assigned = sorted(groups_df["Assigned Group"].unique())
merge_selection = st.multiselect("Groups to merge", all_assigned)
merge_target = st.text_input("Merged group name", value=merge_selection[0] if merge_selection else "")

if st.button("Merge selected groups"):
    if merge_selection and merge_target:
        mask = groups_df["Assigned Group"].isin(merge_selection)
        groups_df.loc[mask, "Assigned Group"] = merge_target
        st.session_state["groups_df"] = groups_df
        st.success(f"Merged {len(merge_selection)} groups into '{merge_target}'.")

# Apply group mapping to main data
stock_to_group = dict(zip(groups_df["Stock Name"], groups_df["Assigned Group"]))
data["Material Group"] = data["Stock Name"].map(stock_to_group).fillna("Unassigned")

# ---------------------------------------------------------
# Step 2.5 – Group preview panel
# ---------------------------------------------------------

st.subheader("Group Preview")

group_summary = (
    data.groupby("Material Group")
    .agg(
        Materials=("Stock Name", "nunique"),
        Lines=("Stock Name", "count"),
        Total_Area_m2=("Total Area m²", "sum"),
    )
    .reset_index()
)

group_summary["Friendly Name"] = group_summary["Material Group"].apply(friendly_group_name)

group_summary = group_summary[
    ["Material Group", "Friendly Name", "Materials", "Lines", "Total_Area_m2"]
]

st.dataframe(group_summary, use_container_width=True)

# ---------------------------------------------------------
# Step 3 – Pricing per material group
# ---------------------------------------------------------

st.sidebar.header("Pricing & Double-Sided Loading")

group_names = sorted(g for g in data["Material Group"].dropna().unique() if str(g).strip())
group_prices = {}

st.sidebar.subheader("Price per m² by material group")

for g in group_names:
    group_prices[g] = st.sidebar.number_input(
        g,
        min_value=0.0,
        value=0.0,
        step=0.1,
        key=f"price_{g}",
    )

double_loading_pct = st.sidebar.number_input(
    "Double-sided loading (%)",
    min_value=0.0,
    value=25.0,
    step=1.0,
    help="Extra percentage added for double-sided lines.",
)

data["Price per m²"] = data["Material Group"].map(group_prices).fillna(0.0)
double_mult = 1.0 + double_loading_pct / 100.0
data["Sided Multiplier"] = np.where(data["Double Sided?"], double_mult, 1.0)

data["Line Value (ex GST)"] = (
    data["Total Area m²"] * data["Price per m²"] * data["Sided Multiplier"]
)

# ---------------------------------------------------------
# Step 4 – Final view & download
# ---------------------------------------------------------

st.subheader("Step 4 – Final Calculated Lines")

data["Friendly Group Name"] = data["Material Group"].apply(friendly_group_name)

pricing_cols = [
    "Stock Name",
    "Material Group",
    "Friendly Group Name",
    "Dimensions",
    "Quantity",
    "Total Area m²",
    "Double Sided?",
    "Price per m²",
    "Sided Multiplier",
    "Line Value (ex GST)",
]
if "Lot ID" in data.columns:
    pricing_cols.insert(0, "Lot ID")
if "Item Description" in data.columns:
    pricing_cols.insert(1, "Item Description")

st.dataframe(data[pricing_cols], use_container_width=True)

total_area = data["Total Area m²"].sum(skipna=True)
total_value = data["Line Value (ex GST)"].sum(skipna=True)

c1, c2 = st.columns(2)
c1.metric("Total Area (m²)", f"{total_area:,.2f}")
c2.metric("Total Value (ex GST)", f"${total_value:,.2f}")

# Excel export
buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
    data.to_excel(writer, index=False, sheet_name="Priced Tender")
    group_summary.to_excel(writer, index=False, sheet_name="Group Summary")

st.download_button(
    "Download priced tender as Excel",
    data=buffer.getvalue(),
    file_name="bp_tender_priced.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
