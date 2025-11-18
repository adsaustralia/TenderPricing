import re
import numpy as np
import pandas as pd
import streamlit as st


# ---------- Helpers ----------

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
    """Detect single/double sided from text. Default to Single Sided."""
    if pd.isna(spec) or str(spec).strip() == "":
        return "Single Sided"

    s = str(spec).lower()
    if "double" in s:
        return "Double Sided"
    return "Single Sided"


def _first_match(pattern: str, text: str):
    m = re.search(pattern, text)
    return m.group(1) if m else None


def material_group_key_medium(stock: str) -> str:
    """Derive a medium-detail material group key from a stock name (Option B)."""
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
        if "maxi t" in s or "maxi t" in s.replace("-", " "):
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
        if "plasnet" in s or "synthetic" in s:
            return f"{gsm}gsm Synthetic"
        return f"{gsm}gsm Paper/Card"

    # --------- Vinyls / SAV by brand and code ---------
    # Avery
    if "avery" in s or "mpi" in s:
        code = _first_match(r"\b(11\d{2}|21\d{2}|29\d{2}|33\d{2})\b", s)
        if not code:
            code = _first_match(r"\b\d{3,4}\b", s)
        if "mpi" in s:
            brand = "Avery MPI"
        else:
            brand = "Avery"
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
    if "hexis" in s or "hex is" in s:
        return "SAV – Hexis"

    # Easy Apply / HiTac / SAV generic families
    if "sav" in s:
        code = _first_match(r"\b(2126|2903|2904|3302|2105)\b", s)
        if code:
            return f"SAV – {code} Family"
        return "SAV – Other"

    # Frosted / dusted glass
    if "glass decor" in s or "frosted" in s or "dusted" in s:
        return "Glass Decor / Frosted Film"

    # Ultra clear / clear window films
    if "ultra clear" in s:
        return "Clear Window Film – Ultra Clear"

    # Catch-all matte/black CCV etc.
    if "ccv" in s and "black" in s:
        return "SAV – Black CCV"

    # ---------- Fallback: cleaned first two tokens ----------
    cleaned = re.sub(r"\(.*?\)", "", s)
    cleaned = re.sub(r"[^a-z0-9]+", " ", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    tokens = cleaned.split(" ")
    if len(tokens) >= 2:
        return " ".join(tokens[:2])
    elif tokens:
        return tokens[0]
    return s_raw.strip()


def friendly_group_name(group_key: str) -> str:
    """Turn a technical group key into a slightly nicer label for display/download."""
    if not isinstance(group_key, str):
        return ""

    g = group_key

    # Common replacements just for nicer labels
    g = g.replace("SAV –", "").strip()
    g = g.replace("Synthetic –", "Synthetic ")
    g = g.replace("Backlit Film –", "Backlit Film ")
    g = g.replace("Glass Decor –", "Glass Decor ")
    g = g.replace(" –", " -")

    # Example: 'Avery MPI 2904' -> 'Avery 2904 Vinyl'
    m = re.match(r"SAV – (Avery MPI\s+\d+)", group_key)
    if m:
        core = m.group(1)
        return f"{core} Vinyl"

    if group_key.startswith("SAV – Avery MPI"):
        core = group_key.replace("SAV – ", "")
        return f"{core} Vinyl"

    if group_key.startswith("SAV –"):
        # 'SAV – Arlon 8000' -> 'Arlon 8000 Vinyl'
        core = group_key.replace("SAV – ", "")
        return f"{core} Vinyl"

    return g


# ---------- Streamlit App ----------

st.set_page_config(page_title="BP Tender SQM Calculator", layout="wide")
st.title("BP Tender – Square Metre Price Calculator (Material Groups – Option B + Tools)")


uploaded_file = st.file_uploader("Upload the tender Excel file", type=["xlsx", "xls"])

if not uploaded_file:
    st.info("Upload the Excel file to continue.")
    st.stop()

df = pd.read_excel(uploaded_file)

required = ["Dimensions", "Print/Stock Specifications", "Total Annual Volume"]
missing = [c for c in required if c not in df.columns]
if missing:
    st.error(f"Missing required columns: {missing}")
    st.stop()

# ---------- Base processing ----------

data = df.copy()
data["Area m² (each)"] = data["Dimensions"].apply(parse_area_m2)
data["Stock Name"] = data["Print/Stock Specifications"].apply(extract_stock_name)
data["Sided (auto)"] = data["Print/Stock Specifications"].apply(detect_sides)
data["Double Sided?"] = data["Sided (auto)"] == "Double Sided"
data["Quantity"] = data["Total Annual Volume"]
data["Total Area m²"] = data["Area m² (each)"] * data["Quantity"]

# ---------- Step 1: Double-sided control ----------

st.subheader("Step 1 – Review & adjust double-sided lines")

base_cols = [
    "Dimensions",
    "Print/Stock Specifications",
    "Quantity",
    "Area m² (each)",
    "Total Area m²",
    "Double Sided?",
]

if "Lot ID" in data.columns:
    base_cols.insert(0, "Lot ID")
if "Item Description" in data.columns:
    base_cols.insert(1, "Item Description")

double_sided_view = st.data_editor(
    data[base_cols],
    use_container_width=True,
    num_rows="dynamic",
    column_config={
        "Double Sided?": st.column_config.CheckboxColumn(
            "Double Sided?",
            help="Tick to apply double-sided loading on this line.",
            default=False,
        )
    },
    key="double_sided_editor",
)

data["Double Sided?"] = double_sided_view["Double Sided?"].fillna(False)

st.caption("Tick/untick **Double Sided?** to enable/disable double-sided pricing per line.")

# ---------- Step 2: Material grouping with search + dropdown + merge ----------

st.subheader("Step 2 – Group similar materials for pricing (Option B)")

unique_stocks = sorted(
    s for s in data["Stock Name"].dropna().unique() if str(s).strip()
)

if "stock_materials_df" not in st.session_state:
    initial_groups = [material_group_key_medium(s) for s in unique_stocks]
    st.session_state["stock_materials_df"] = pd.DataFrame(
        {
            "Stock Name": unique_stocks,
            "Initial Group": initial_groups,
            "Assigned Group": initial_groups,  # start in the auto group
        }
    )
else:
    sm = st.session_state["stock_materials_df"]

    existing = set(sm["Stock Name"])
    new_stocks = [s for s in unique_stocks if s not in existing]
    if new_stocks:
        new_rows = pd.DataFrame(
            {
                "Stock Name": new_stocks,
                "Initial Group": [material_group_key_medium(s) for s in new_stocks],
            }
        )
        new_rows["Assigned Group"] = new_rows["Initial Group"]
        sm = pd.concat([sm, new_rows], ignore_index=True)

    sm = sm[sm["Stock Name"].isin(unique_stocks)].reset_index(drop=True)
    st.session_state["stock_materials_df"] = sm

stock_materials_df = st.session_state["stock_materials_df"]

st.markdown(
    """The app has auto-grouped materials using **thickness + substrate** and,
for vinyls/SAV, **brand + code** (Option B).

- **Assigned Group** controls pricing
- Multiple stocks with the same Assigned Group share the same price per m²
- You can reassign any stock via the dropdown
"""
)

# --- Search bar for quickly finding stocks ---
search_term = st.text_input("Search stock names / groups", value="").lower().strip()

filtered_df = stock_materials_df.copy()
if search_term:
    mask = (
        filtered_df["Stock Name"].str.lower().str.contains(search_term)
        | filtered_df["Initial Group"].str.lower().str.contains(search_term)
        | filtered_df["Assigned Group"].str.lower().str.contains(search_term)
    )
    filtered_df = filtered_df[mask]

# --- Group options (for dropdowns & merging) ---
group_options = sorted(
    set(stock_materials_df["Initial Group"]).union(stock_materials_df["Assigned Group"])
)

# --- Editable table with dropdown grouping ---
edited_df = st.data_editor(
    filtered_df,
    use_container_width=True,
    num_rows="fixed",
    column_config={
        "Stock Name": st.column_config.TextColumn(disabled=True),
        "Initial Group": st.column_config.TextColumn(disabled=True),
        "Assigned Group": st.column_config.SelectboxColumn(
            "Material Group",
            help="Choose which material group this stock belongs to.",
            options=group_options,
        ),
    },
    key="stock_group_editor",
)

# Write back edits into the master stock_materials_df
for idx, row in edited_df.iterrows():
    stock = row["Stock Name"]
    st.session_state["stock_materials_df"].loc[
        st.session_state["stock_materials_df"]["Stock Name"] == stock, "Assigned Group"
    ] = row["Assigned Group"]

stock_materials_df = st.session_state["stock_materials_df"]

# --- Merge groups tool ---
st.markdown("### Merge groups")

merge_cols = st.columns([2, 1])
with merge_cols[0]:
    groups_to_merge = st.multiselect(
        "Select two or more groups to merge",
        options=sorted(stock_materials_df["Assigned Group"].unique()),
    )
with merge_cols[1]:
    target_name = st.text_input(
        "Merged group name",
        value=groups_to_merge[0] if groups_to_merge else "",
        help="Name of the group after merge (e.g. 'SAV – Avery MPI 2126').",
    )

do_merge = st.button("Merge selected groups into target")

if do_merge and groups_to_merge and target_name:
    mask = stock_materials_df["Assigned Group"].isin(groups_to_merge)
    stock_materials_df.loc[mask, "Assigned Group"] = target_name
    st.session_state["stock_materials_df"] = stock_materials_df
    st.success(f"Merged {len(groups_to_merge)} groups into '{target_name}'.")

# Apply group mapping to main data
stock_to_group = dict(
    zip(stock_materials_df["Stock Name"], stock_materials_df["Assigned Group"])
)

data["Material Group"] = data["Stock Name"].map(stock_to_group).fillna("Unassigned")

# ---------- Step 2.5: Group Preview Panel ----------

st.subheader("Group Preview")

group_summary = (
    data.groupby("Material Group")
    .agg(
        Materials=("Stock Name", lambda x: x.nunique()),
        Lines=("Stock Name", "count"),
        Total_Area_m2=("Total Area m²", "sum"),
    )
    .reset_index()
)

group_summary["Friendly Name"] = group_summary["Material Group"].apply(friendly_group_name)

# Order columns nicely
group_summary = group_summary[
    ["Material Group", "Friendly Name", "Materials", "Lines", "Total_Area_m2"]
]

st.dataframe(group_summary, use_container_width=True)

# ---------- Step 3: Pricing per material group ----------

st.sidebar.header("Pricing & double-sided loading")

group_names = sorted(
    g for g in data["Material Group"].dropna().unique() if str(g).strip()
)

group_prices = {}
st.sidebar.subheader("Price per m² by material group")

for group in group_names:
    group_prices[group] = st.sidebar.number_input(
        f"{group}",
        min_value=0.0,
        value=0.0,
        step=0.10,
        key=f"price_group_{group}",
    )

double_loading_pct = st.sidebar.number_input(
    "Double-sided loading (%)",
    min_value=0.0,
    value=25.0,
    step=1.0,
    help="Extra percentage applied to double-sided lines.",
)

data["Price per m²"] = data["Material Group"].map(group_prices).fillna(0.0)

double_mult = 1.0 + double_loading_pct / 100.0
data["Sided Multiplier"] = np.where(data["Double Sided?"], double_mult, 1.0)

data["Line Value (ex GST)"] = (
    data["Total Area m²"] * data["Price per m²"] * data["Sided Multiplier"]
)

# ---------- Step 4: Pricing view & totals ----------

st.subheader("Step 3 – Calculated pricing")

pricing_cols = [
    "Stock Name",
    "Material Group",
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

# ---------- Step 5: Final preview & download ----------

st.subheader("Step 4 – Final Excel preview")

# Also include friendly group name in export
data["Friendly Group Name"] = data["Material Group"].apply(friendly_group_name)

st.dataframe(data, use_container_width=True)

output_filename = "bp_tender_priced.xlsx"

with pd.ExcelWriter(output_filename, engine="xlsxwriter") as writer:
    data.to_excel(writer, index=False, sheet_name="Priced Tender")
    group_summary.to_excel(writer, index=False, sheet_name="Group Summary")

with open(output_filename, "rb") as f:
    st.download_button(
        "Download Final Priced Excel",
        data=f,
        file_name=output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
