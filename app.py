import numpy as np
import pandas as pd
import streamlit as st


# ---------- Helpers ----------

def parse_area_m2(dimensions: str):
    """Convert '841mm x 1189mm' → m²"""
    if pd.isna(dimensions):
        return np.nan

    s = str(dimensions).lower().replace(" ", "").replace("mm", "").replace("×", "x")
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
    if pd.isna(spec):
        return ""
    return str(spec).split(",")[0].strip()


def detect_sides(spec: str):
    """Detect single/double sided. Default to Single Sided."""
    if pd.isna(spec) or str(spec).strip() == "":
        return "Single Sided"

    s = str(spec).lower()
    if "double" in s:
        return "Double Sided"
    return "Single Sided"


# ---------- Streamlit App ----------

st.set_page_config(page_title="BP Tender SQM Calculator", layout="wide")
st.title("BP Tender – Square Metre Price Calculator")

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

st.subheader("Step 1 – Review & adjust double-sided lines")

display_cols = [
    col for col in [
        "Lot ID" if "Lot ID" in data.columns else None,
        "Item Description" if "Item Description" in data.columns else None,
        "Dimensions",
        "Print/Stock Specifications",
        "Quantity",
        "Area m² (each)",
        "Total Area m²",
        "Sided (auto)",
        "Double Sided?",
    ] if col is not None
]

edited = st.data_editor(
    data[display_cols],
    use_container_width=True,
    num_rows="dynamic",
)

# Update the main frame with edited double-sided flags
data["Double Sided?"] = edited["Double Sided?"].astype(bool)

st.markdown(
    "You can tick/untick **Double Sided?** above to enable/disable double-sided pricing for any line."
)

# ---------- Stock grouping & pricing ----------

st.sidebar.header("Pricing & Stock Grouping")

# Unique stocks
unique_stocks = sorted(x for x in data["Stock Name"].dropna().unique() if str(x).strip())

# Initialise grouping in session_state
if "stock_groups" not in st.session_state:
    st.session_state["stock_groups"] = {s: s for s in unique_stocks}
else:
    # Ensure all current stocks are present
    for s in unique_stocks:
        st.session_state["stock_groups"].setdefault(s, s)

st.sidebar.subheader("Stock → Group mapping")

stock_groups = {}
for stock in unique_stocks:
    default_group = st.session_state["stock_groups"].get(stock, stock)
    group_name = st.sidebar.text_input(
        f"Group for '{stock}'",
        value=default_group,
        key=f"group_{stock}",
    )
    stock_groups[stock] = group_name
    st.session_state["stock_groups"][stock] = group_name

data["Stock Group"] = data["Stock Name"].map(stock_groups).fillna("Unassigned")

# Groups = all distinct group names
group_names = sorted(set(stock_groups.values()))

st.sidebar.subheader("Price per m² (by group)")

group_prices = {}
for group in group_names:
    group_prices[group] = st.sidebar.number_input(
        f"{group}",
        min_value=0.0,
        value=0.0,
        step=0.1,
        key=f"price_group_{group}",
    )

# Double-sided loading
double_loading_pct = st.sidebar.number_input(
    "Double-sided loading (%)",
    min_value=0.0,
    value=25.0,
    step=1.0,
)

data["Price per m²"] = data["Stock Group"].map(group_prices).fillna(0.0)

double_mult = 1.0 + double_loading_pct / 100.0
data["Sided Multiplier"] = np.where(data["Double Sided?"], double_mult, 1.0)

data["Line Value (ex GST)"] = (
    data["Total Area m²"] * data["Price per m²"] * data["Sided Multiplier"]
)

st.subheader("Step 2 – Calculated pricing")

pricing_cols = [
    col for col in [
        "Lot ID" if "Lot ID" in data.columns else None,
        "Item Description" if "Item Description" in data.columns else None,
        "Stock Name",
        "Stock Group",
        "Dimensions",
        "Quantity",
        "Total Area m²",
        "Double Sided?",
        "Price per m²",
        "Sided Multiplier",
        "Line Value (ex GST)",
    ] if col is not None
]

st.dataframe(data[pricing_cols], use_container_width=True)

total_area = data["Total Area m²"].sum(skipna=True)
total_value = data["Line Value (ex GST)"].sum(skipna=True)

col1, col2 = st.columns(2)
col1.metric("Total Area (m²)", f"{total_area:,.2f}")
col2.metric("Total Value (ex GST)", f"${total_value:,.2f}")

st.subheader("Step 3 – Final Excel preview")

st.dataframe(data, use_container_width=True)

# ---- Excel Download ----
output_filename = "bp_tender_priced.xlsx"

with pd.ExcelWriter(output_filename, engine="xlsxwriter") as writer:
    data.to_excel(writer, index=False, sheet_name="Priced Tender")

with open(output_filename, "rb") as f:
    st.download_button(
        "Download Final Priced Excel",
        data=f,
        file_name=output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
