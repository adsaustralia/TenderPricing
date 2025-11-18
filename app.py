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


# ---------- Streamlit App ----------

st.set_page_config(page_title="BP Tender SQM Calculator", layout="wide")
st.title("BP Tender – Square Metre Price Calculator (Grouped Stocks)")

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

# Update main data with edited Double Sided? values
data["Double Sided?"] = double_sided_view["Double Sided?"].fillna(False)

st.caption("You can enable/disable double-sided pricing per line using the checkboxes above.")

# ---------- Step 2: Stock grouping ----------

st.subheader("Step 2 – Group similar materials for pricing")

unique_stocks = sorted(
    s for s in data["Stock Name"].dropna().unique() if str(s).strip()
)

if "stock_groups_df" not in st.session_state:
    st.session_state["stock_groups_df"] = pd.DataFrame({
        "Stock Name": unique_stocks,
        "Group Name": unique_stocks,  # default: each stock is its own group
    })
else:
    sg = st.session_state["stock_groups_df"]
    existing = set(sg["Stock Name"])
    new_rows = [s for s in unique_stocks if s not in existing]
    if new_rows:
        sg = pd.concat(
            [sg, pd.DataFrame({"Stock Name": new_rows, "Group Name": new_rows})],
            ignore_index=True,
        )
    sg = sg[sg["Stock Name"].isin(unique_stocks)].reset_index(drop=True)
    st.session_state["stock_groups_df"] = sg

st.markdown(
    """Type the **Group Name** for each stock below:

- Stocks with the **same Group Name** share one price per m²  
- Stocks with **different Group Names** have different prices  

Example: put multiple synthetics into **\"Synthetic Group\"**.
"""
)

stock_groups_df = st.data_editor(
    st.session_state["stock_groups_df"],
    use_container_width=True,
    num_rows="fixed",
    key="stock_group_editor",
)

st.session_state["stock_groups_df"] = stock_groups_df

stock_to_group = dict(
    zip(stock_groups_df["Stock Name"], stock_groups_df["Group Name"])
)

data["Stock Group"] = data["Stock Name"].map(stock_to_group).fillna("Unassigned")

# ---------- Step 3: Pricing per group & calculation ----------

st.sidebar.header("Pricing & double-sided loading")

group_names = sorted(
    g for g in data["Stock Group"].dropna().unique() if str(g).strip()
)

group_prices = {}
st.sidebar.subheader("Price per m² by group")

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

data["Price per m²"] = data["Stock Group"].map(group_prices).fillna(0.0)

double_mult = 1.0 + double_loading_pct / 100.0
data["Sided Multiplier"] = np.where(data["Double Sided?"], double_mult, 1.0)

data["Line Value (ex GST)"] = (
    data["Total Area m²"] * data["Price per m²"] * data["Sided Multiplier"]
)

st.subheader("Step 3 – Calculated pricing")

pricing_cols = [
    "Stock Name",
    "Stock Group",
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

# ---------- Step 4: Final preview & download ----------

st.subheader("Step 4 – Final Excel preview")

st.dataframe(data, use_container_width=True)

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
