
import streamlit as st
import openpyxl
import pandas as pd
import re
from io import BytesIO
from openpyxl.utils import column_index_from_string, get_column_letter

st.set_page_config(page_title="Excel SQM & Pricing Tool", layout="wide")
st.title("📊 Excel SQM & Pricing Calculator — Multi-Sheet Version")

st.write("""
Upload your Excel file, define column/row settings, and input pricing rules.  
The app will calculate **SQM and prices** for each sheet, display previews, and let you **download the updated Excel file or a combined summary**.
""")


def normalize(s):
    return re.sub(r'[^a-z0-9]+', '', str(s).lower()) if s else ""


SIDE_MULTIPLIERS = {
    "2mm Screenboard": {"Single sided": 1.0, "Double sided": 1.2},
    "3mm Screenboard": {"Single sided": 1.0, "Double sided": 1.2},
    "3mm Corflute": {"Single sided": 1.0, "Double sided": 1.2},
    "280gsm Synthetic (Plasnet)": {"Single sided": 1.0, "Double sided": 1.3},
    "300gsm Silk": {"Single sided": 1.0, "Double sided": 1.2},
    "250gsm Silk": {"Single sided": 1.0, "Double sided": 1.2},
    "200gsm Gloss": {"Single sided": 1.0, "Double sided": 1.2},
    "200gsm Matt": {"Single sided": 1.0, "Double sided": 1.2},
    "200gsm Satin": {"Single sided": 1.0, "Double sided": 1.2},
    "150gsm Silk": {"Single sided": 1.0, "Double sided": 1.2},
    "Jellyfish Supercling (Synthetic)": {"Single sided": 1.0, "Double sided": 1.3},
    "Jellyfish Supercling on Ferrous Substrate": {"Single sided": 1.0, "Double sided": 1.3},
    "Jellyfish": {"Single sided": 1.0, "Double sided": 1.3},
    "0.6mm Magnetic": {"Single sided": 1.0, "Double sided": 1.3},
    "Duratran Backlit": {"Single sided": 1.0, "Double sided": 1.0},
    "Yupo Synthetic Paper": {"Single sided": 1.0, "Double sided": 1.2},
    "MPI 2126 Hi-Tack SAV": {"Single sided": 1.0, "Double sided": 1.2},
    "MPI 2904 Easy Apply SAV": {"Single sided": 1.0, "Double sided": 1.2},
    "MPI 2903 SAV": {"Single sided": 1.0, "Double sided": 1.2},
    "SAV 3302": {"Single sided": 1.0, "Double sided": 1.2},
    "3M 7725 SAV": {"Single sided": 1.0, "Double sided": 1.2},
    "Metamark M7 SAV": {"Single sided": 1.0, "Double sided": 1.2},
    "Arlon 8000 SAV": {"Single sided": 1.0, "Double sided": 1.2},
    "Arlon 6700 SAV": {"Single sided": 1.0, "Double sided": 1.2},
    "Ultra Tac SAV": {"Single sided": 1.0, "Double sided": 1.3},
    "Polymeric SAV + ACM": {"Single sided": 1.0, "Double sided": 1.3},
    "Generic SAV": {"Single sided": 1.0, "Double sided": 1.2},
    "Mactac Glass Decor Dusted": {"Single sided": 1.0, "Double sided": 1.0},
    "Frosted/Dusted Glass SAV": {"Single sided": 1.0, "Double sided": 1.0},
    "Ultra Clear SAV": {"Single sided": 1.0, "Double sided": 1.0},
    "Cool Grey SAV": {"Single sided": 1.0, "Double sided": 1.0},
    "Matt Black CCV 900 Series": {"Single sided": 1.0, "Double sided": 1.0},
    "Metamark Gloss Black M7": {"Single sided": 1.0, "Double sided": 1.0},
    "Braille Acrylic Signs": {"Single sided": 1.0, "Double sided": 1.0},
    "3mm Satin Black Acrylic": {"Single sided": 1.0, "Double sided": 1.0},
    "Maxi T + PVC Assembly": {"Single sided": 1.0, "Double sided": 1.0},
}


def detect_category(material: str):
    if not material:
        return None

    s = str(material)
    m = normalize(s)

    if "screenboard" in m and "2mm" in m:
        return "2mm Screenboard"
    if "screenboard" in m and "3mm" in m:
        return "3mm Screenboard"

    if ("corflute" in m or "coreflute" in m) and "3mm" in m:
        return "3mm Corflute"

    if "expandedpvc" in m and "1mm" in m:
        return "1mm Expanded PVC"
    if "1mmpvc" in m and "expanded" not in m:
        return "1mm PVC"

    if "anodised" in m and "aluminium" in m:
        return "1mm Anodised Aluminium"

    if "satinblackacrylic" in m or ("satin" in m and "black" in m and "acrylic" in m):
        return "3mm Satin Black Acrylic"

    if "maxit" in m and "pvc" in m:
        return "Maxi T + PVC Assembly"

    if "brailleacrylicsigns" in m:
        return "Braille Acrylic Signs"

    if "plasnet" in m or ("synthetic" in m and "280gsm" in m):
        return "280gsm Synthetic (Plasnet)"

    if "titansilk" in m or ("300gsm" in m and "silk" in m):
        return "300gsm Silk"
    if "200gsm" in m and "silk" in m:
        return "200gsm Silk"
    if "150gsm" in m and "silk" in m:
        return "150gsm Silk"

    if "250gsm" in m and ("ecomatt" in m or "satin" in m or "silk" in m):
        return "250gsm Silk"

    if "200gsm" in m and "gloss" in m:
        return "200gsm Gloss"
    if "200gsm" in m and "matt" in m:
        return "200gsm Matt"
    if "200gsm" in m and "satin" in m:
        return "200gsm Satin"

    if "duratran" in m or ("backlit" in m and "duratran" in m):
        return "Duratran Backlit"

    if "yuppo" in m or "yuposyntheticpaper" in m:
        return "Yupo Synthetic Paper"

    if "jellyfish" in m and "supercling" in m:
        if "ferrous" in m:
            return "Jellyfish Supercling on Ferrous Substrate"
        return "Jellyfish Supercling (Synthetic)"

    if "jellyfish" in m:
        return "Jellyfish"

    if "magnetic" in m:
        return "0.6mm Magnetic"

    if "2126" in m and "sav" in m:
        return "MPI 2126 Hi-Tack SAV"
    if "2904" in m and "sav" in m:
        return "MPI 2904 Easy Apply SAV"
    if "2903" in m and "sav" in m:
        return "MPI 2903 SAV"
    if "3302" in m and "sav" in m:
        return "SAV 3302"
    if "7725" in m and "3m" in m:
        return "3M 7725 SAV"
    if "metamark" in m and "m7" in m:
        return "Metamark M7 SAV"
    if "arlon8000" in m or ("arlon" in m and "8000" in m):
        return "Arlon 8000 SAV"
    if "arlon6700" in m or ("arlon" in m and "6700" in m):
        return "Arlon 6700 SAV"
    if "ultratac" in m:
        return "Ultra Tac SAV"
    if "polymericsav" in m and "acm" in m:
        return "Polymeric SAV + ACM"
    if "sav" in m:
        return "Generic SAV"

    if "mactacglassdecor" in m or ("glassdecor" in m and "mactac" in m):
        return "Mactac Glass Decor Dusted"
    if "frosteddustedglasssav" in m or ("frosted" in m and "glass" in m):
        return "Frosted/Dusted Glass SAV"
    if "ultraclearsav" in m or ("ultra" in m and "clear" in m and "sav" in m):
        return "Ultra Clear SAV"

    if "coolgrey" in m and "sav" in m:
        return "Cool Grey SAV"
    if "mattblackccv900series" in m or ("matt" in m and "black" in m and "ccv900" in m):
        return "Matt Black CCV 900 Series"
    if "metamarkglossblackm7" in m or ("metamark" in m and "black" in m and "m7" in m):
        return "Metamark Gloss Black M7"

    return s.strip()


def parse_size(raw):
    if not raw:
        return None, None
    s = str(raw).replace("×", "x").replace("X", "x").replace("*", "x")
    nums = re.findall(r'\d+(?:\.\d+)?', s)
    if len(nums) >= 2:
        return float(nums[0]), float(nums[1])
    return None, None


def parse_qty(raw):
    if raw is None:
        return None
    if isinstance(raw, (int, float)):
        return float(raw)
    m = re.search(r'\d+(?:\.\d+)?', str(raw).replace(",", ""))
    return float(m.group(0)) if m else None


def clean_value(v):
    if v is None:
        return None
    if isinstance(v, str) and v.strip().startswith("="):
        return None
    return v


uploaded_file = st.file_uploader("📤 Upload Excel file", type=["xlsx"])

if uploaded_file:

    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    sheet_names = wb.sheetnames

    st.subheader("⚙️ Excel Structure Settings")
    col1, col2 = st.columns(2)
    with col1:
        start_col = st.text_input("Start Column", value="AC")
        end_col = st.text_input("End Column", value="IG")
    with col2:
        row_size = st.number_input("Row (Size)", value=5)
        row_material = st.number_input("Row (Material)", value=6)
        row_qty = st.number_input("Row (Quantity)", value=155)
        row_sqm = st.number_input("Row (SQM Output)", value=156)
        row_price = st.number_input("Row (Price Output)", value=157)

    st.subheader("🧮 Print Side")
    side_option = st.radio(
        "Select print side:",
        options=["Single sided", "Double sided"],
        index=0,
        horizontal=True,
    )

    st.subheader("📄 Sheet Selection")
    process_all = st.checkbox("Process all sheets", value=True)
    sheet_choice = None
    if not process_all:
        sheet_choice = st.selectbox("Select sheet", sheet_names)

    start_idx = column_index_from_string(start_col)
    end_idx = column_index_from_string(end_col)

    st.subheader("🧾 Materials & Category Overrides")

    detected_materials = set()
    source_sheets = sheet_names if process_all else [sheet_choice]
    for sname in source_sheets:
        ws = wb[sname]
        for c in range(start_idx, end_idx + 1):
            col = get_column_letter(c)
            raw_mat = clean_value(ws[f"{col}{row_material}"].value)
            if raw_mat:
                detected_materials.add(str(raw_mat).strip())

    if not detected_materials:
        st.error("❌ No materials found. Check row/column settings.")
        st.stop()

    material_categories = {}
    for mat in sorted(detected_materials):
        default_cat = detect_category(mat) or mat
        c1, c2 = st.columns([2, 2])
        with c1:
            st.write(mat)
        with c2:
            override = st.text_input(
                "Category",
                value=default_cat,
                key=f"cat_{normalize(mat)}",
                help="Edit this if the auto group is wrong. Materials with the same category share one rate.",
            )
        cat_final = override.strip() or default_cat
        material_categories[mat] = cat_final

    categories_present = sorted(set(material_categories.values()))

    st.markdown("---")
    st.subheader("💰 Category Pricing & Enable/Disable")

    base_rates = {}
    category_enabled = {}

    for cat in categories_present:
        col_a, col_b = st.columns([1, 2])
        with col_a:
            enabled = st.checkbox(
                f"Use '{cat}'",
                value=True,
                key=f"enable_{normalize(cat)}",
            )
        with col_b:
            base = st.number_input(
                f"Base rate (Single sided, AUD/m²) for '{cat}'",
                min_value=0.0,
                value=0.0,
                step=0.1,
                key=f"base_{normalize(cat)}",
            )
        category_enabled[cat] = enabled
        base_rates[cat] = base

    def get_effective_rate(material: str, side: str):
        if material is None:
            return None
        cat = material_categories.get(material) or detect_category(material) or material
        if not category_enabled.get(cat, True):
            return None
        base = base_rates.get(cat, 0.0)
        if base <= 0:
            return None
        mult = SIDE_MULTIPLIERS.get(cat, {}).get(side, 1.0)
        return base * mult

    if st.button("🚀 Process & Calculate"):
        summary = []

        def process_sheet(ws, sheetname):
            rows = []
            total_cost = 0

            for c in range(start_idx, end_idx + 1):
                col = get_column_letter(c)

                raw_size = clean_value(ws[f"{col}{row_size}"].value)
                raw_mat = clean_value(ws[f"{col}{row_material}"].value)
                raw_qty = clean_value(ws[f"{col}{row_qty}"].value)

                w, h = parse_size(raw_size)
                qty = parse_qty(raw_qty)
                rate = get_effective_rate(raw_mat, side_option)
                sqm = price = None

                if w and h and qty:
                    sqm = (w / 1000) * (h / 1000) * qty
                    if rate is not None:
                        price = round(sqm * rate, 2)
                        total_cost += price
                    ws[f"{col}{row_sqm}"].value = sqm
                    ws[f"{col}{row_price}"].value = price

                rows.append({
                    "Column": col,
                    "Material": raw_mat,
                    "Category": material_categories.get(raw_mat) or detect_category(raw_mat) or raw_mat,
                    "Size": raw_size,
                    "Qty": qty,
                    "Rate": rate,
                    "SQM": sqm,
                    "Price": price,
                })

            ws[f"{end_col}{row_sqm}"] = "TOTAL"
            ws[f"{end_col}{row_price}"] = total_cost
            return pd.DataFrame(rows), total_cost

        if process_all:
            for sname in sheet_names:
                df, total = process_sheet(wb[sname], sname)
                st.markdown(f"### 📄 {sname}")
                st.dataframe(df)
                st.success(f"Subtotal: ${total:,.2f}")
                summary.append({"Sheet": sname, "Total": total})
        else:
            df, total = process_sheet(wb[sheet_choice], sheet_choice)
            st.markdown(f"### 📄 {sheet_choice}")
            st.dataframe(df)
            st.success(f"Total: ${total:,.2f}")
            summary.append({"Sheet": sheet_choice, "Total": total})

        st.subheader("📘 Combined Totals")
        summary_df = pd.DataFrame(summary)
        st.dataframe(summary_df)
        st.success(f"GRAND TOTAL: ${summary_df['Total'].sum():,.2f}")

        excel_bytes = BytesIO()
        wb.save(excel_bytes)
        excel_bytes.seek(0)

        st.download_button(
            "⬇️ Download Updated Excel",
            excel_bytes,
            "Updated_Pricing.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
