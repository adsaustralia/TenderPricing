# BP Tender SQM Calculator (Streamlit) – Grouped Stocks

This Streamlit app loads a BP tender Excel file, calculates square metre usage,
groups similar materials for pricing, lets you break stocks out of groups, and
allows you to enable/disable double-sided pricing per line.

---

## Key Features

- Reads Excel columns:
  - **Dimensions** – used to calculate m² per item (assumes mm × mm)
  - **Print/Stock Specifications** – used to derive stock name and auto-detect single/double sided
  - **Total Annual Volume** – used as quantity

- **Step 1 – Double-sided control**
  - Auto-detects "Double Sided" based on text
  - Shows an editable table with a **Double Sided?** checkbox per line
  - You can manually tick/untick any line to override the automatic detection

- **Step 2 – Grouped stock pricing**
  - Each unique stock name can be assigned to a **group** in the sidebar
  - All stocks in the same group share a **single price per m²**
  - To:
    - **Combine similar materials** → give them the same group name  
    - **Separate a stock from a group** → give it a unique group name  
  - You then enter **price per m² per group** (not per individual stock)

- **Double-sided loading**
  - Configurable **Double-sided loading (%)** in the sidebar
  - Double-sided lines are multiplied by: `1 + loading% / 100`

- **Preview and export**
  - Shows:
    1. A table to adjust double-sided flags
    2. Calculated pricing table
    3. Full final dataset
  - Exports `bp_tender_priced.xlsx` with all calculated values.

---

## Requirements

```bash
pip install -r requirements.txt
```

---

## Run

```bash
streamlit run app.py
```

---

## Excel Format Requirements

The Excel must contain at least these columns:

- `Dimensions`
- `Print/Stock Specifications`
- `Total Annual Volume`

Other columns (e.g. `Lot ID`, `Item Description`) are optional. If present, they
are preserved and shown in the app and output.
