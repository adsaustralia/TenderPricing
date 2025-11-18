# BP Tender SQM Calculator (Streamlit) – Grouped Stocks & Double-Sided Control

This Streamlit app loads a BP tender Excel file, calculates square metre usage,
lets you group similar materials for pricing, and gives you full control over
double-sided lines.

---

## Features

### 1. Base tender handling

- Reads Excel columns:
  - **Dimensions** – used to calculate m² per item (assumes mm × mm)
  - **Print/Stock Specifications** – used to derive stock name and auto-detect single/double sided
  - **Total Annual Volume** – used as quantity

- Calculates:
  - Area m² per item
  - Total Area m² per line
  - Line value based on group price and double-sided loading

---

### 2. Step 1 – Double-sided control

- Auto-detects "Double Sided" from the specification text.
- Shows an editable table with a **Double Sided?** checkbox per line.
- You can manually tick/untick any line:
  - Tick → double-sided loading applied
  - Untick → no extra loading

---

### 3. Step 2 – Stock grouping for pricing

- Each unique **Stock Name** is listed in a table with an editable **Group Name**.
- Rules:
  - Stocks with the **same Group Name** share one price per m².
  - Stocks with **different Group Names** have different prices.

This allows you to:

- Combine similar materials (e.g. different synthetic variations) into a single group.
- Separate any stock out of a group by giving it a unique Group Name.
- Keep your original Excel unchanged while controlling pricing structure from the UI.

---

### 4. Step 3 – Pricing per group

- In the sidebar, you enter **Price per m²** for each **Stock Group**.
- You also set a **Double-sided loading (%)**.
- For each line, the app calculates:

`Line Value = Total Area m² × Price per m² × Sided Multiplier`

where:

- `Sided Multiplier = 1 + (loading% / 100)` for double-sided lines
- `Sided Multiplier = 1` for single-sided lines

---

### 5. Step 4 – Preview & export

- Shows a full preview of the final dataset.
- Exports to **bp_tender_priced.xlsx** with all calculated values.

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
