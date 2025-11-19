# BP Tender SQM Calculator v9 (Streamlit)

This version adds **total value per group** into the **Group Preview** panel and Group Summary sheet.

Features:

- Option B material grouping (thickness + substrate / GSM + finish / SAV brand+code)
- Editable group assignments with dropdown
- Search view for stocks & groups
- Merge-groups tool
- Group Preview with:
  - Material Group
  - Friendly Name
  - **Price per m²**
  - **Group value (ex GST)**
  - Number of materials
  - Number of lines
  - Total area (m²)
- Double-sided override + global loading %
- Per-group price per m²
- Excel export with:
  - `Priced Tender` sheet
  - `Group Summary` sheet (includes price per m² and group value)

## Required Excel Columns

- `Dimensions`
- `Print/Stock Specifications`
- `Total Annual Volume`

## Install

```bash
pip install -r requirements.txt
```

## Run

```bash
streamlit run app.py
```
