# BP Tender SQM Calculator v7 (Streamlit)

This app reads a tender Excel file and prices items by **square metre** with:

- Option B material grouping (thickness + substrate / GSM + finish / SAV brand+code)
- Editable group assignments with dropdown
- Search view for stocks & groups
- Group merge tool
- Group preview panel (materials, lines, total m²)
- Double-sided override column + global loading %
- Per-group price per m²
- Excel export with:
  - `Priced Tender` sheet
  - `Group Summary` sheet

## Required Excel Columns

Your Excel must contain at least:

- `Dimensions` (e.g. `841mm x 1189mm`, assumed mm)
- `Print/Stock Specifications`
- `Total Annual Volume`

Other columns (e.g. `Lot ID`, `Item Description`) are optional and will be passed through.

## Install

```bash
pip install -r requirements.txt
```

## Run

```bash
streamlit run app.py
```
