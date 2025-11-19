# BP Tender SQM Calculator v10 (Streamlit)

This version adds **persistent price memory**:

- Remembers **price per Material Group** across runs.
- Remembers **price per Stock Name** across runs.
- Uses a local JSON file: `price_memory.json` in the app folder.

## Price logic

For each line:

1. If a **stock-specific price** is set (> 0), it is used.
2. Otherwise, the **material group price** is used.
3. Double-sided lines apply the global double-sided loading %.

## Features

- Option B material grouping (thickness + substrate / GSM + finish / SAV brand+code)
- Editable group assignments with dropdown
- Search view for stocks & groups
- Merge-groups tool
- Group Preview with:
  - Material Group
  - Friendly Name
  - Price per m² (max within group)
  - Group value (ex GST)
  - Number of materials
  - Number of lines
  - Total area (m²)
- Double-sided override + global loading %
- Per-group price per m² + per-stock overrides
- Excel export with:
  - `Priced Tender` sheet
  - `Group Summary` sheet

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

A file called `price_memory.json` will be created/updated in the same folder, storing your group and stock prices for future runs.
