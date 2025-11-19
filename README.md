# BP Tender SQM Calculator v11 (Streamlit, Orange UI)

This version adds a **clean orange theme** and keeps all the v10 functionality.

## Visual changes

- Soft orange background for app and sidebar
- Orange pill tag under the title
- Orange rounded buttons with hover state
- Chip-style tips above the double-sided table
- Card-style containers around summary metrics
- Steps clearly labelled 1–4

## Functional features

- Option B material grouping (thickness + substrate / GSM + finish / SAV brand+code)
- Editable group assignments with dropdown
- Search view for stocks & groups (under an expander)
- Merge-groups tool
- Persistent price memory:
  - Per Material Group
  - Per Stock Name
  - Stored locally in `price_memory.json`
- Pricing logic:
  1. If stock-specific price > 0 → use that
  2. Else use material group price
  3. Apply double-sided loading % if flagged
- Group Preview with:
  - Material Group
  - Friendly Name
  - Price per m² (max within group)
  - Group value (ex GST)
  - Number of materials
  - Number of lines
  - Total area (m²)
- Double-sided override + global loading %
- Excel export:
  - `Priced Tender`
  - `Group Summary`

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
