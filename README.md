# ADS Tender SQM Calculator v12 (Streamlit, ADS Orange + Navy)

This version applies **ADS Australia branding** and keeps all functionality:

- Dark navy header bar with orange stripe (via Streamlit header styling)
- ADS logo on the **top-left** of the main console
- ADS orange (`#FF5E19`) used for pill label, buttons, and chips
- Soft warm background for the working area

Functional features:

- Option B material grouping (thickness + substrate / GSM + finish / SAV brand+code)
- Editable group assignments with dropdown
- Search view for stocks & groups (within an expander)
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

The ADS logo is included as `ads_logo.png` and is displayed at the top-left of the console.
