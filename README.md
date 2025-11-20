# ADS Tender SQM Calculator v12.5 (Streamlit, ADS Orange + Navy)

This version adds **per-run metrics** on top of the per-annum SQM logic:

- ADS orange + navy theme
- ADS logo on the top-left of the console (`ads_logo.png`)
- Option B material grouping
- Group + stock price memory (`price_memory.json`)
- Double-sided loading logic (configurable %)
- Per-annum and per-run calculations:
  - Total Area m² per annum
  - Area m² per run
  - Line Value (ex GST) per annum
  - Value per Run (ex GST)
- KPI cards:
  - Total Area (m² per annum)
  - Total Value (ex GST)
  - Average m² per Run
  - Average Value per Run (ex GST)
- Group preview with price and total value
- Final calculated table and Excel export

## Install

```bash
pip install -r requirements.txt
```

## Run

```bash
streamlit run app.py
```
