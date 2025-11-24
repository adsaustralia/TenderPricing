# Tender Pricing App (Excel-style, SQM + Pricing, Currency Conversion)

This Streamlit app lets you:

1. Upload an Excel file and view it in an **Excel-style grid**:
   - Column headers: `A, B, C, ...`
   - Row numbers: `1, 2, 3, ...`
   - Row 1 contains your original header names.
2. Hide/unhide rows and columns (without deleting them) and export a workbook with those rows/columns marked as hidden.
3. Map fields like Size, Material, Qty per annum, Qty per run, Runs per annum, and DS/SS information to calculate:
   - SQM per unit / per annum / per run
   - Price per unit / per annum / per run
4. Handle **Qty per run** either as:
   - An explicit "Qty per run" column/row, OR
   - `Qty per annum ÷ Runs per annum`.
5. Enter **Material Price per SQM (AUD)** and apply:
   - Double-sided loading % for DS items
   - Optional **currency conversion** to another display currency (e.g. USD, EUR).

All prices are rounded to **2 decimal places**. In the UI, price columns are displayed with a leading `$` and you can configure the display currency code (default AUD) and a conversion rate (e.g. `1 AUD = 0.65 USD`). The exported Excel file contains the **numeric values** (rounded to 2 decimals) for further use.

## How to run locally

1. Create and activate a virtual environment (optional but recommended).
2. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

3. Run the app:

   ```bash
   streamlit run app.py
   ```

4. Open the URL shown in your terminal (usually http://localhost:8501) in your browser.

## Deployment (e.g., Streamlit Cloud)

1. Create a new GitHub repository and upload:
   - `app.py`
   - `requirements.txt`
   - `README.md`

2. On Streamlit Cloud:
   - Point it to your GitHub repo and `app.py` as the entry file.
   - Deploy and use the web UI to upload your Excel files.
