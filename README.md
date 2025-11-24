# Tender Pricing App (Steps 1–3)

This Streamlit app lets you:

1. Upload an Excel file and view all rows/columns exactly as they appear in Excel (only the header row, no extra mapping tables).
2. Hide/unhide rows and columns (without deleting them) and export a workbook with those rows/columns marked as hidden.
3. Map fields like Size, Material, Qty per annum, Qty per run, and DS/SS information to calculate:
   - SQM per unit / per annum / per run
   - Price per unit / per annum / per run (based on Material $/sqm and DS loading %).

Column selections in the UI use Excel-style labels like `A - Lot ID`, `B - Country` etc., so you can think in Excel terms while the data still displays like your original sheet.

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
