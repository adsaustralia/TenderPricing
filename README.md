# Tender Pricing App (Steps 1–3) – Excel-style Preview

This Streamlit app lets you:

1. Upload an Excel file and view it in an **Excel-style grid**:
   - Column headers: `A, B, C, ...`
   - Row numbers: `1, 2, 3, ...`
   - Row 1 contains your original header names.
2. Hide/unhide rows and columns (without deleting them) and export a workbook with those rows/columns marked as hidden.
3. Map fields like Size, Material, Qty per annum, Qty per run, and DS/SS information to calculate:
   - SQM per unit / per annum / per run
   - Price per unit / per annum / per run (based on Material $/sqm and DS loading %).

All row references and column references in the UI are **Excel-style** so they match what you see in your original workbook.

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
