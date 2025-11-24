# Tender Pricing App (Excel-style, Grouped Pricing, Default Preset, JSON Save Options)

This Streamlit app lets you:

1. Upload an Excel file and view it in an **Excel-style grid**:
   - Column headers: `A, B, C, ...`
   - Row numbers: `1, 2, 3, ...`
   - Row 1 contains your original header names.
2. Hide/unhide rows and columns (without deleting them) and export a workbook with those rows/columns marked as hidden.
3. Map fields like Size, Material, Qty per annum, Qty per run, Runs per annum, and DS/SS information to calculate:
   - SQM per unit / per annum / per run
4. Handle **Qty per run** either as:
   - An explicit "Qty per run" column/row, OR
   - `Qty per annum ÷ Runs per annum`.
5. Use an interactive **Material Groups & Pricing Presets** UI to:
   - Assign each material to a **Group name** (e.g. "3mm ACM", "Posters", "Window Vinyl").
   - Enter a **Group Price per SQM (AUD)** which applies to all materials in that group.
   - Optionally set **per-material override prices**.
6. Automatically load a **default preset** (`material_groups_default.json`) from the repo on startup.
7. After you update groups/prices in the UI, choose how to save them:
   - **Only for this session** (no JSON file is modified).
   - **Update default JSON on server** (`material_groups_default.json` is overwritten, if the environment allows writing).
   - Always have the option to **download** the latest preset as `material_groups_preset.json` and commit it to GitHub.
8. Enter a **display currency** and conversion rate:
   - Base prices are in **AUD**.
   - The app can also calculate and display converted prices (e.g. `1 AUD = 0.65 USD`).

All SQM and price values are rounded to **2 decimal places**.  
Price columns in the UI are formatted with a leading `$` for readability, and the downloaded Excel file contains the underlying numeric values.

## JSON saving behaviour

- On startup, the app tries to load `material_groups_default.json` to pre-fill:
  - `group_assignments` (material → group)
  - `group_prices` (group → AUD price per SQM)
  - `material_overrides` (material → AUD price per SQM override)
- After editing material/group prices, you can choose:
  - **Only for this session** – keep everything in memory; the JSON file is not changed.
  - **Update default JSON on server** – the app attempts to overwrite `material_groups_default.json` in the current environment.
    - On local runs, this will normally succeed and persist.
    - On Streamlit Cloud or similar, the filesystem may be ephemeral; use the download option and commit to Git instead.
- You can always click **Download current material group preset (JSON)** to get a file you can store in version control.

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
   - `material_groups_default.json` (optional but recommended)

2. On Streamlit Cloud:
   - Point it to your GitHub repo and `app.py` as the entry file.
   - Deploy and use the web UI to upload your Excel files.
