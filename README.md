
# Excel SQM & Pricing Calculator (Streamlit) — v2

This version includes:
- Automatic grouping of similar materials into categories
- Per-material category override in the UI
- Enable/disable checkbox per category
- Per-client base rate entry (single sided)
- Single/Double sided pricing via multipliers

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the URL shown in the terminal (usually http://localhost:8501).

Deploy to Streamlit Cloud by pushing these files to GitHub and selecting `app.py` as the entrypoint.
