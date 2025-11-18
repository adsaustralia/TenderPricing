# BP Tender SQM Calculator (Streamlit) – Option B + Tools

This version adds all of your requested tools on top of **Option B** grouping:

- **Group Preview Panel** – see groups, counts, and total m²
- **Search bar** for quickly finding materials
- **Auto-generated human-readable group names**
- **Merge groups button** – select groups and merge into one

---

## Key Features

### 1. Base tender logic

- Reads Excel columns:
  - `Dimensions` – used to calculate m² (assumes mm × mm)
  - `Print/Stock Specifications` – used to derive Stock Name & sides
  - `Total Annual Volume` – used as quantity

- Calculates:
  - Area m² per item
  - Total area m² per line
  - Line value based on:
    - Material Group price per m²
    - Double-sided loading (if enabled)

---

### 2. Double-sided control (Step 1)

- Auto-detects "Double Sided" from the specification text.
- Shows a table with a **Double Sided?** checkbox per line.
- You can override detection manually.
- Sidebar control: **Double-sided loading (%)**.

---

### 3. Option B material grouping with search & dropdown (Step 2)

- Each unique **Stock Name** is assigned an **Initial Group** using Option B rules:
  - Thickness + substrate for boards
  - GSM + finish for papers
  - Brand + code family for vinyls/SAV
  - Special handling for Jellyfish, Duratran, Yuppo, Braille, glass decor, etc.

- A table shows:
  - `Stock Name` (read-only)
  - `Initial Group` (read-only)
  - `Assigned Group` (editable **dropdown**)

- A **search bar** filters the table by:
  - Stock Name
  - Initial Group
  - Assigned Group

You can:

- Keep the default grouping (do nothing).
- Move a material into a different group by changing its Assigned Group.
- Isolate a material by assigning it to a unique group name.

---

### 4. Merge groups tool

Under **"Merge groups"**:

- Select two or more existing groups from a multiselect.
- Type the **Merged group name** you want.
- Click **"Merge selected groups into target"**.

The app will:

- Update all materials whose `Assigned Group` is in the selected list.
- Set their `Assigned Group` to the target name.
- That new group will then appear in the pricing section.

This makes it easy to consolidate similar groups after initial auto-grouping.

---

### 5. Group Preview Panel (Step 2.5)

A **Group Preview** table is shown with:

- `Material Group` – the internal group key
- `Friendly Name` – a nicer label for display/exports
- `Materials` – number of unique Stock Names in the group
- `Lines` – number of tender lines in that group
- `Total_Area_m2` – total area for that group

This gives you a quick overview of:

- How many items are in each group
- Which groups are big vs small
- Total m² by group (for sanity checks / weighting)

---

### 6. Pricing per material group (Step 3)

In the sidebar:

- For each **Material Group**, you set a **Price per m²**.
- You also set a **Double-sided loading (%)**.

The app computes:

`Line Value = Total Area m² × Price per m² × Sided Multiplier`

Where:

- `Sided Multiplier = 1 + loading% / 100` for double-sided lines.
- `Sided Multiplier = 1` for single-sided lines.

You see a **Calculated pricing** table including:

- Group, dimensions, quantity, area, price per m², multiplier, line value.

---

### 7. Final preview & Excel export (Step 4)

The final table includes:

- All original columns
- `Stock Name`
- `Material Group`
- `Friendly Group Name`
- `Area m² (each)`
- `Total Area m²`
- `Price per m²`
- `Sided Multiplier`
- `Line Value (ex GST)`

Export:

- `Priced Tender` sheet – line-level data
- `Group Summary` sheet – the Group Preview panel (by group)

---

## Requirements

```bash
pip install -r requirements.txt
```

---

## Run

```bash
streamlit run app.py
```

---

## Excel Requirements

Your Excel must contain:

- `Dimensions`
- `Print/Stock Specifications`
- `Total Annual Volume`

Other columns (e.g. `Lot ID`, `Item Description`) are optional and will be preserved.
