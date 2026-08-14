# Stage 02 Output — Design

## Architecture
- **Engine:** `golden_inventory_app.py` (Main orchestrator)
- **Data Pipeline:**
  1. **Input:** `input/Inventory Tool*.xlsx` (Source data)
  2. **Processing:** Python scripts (`scripts/excel_utils.py`, `scripts/export_service.py`)
  3. **Output:** 
     - `output/` (Reports and master workbook)
     - `webapp/data/inventory.json` (Web app database)
- **Deployment:** `scripts/deploy_to_github.py` (Push to GitHub Pages)

## Data Model
- **JSON Schema (`inventory.json`):**
  - `last_updated`: Date string
  - `downloads`: Links to PDF/Excel snapshots
  - `items`: Array of product objects:
    - `id`: Internal Reference
    - `name`: Product Name
    - `category`: Web category (derived)
    - `parent_category`: Top-level grouping
    - `sub_category`: Detailed grouping
    - `price`: Sales Price
    - `available`: Quantity On Hand
    - `image`: Path to product image

## UI Decisions
- **Frontend:** Vanilla JS with Grid/Flexbox layout.
- **Features:** 
  - Dynamic filtering by category/subcategory.
  - Real-time fuzzy search.
  - "Selection Cart" saved in LocalStorage.
  - Export selection to HTML table for Gmail, PDF, or Excel.
- **Aesthetics:** High-contrast "Golden" theme (Dark blue, gold highlights, white cards).
