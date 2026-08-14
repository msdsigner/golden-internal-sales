# Skill: Golden Opportunity Master Protocol

## 1. Inventory Taxonomy & Categorization
This protocol governs product classification for both Excel reports and the Web Dashboard.

### Mandatory Logic Rules
- **Heating & Cooking:** Merge "Heating" and "Cooking" into a single "Heating & Cooking" sub-category.
- **Major Appliances:** 
    - Parent Category: **"Major Appliances"**.
    - Sub-Category (Catch-all): **"Other Appliances"** (includes Dishwashers, Water Dispensers).
- **Personal Care:** Group all "Hair Dryer", "Clipper", and "Trimmer" items here.
- **Beverage Centers:** Merge all "Wine Cellars" here.

---

## 2. Data Integrity & Sync (Inventory Audit)
Ensures perfect parity between the Odoo/Excel source and the Web JSON database.

### The Audit Workflow
1. **Sync Process:** Run `scripts/sync_inventory.py` to compare Excel Sheet dates with Web JSON.
2. **Quantity/Price Check:** Ensure SKU-level data matches the master workbook exactly.
3. **Timestamping:** Website "Last Updated" date must match the **Excel source sheet date**.
4. **Omission Scanning:** Flag items missing from either the Excel source or the Web catalog.

---

## 3. Quality Control: Output & Formatting
Guarantees high-fidelity exports and stable web performance.

### Excel & PDF Formatting
- **Image Scaling:** Maintain aspect ratios to fit `135x135` bounds within `140px` cells.
- **Centering:** Use dynamic EMU padding (`1px = 9525 EMUs`) for precise centering.
- **Anchoring:** Use `oneCellAnchor` with `TwoCellAnchor` constructor so images move with cells during sorting.
- **PDF Export:** Set `PageSetup.Zoom = False` and `FitToPagesWide = 1` to prevent column clipping or distortion.

### Documentation & Artifacts
- **Sync Status:** Maintain perfect synchronization between code logic and project documentation (`README.md`, skills).
- **Task Verification:** Mark sub-conditions as completed only after concrete verification in the local environment.

---

## 4. Execution & Readiness Checklist
Triggered automatically via `Golden Opportunity Catalog Update.bat`:
1. **Unified Sequence:**
   - **Report Generation:** `golden_inventory_app.py` (Master Excel + Weekly PDF).
   - **Web Sync:** `scripts/sync_inventory.py` (Quick sync of core metrics).
   - **Parity Audit:** `scripts/ensure_parity.py` (Deep verification + Auto-Fix).
   - **Health Check:** `scripts/validate_system.py` (Final integrity audit + logging).
   - **Auto-Deploy:** Push changes to GitHub Pages for live viewing.
