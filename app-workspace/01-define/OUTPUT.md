# Stage 01 Output — Define

## The problem
- **Who:** Golden Auto sales team and customers.
- **Workaround today:** Sending manual Excel files or PDFs to customers; manual data entry for inventory updates.
- **Pain:** Inventory data becomes stale quickly; customers can't easily search or filter available products; risk of over-selling.

## The app
- **One-liner:** A real-time web-based inventory dashboard synced automatically from local Excel inventory tools.
- **Core action:** Processing "Inventory Tool" Excel files to update a live web catalog.
- **Success looks like:** Customers can browse, search, and select items on a premium web interface that reflects the latest stock levels and prices.

## Scope — v1 only

**Must have:**
1. Automated processing of "Inventory Tool" Excel files.
2. Generation of high-quality PDF and Excel reports for offline use.
3. Live web dashboard with search, filtering, and "Add to Selection" features.

**Not in v1:**
1. Direct Odoo API integration (currently uses Excel exports).
2. Full e-commerce checkout (uses "Copy to Gmail" selection instead).
3. Multi-user login/permissions for the dashboard.

**Biggest risk or unknown:** Parity between the source Excel data and the web-app JSON database.

## Constraints
- **Stack:** Python (OpenPyXL, jsons) for backend; HTML/JS/CSS for frontend; GitHub Pages for deployment.
- **Timeline:** Continuous improvement (v1 already deployed).
- **Non-negotiables:** Premium visual design; seamless Excel-to-Web sync.
