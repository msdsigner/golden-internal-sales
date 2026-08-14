# Active Task: Advanced Automation Control & UI Enhancements

## Goal
Implement command-line controls for the automation and enhance the web dashboard with internal health checks and selection review.

## Tasks
### Phase 1: Automation Controls (Task 5)
- [ ] Add `--force` flag to `golden_inventory_app.py` to bypass incremental skipping.
- [ ] Add `--dry-run` flag to preview changes without modifying files or deploying.
- [ ] Update the batch script to support these new options.

### Phase 2: Data Health & Selection Review (Tasks 4 & 3)
- [ ] Create a "Health Check" data endpoint/view to identify catalog inconsistencies (missing images, 0 price).
- [ ] Implement an expanded "Review Selection" view in the web app for final cart verification.
- [ ] Ensure all UI additions follow the established "Golden" aesthetic.

## Context
- **Affected Files:** `golden_inventory_app.py`, `webapp/index.html`, `webapp/script.js`, `webapp/style.css`.
- **Logic:** Use `argparse` for flags; extend `script.js` for new dashboard views.
