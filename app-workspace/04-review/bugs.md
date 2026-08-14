# Bug Log

## Active Issues

| ID | Description | Reported | Status | Priority |
|----|-------------|----------|--------|----------|
| BUG-001 | "Last Updated" date on website is stuck (e.g., April 22, 2026) | 2026-04-28 | Open | High |
| BUG-002 | Specific items missing from website (e.g., 250BX390K, 250BX496K, etc.) | 2026-04-28 | Open | High |
| BUG-003 | Local inventory.json updates but website doesn't reflect changes | 2026-04-28 | Open | High |

## Missing Items List
- 250BX390K
- 250BX496K
- 250F400W
- 250HPBT546R
- FBHP660K
- HPBT195K
- FWPB515K
- FWPB515W
- 250BT344R
- 250BX210K
- FBX643K
- FBX647K
- 303DGNX
- MAW08AV1QWT-C

## Investigation Notes
- Local `webapp/data/inventory.json` shows "April 29, 2026" but user sees "April 22".
- This suggests a deployment failure or a caching issue on the live site.
- Missing items were not found in local `inventory.json` grep search.
