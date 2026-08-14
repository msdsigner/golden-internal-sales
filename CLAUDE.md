# App Workspace — System Prompt (Layer 0)
> Always load this file first. It is the DNS of this workspace.

## 🧠 Methodology: Interpretable Context (ICM)
This project uses a staged pipeline to ensure AI consistency and human oversight:
1. **Load Context:** Tell the AI: "Load CLAUDE.md, then app-workspace/[stage]/CONTEXT.md."
2. **Review Output:** The AI performs the task and updates `OUTPUT.md` in that stage folder.
3. **Approve & Move:** You review the output before moving to the next stage.

## 📂 Project Structure

### AI Methodology (app-workspace/)
| Folder / File | Purpose | Key Document |
| :--- | :--- | :--- |
| `00-system/` | Meta | `CONTEXT.md` (Role & Identity) |
| `01-define/` | Define | `OUTPUT.md` (Goals & Scope) |
| `02-design/` | Design | `OUTPUT.md` (Architecture & Data Model) |
| `03-build/` | Build | `ACTIVE-TASK.md` (Current Work) |
| `04-review/` | Review | `bugs.md` (Issue Tracking) |
| `05-ship/` | Ship | `OUTPUT.md` (Launch Checklist) |
| `_memory/` | Memory | `decisions.md` (Architectural Log) |
| `_skills/` | Reference | Prompt patterns and reusable skills |

### Implementation
- **Core Orchestrator:** `golden_inventory_app.py`
- **Web App:** `webapp/` (Frontend assets)
- **Scripts:** `scripts/` (Data processing & Deployment)
- **Data:** `input/` (Source Excel) & `output/` (Reports)
- **Assets:** `images/` (Product photography)
- **Logs:** `system_health_log.txt` (Sync & Deployment status)

## 📜 Rules
1. **Single Source of Truth:** Never skip a stage; always update the relevant `OUTPUT.md` before moving on.
2. **Persistence:** Always append significant choices to `app-workspace/_memory/decisions.md`.
3. **Data Integrity:** Updates must respect MD5 hashing and incremental skipping logic.
4. **Premium Aesthetics:** Maintain the "Golden" design system (Dark blue, Gold accents, Glassmorphism, 4-column grid).
5. **Efficiency:** Keep methodology files under 800 tokens; split if they become too long.
6. **Task Focus:** Only one active task at a time in `app-workspace/03-build/ACTIVE-TASK.md`.
7. **Local/Web Parity (Crucial Skill):** The local site and the live web app MUST ALWAYS be in sync. Any time you make edits to files in the `webapp/` directory (like `index.html`, `style.css`, or `script.js`), you MUST proactively run `python scripts/deploy_to_github.py` to push those changes to the live site. Never let the local code diverge from the published site.
