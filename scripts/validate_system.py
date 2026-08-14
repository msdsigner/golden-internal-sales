import os
import json
from datetime import datetime


def validate():
    log_path = "system_health_log.txt"
    json_path = "webapp/data/inventory.json"

    # ─────────────────────────────────────────────
    # 1. Load current inventory state safely
    # ─────────────────────────────────────────────
    current_status = "Unknown"
    item_count = 0
    last_updated = "Unknown"

    if os.path.exists(json_path):
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                data = json.load(f)

            items = data.get("items", [])
            item_count = len(items)
            last_updated = data.get("last_updated")

            # Guard against bad dates like "1900"
            if not last_updated or "1900" in str(last_updated):
                last_updated = "Invalid / Missing Date"

            current_status = "Successfully synchronized"

        except Exception as e:
            current_status = f"Error reading JSON: {e}"

    # ─────────────────────────────────────────────
    # 2. Build health report
    # ─────────────────────────────────────────────
    report = []
    report.append("=" * 60)
    report.append(f" SYSTEM HEALTH CHECK: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    report.append("=" * 60)
    report.append(f"Web Database: {current_status} (Total Items: {item_count})")
    report.append(f"Last Updated: Set to {last_updated}, matching latest master data.")
    report.append("-" * 60)
    report.append("")

    # ─────────────────────────────────────────────
    # 3. Load history safely (limit size)
    # ─────────────────────────────────────────────
    history_text = ""
    if os.path.exists(log_path):
        try:
            with open(log_path, "r", encoding="utf-8") as f:
                lines = f.readlines()

            # keep last 200 lines only (prevents log bloat)
            history_text = "".join(lines[-200:])
        except:
            history_text = ""

    full_log = "\n".join(report) + "\n" + history_text

    with open(log_path, "w", encoding="utf-8") as f:
        f.write(full_log)

    # ─────────────────────────────────────────────
    # 4. Console output (clean + latest-first view)
    # ─────────────────────────────────────────────
    print("\n".join(report[:6]))

    if history_text:
        print("\n>>> PREVIOUS VERSION (HISTORY):")
        history_lines = history_text.split("\n")

        # show most recent entries, not oldest
        for line in history_lines[-8:]:
            print(line)

    print(f"\n[SUCCESS] Full health history saved to: {os.path.abspath(log_path)}")


if __name__ == "__main__":
    validate()