import os
import sys
# Ensure we can import from the root directory
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import openpyxl
import json
from datetime import datetime

from scripts.excel_utils import get_latest_inventory_sheet, normalize_sku

def ensure_parity():
    """
    Cross-references the Web Database (JSON) with the Master Output (Excel).
    Detects discrepancies in Price or Quantity, auto-fixes the JSON,
    and logs all actions to parity_audit.txt.
    """
    year = datetime.now().year
    excel_path = os.path.join("output", f"{year} Golden Inventory.xlsx")
    json_path = os.path.join("webapp", "data", "inventory.json")
    error_log_path = 'parity_errors.txt'
    
    print("\n" + "="*60)
    print("  PARITY AUDIT & AUTO-FIX ENGINE")
    print("="*60)

    if not os.path.exists(excel_path):
        print(f"[Error] Excel file not found: {excel_path}")
        return
    if not os.path.exists(json_path):
        print(f"[Error] JSON file not found: {json_path}")
        return

    # 1. Load Excel Master Data (Latest Sheet)
    wb = openpyxl.load_workbook(excel_path, data_only=True)

    sheet_name, _ = get_latest_inventory_sheet(wb)
    sheet = wb[sheet_name]
    
    excel_data = {}
    # Build header map (same logic as sync_inventory.py)

    rows = list(sheet.iter_rows(values_only=True))

    headers = [str(h).strip().lower() if h else "" for h in rows[0]]

    col = {
        "item": None,
        "price": None,
        "qty": None
    }

    for i, h in enumerate(headers):
        if "item" in h or "internal" in h or "sku" in h:
            col["item"] = i
        elif "price" in h:
            col["price"] = i
        elif "qty" in h or "quantity" in h or "available" in h or "on hand" in h:
            col["qty"] = i

    if col["item"] is None:
        raise ValueError("Item/SKU column not found in parity audit.")

    for row in rows[1:]:
        sku = str(row[col["item"]] or "").strip()
        if not sku:
            continue

        price = row[col["price"]] if col["price"] is not None else 0
        qty = row[col["qty"]] if col["qty"] is not None else 0

        try:
            qty = int(float(qty))
        except:
            qty = 0

        try:
            price = f"{float(price):.2f}"
        except:
            price = "0.00"

        normalized_excel_sku = normalize_sku(sku)
        excel_data[normalized_excel_sku] = {
            "raw_sku": sku,
            "price": price,
            "available": qty
        }

    # 2. Load Web Database (JSON)
    with open(json_path, 'r', encoding='utf-8') as f:
        web_db = json.load(f)
    
    items = web_db.get('items', [])
    mismatches = []
    omissions = []
    fixes_applied = 0

    # 3. Perform Parity Check
    json_skus_normalized = set()
    publish_pricing = web_db.get("publish_pricing", True)
    
    for item in items:
        sku = str(item.get('id', '')).strip()
        normalized_web_sku = normalize_sku(sku)
        json_skus_normalized.add(normalized_web_sku)
        
        if normalized_web_sku in excel_data:
            if publish_pricing:
                ex_price = excel_data[normalized_web_sku]['price']
                ex_qty = excel_data[normalized_web_sku]['available']
            else:
                ex_price = "Hidden"
                ex_qty = "Hidden"
            
            web_price = str(item.get('price', '0.00'))
            # Format web price to match float string
            if publish_pricing:
                try:
                    web_price = f"{float(web_price):.2f}"
                except: pass
            
            web_qty = item.get('available', 0) if publish_pricing else "Hidden"
            
            if web_price != ex_price or web_qty != ex_qty:
                mismatches.append(f"Mismatch [{sku}]: Web(Q:{web_qty}, P:${web_price}) vs Excel(Q:{ex_qty}, P:${ex_price})")
                # AUTO-FIX
                item['price'] = ex_price
                item['available'] = ex_qty
                fixes_applied += 1
        else:
            # Item in JSON but NOT in Excel (Omission or Stale)
            omissions.append(f"Stale Item [{sku}]: Found in Web but missing from Excel Master.")

    # 4. Check for Omissions (Items in Excel but missing from JSON)
    for normalized_sku_key in excel_data:
        if normalized_sku_key not in json_skus_normalized:
            raw_sku = excel_data[normalized_sku_key]["raw_sku"]
            omissions.append(f"Missing Item [{raw_sku}]: Found in Excel but missing from Web Database.")

    # 5. Save Results & Record Errors (Append mode)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = [
        f"--- AUDIT LOG: {timestamp} ---",
        f"Excel Source: {sheet_name}",
        f"Items Checked: {len(items)}",
        f"Mismatches Fixed: {fixes_applied}",
        f"Total Omissions/Stale: {len(omissions)}",
        "Details:"
    ]
    log_entry.extend(mismatches)
    log_entry.extend(omissions)
    log_entry.append("-" * 40 + "\n")

    # Record errors/mismatches to a persistent text file by appending
    with open(error_log_path, 'a', encoding='utf-8') as f:
        f.write("\n".join(log_entry))

    # Save fixed JSON if changes were made
    if fixes_applied > 0:
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(web_db, f, indent=4)
        print(f"[Success] Applied {fixes_applied} auto-fixes to the Web Database.")
    else:
        print("[Status] Parity check passed. No mismatches found.")

    if omissions:
        print(f"[Warning] Found {len(omissions)} omissions/stale items. Check {error_log_path} for details.")

    print(f"[Log] Audit details recorded in {error_log_path}")

if __name__ == "__main__":
    ensure_parity()
