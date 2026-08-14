"""
extract_images.py
Extracts all product images from '2026 Golden Inventory.xlsx'
and saves them to the 'images/' folder, named by item code (e.g. CK-CB26GR.png).
The banner image (row 0) is saved as 'banner.png'.
"""
import openpyxl
import os
import io
from pathlib import Path

SOURCE_FILE = "input/2026 Golden Inventory.xlsx"
IMAGES_DIR = "images"
SHEET_NAME = "Item Listing and Pricing"


def extract_images():
    os.makedirs(IMAGES_DIR, exist_ok=True)

    print(f"Loading {SOURCE_FILE} ...")
    wb = openpyxl.load_workbook(SOURCE_FILE)
    ws = wb[SHEET_NAME]

    # Build 0-indexed row -> item code map from column B
    row_to_item = {}
    for row in ws.iter_rows(min_row=1, min_col=2, max_col=2):
        cell = row[0]
        if cell.value and str(cell.value).strip():
            row_to_item[cell.row - 1] = str(cell.value).strip()

    images = ws._images
    print(f"Found {len(images)} images in sheet.")

    extracted = 0
    skipped = 0
    banner_saved = False

    # Track used codes to avoid duplicates if possible, or handle them
    for img in sorted(images, key=lambda i: i.anchor._from.row):
        anchor_row = img.anchor._from.row
        
        # 0 is banner
        if anchor_row < 3: 
            if not banner_saved:
                file_name = "banner.png"
                banner_saved = True
            else:
                skipped += 1
                continue
        else:
            # Robust matching: find nearest item row
            nearest_dist = 999
            nearest_code = None
            for r_idx, code in row_to_item.items():
                dist = abs(r_idx - anchor_row)
                if dist < nearest_dist:
                    nearest_dist = dist
                    nearest_code = code
            
            if nearest_code and nearest_dist <= 1: # 1 row tolerance
                safe_code = nearest_code.replace("/", "-").replace("\\", "-")
                file_name = f"{safe_code}.png"
            else:
                skipped += 1
                continue

        out_path = os.path.join(IMAGES_DIR, file_name)

        # Extract raw image bytes
        try:
            img_data = img._data()
            with open(out_path, "wb") as f:
                f.write(img_data)
            extracted += 1
            if extracted % 100 == 0:
                print(f"  Extracted {extracted} images so far...")
        except Exception as e:
            print(f"  Warning: could not extract image at row {anchor_row}: {e}")
            skipped += 1

    print(f"\nDone! Extracted: {extracted}, Skipped: {skipped}")
    if banner_saved:
        print(f"Banner saved as: {os.path.join(IMAGES_DIR, 'banner.png')}")
    return extracted


if __name__ == "__main__":
    extract_images()
