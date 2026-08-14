import openpyxl

wb = openpyxl.load_workbook('input/2026 Golden Inventory.xlsx')
ws = wb['Item Listing and Pricing']

# Build row -> item code map from column B (item codes start at row 4)
row_to_item = {}
for row in ws.iter_rows(min_row=4, min_col=2, max_col=2, values_only=False):
    cell = row[0]
    if cell.value and str(cell.value).strip():
        row_to_item[cell.row - 1] = str(cell.value).strip()  # anchor is 0-indexed

print('Row-to-item map (first 20):')
for r, item in sorted(row_to_item.items())[:20]:
    print(f'  row_0idx={r}: {item}')

# Sort images by row
images = ws._images
print(f'\nImages by anchor row (first 20):')
for img in sorted(images, key=lambda i: i.anchor._from.row)[:20]:
    fr = img.anchor._from
    matched_item = row_to_item.get(fr.row, '???')
    print(f'  anchor_row={fr.row}, col={fr.col}, item={matched_item}')
