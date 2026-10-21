import openpyxl
from openpyxl import load_workbook

# === 1. Check third tab "300 SKU с ценами" in assortment ===
wb = load_workbook("/home/z/my-project/upload/Ассортимент_Олимп_300_SKU.xlsx", data_only=True)
ws = wb["300 SKU с ценами"]

print(f"=== Вкладка '300 SKU с ценами' (rows={ws.max_row}, cols={ws.max_column}) ===")
# Print header (row 4) and first 5 data rows
print("\n--- Row 1 (title) ---")
print([c.value for c in ws[1]])
print("\n--- Row 2 ---")
print([c.value for c in ws[2]])
print("\n--- Row 4 (headers) ---")
print([c.value for c in ws[4]])
print("\n--- First 3 data rows ---")
for r in range(5, 8):
    print(f"Row {r}:", [c.value for c in ws[r]])

# Check 8 previously unmatched items
print("\n=== 8 ранее не сопоставленных SKU на вкладке '300 SKU с ценами' ===")
unmatched_idx = [37, 38, 56, 61, 62, 63, 150, 169]
for r in range(5, ws.max_row + 1):
    idx_val = ws.cell(row=r, column=1).value
    if idx_val in unmatched_idx:
        print(f"\nRow {r} (idx={idx_val}):")
        for col_idx, cell in enumerate(ws[r], 1):
            header = ws.cell(row=4, column=col_idx).value
            print(f"  {header}: {cell.value}")
