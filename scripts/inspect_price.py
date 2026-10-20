import openpyxl
from openpyxl import load_workbook

wb2 = load_workbook("/home/z/my-project/upload/21.08.2026 РОССИЯ с НДС, руб..xlsx", data_only=True)

# Inspect AvantGarde sheet — print all rows until headers and several data rows
for sn in ['AvantGarde', 'ОРЕОЛ TITAN', 'Ореол', 'Расцвет']:
    ws = wb2[sn]
    print(f"\n{'=' * 100}")
    print(f"Sheet: '{sn}' (max_row={ws.max_row})")
    print('=' * 100)
    for row_idx, row in enumerate(ws.iter_rows(min_row=8, max_row=20, values_only=True), 8):
        # show all 14 cols
        print(f"Row {row_idx}: {row}")

# Try to find header row pattern
print("\n\n=== Detecting header row in 'AvantGarde' ===")
ws = wb2['AvantGarde']
for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=20, values_only=True), 1):
    non_empty = [c for c in row if c is not None]
    if non_empty:
        print(f"Row {row_idx}: non_empty count={len(non_empty)}, first vals: {row[:14]}")
