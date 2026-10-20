import openpyxl
from openpyxl import load_workbook

wb = load_workbook("/home/z/my-project/upload/Ассортимент для Олипма.xlsx", data_only=True)
ws = wb["АBС анализ"]

print(f"=== ABC-анализ: {ws.max_row} строк × {ws.max_column} колонок ===\n")

# Print first 20 rows
for r in range(1, 25):
    vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
    if any(v is not None for v in vals):
        print(f"Row {r}: {vals}")
