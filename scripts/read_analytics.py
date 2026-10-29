"""Прочитать данные аналитики по Красноярскому краю и Хакасии."""
import openpyxl
from openpyxl import load_workbook

for filename, sheet_name in [
    ("ДФО + В Сиб Красноярский край ver2.xlsx", "Красноярский край"),
    ("ДФО + В Сиб Хакасия республика ver2.xlsx", "Хакасия республика"),
]:
    path = f"/home/z/my-project/upload/{filename}"
    wb = load_workbook(path, data_only=True)
    ws = wb[sheet_name]
    
    print(f"\n{'=' * 100}")
    print(f"ФАЙЛ: {filename} — лист '{sheet_name}'")
    print(f"Строк: {ws.max_row}, колонок: {ws.max_column}")
    print(f"{'=' * 100}\n")
    
    # Print all rows
    for r in range(1, min(ws.max_row + 1, 110)):
        vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        if any(v is not None for v in vals):
            print(f"Row {r}: {vals}")
