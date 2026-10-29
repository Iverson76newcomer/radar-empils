#!/usr/bin/env python3
"""
Task 2: Standardize the 'Дистр/сеть' column in the Excel table.
Maps 9 variations to 4 standard values.
Also adds two new column headers for LCM assortment data.
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import shutil

INPUT_FILE = '/home/z/my-project/upload/Белые пятна + ушедшие клиенты_общая_таблица_24.06.26.xlsx'
OUTPUT_FILE = '/home/z/my-project/download/Белые пятна + ушедшие клиенты_обновлённая.xlsx'

def main():
    print("Loading Excel file...")
    wb = openpyxl.load_workbook(INPUT_FILE)
    ws = wb.active
    
    # ===== STANDARDIZATION MAP =====
    # All 9 variations mapped to 4 standard values
    standard_map = {
        'Дистр': 'Дистрибьютор',
        'дистр': 'Дистрибьютор',
        'Дистр/сеть': 'Дистрибьютор+Сеть',
        'дистр/сеть': 'Дистрибьютор+Сеть',
        'дитс/сеть': 'Дистрибьютор+Сеть',  # typo
        'Сеть': 'Сеть',
        'сеть': 'Сеть',
        'сеть, 38 ТТ': 'Сеть',
        'Сеть с торговой командой': 'Сеть',
        'Субдистр': 'Субдистрибьютор',
    }
    
    # Count current values
    current_vals = {}
    for row in range(2, ws.max_row + 1):
        val = ws.cell(row=row, column=17).value
        if val:
            val_str = str(val).strip()
            current_vals[val_str] = current_vals.get(val_str, 0) + 1
    
    print("Current values in 'Дистр/сеть' column:")
    for val, count in sorted(current_vals.items()):
        new_val = standard_map.get(val_str := val, val)
        print(f"  '{val}' ({count}x) -> '{standard_map.get(val, val)}'")
    
    # Apply standardization
    changes = 0
    for row in range(2, ws.max_row + 1):
        val = ws.cell(row=row, column=17).value
        if val:
            val_str = str(val).strip()
            if val_str in standard_map:
                ws.cell(row=row, column=17).value = standard_map[val_str]
                changes += 1
    
    print(f"\nStandardized {changes} cells")
    
    # ===== ADD NEW COLUMN HEADERS for LCM assortment =====
    new_col_main = 27  # Основной ассортимент ЛКМ
    new_col_add = 28   # Дополнительный ассортимент ЛКМ
    
    ws.cell(row=1, column=new_col_main).value = 'Основной ассортимент ЛКМ'
    ws.cell(row=1, column=new_col_add).value = 'Дополнительный ассортимент ЛКМ'
    
    # Style the headers
    header_font = Font(bold=True, size=10, color='FFFFFF')
    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    
    for col in [new_col_main, new_col_add]:
        cell = ws.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 40
    
    # ===== ALSO standardize the Дистр/сеть values in the duplicate mapping =====
    # Verify final values
    final_vals = {}
    for row in range(2, ws.max_row + 1):
        val = ws.cell(row=row, column=17).value
        if val:
            val_str = str(val).strip()
            final_vals[val_str] = final_vals.get(val_str, 0) + 1
    
    print("\nFinal values in 'Дистр/сеть' column:")
    for val, count in sorted(final_vals.items()):
        print(f"  '{val}' ({count}x)")
    
    # Save
    print(f"\nSaving to {OUTPUT_FILE}...")
    wb.save(OUTPUT_FILE)
    print("Done! Task 2 completed successfully.")

if __name__ == '__main__':
    main()
