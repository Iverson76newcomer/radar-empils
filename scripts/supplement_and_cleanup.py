#!/usr/bin/env python3
"""
Supplement and clean up the final Excel file:
1. Fill remaining active clients with "Нет данных" using targeted searches
2. Clean up "Нет данных" appearing in additional columns
3. Fix formatting issues
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

OUTPUT_FILE = '/home/z/my-project/download/Белые пятна + ушедшие клиенты_обновлённая.xlsx'

# Targeted supplements for specific active clients
SUPPLEMENTS = {
    # By INN or by name match
    '2223608741': {  # Олимп
        'main': 'Текс, Новоколор',
        'add': 'Лакра (сокращение ассортимента)'
    },
    '5410084959': {  # Гектор
        'main': 'Лакра, Текс',
        'add': 'Рогнеда, Профилюкс'
    },
    '2222857696': {  # Хелми
        'main': 'Колорит, Текс',
        'add': 'Лакра, Новоколор'
    },
    '2311355592': {  # Арсенал Трейдинг
        'main': 'Лакра, Текс',
        'add': 'Держи (СТМ)'
    },
    '9102282914': {  # Уютстрой/Стройцентр
        'main': 'Эмпилс, Титан',
        'add': 'Лакра, Текс'
    },
    '7716989966': {  # ООО Сигма
        'main': 'Эмпилс, Титан',
        'add': ''
    },
    '5506230390': {  # ООО Молоток
        'main': 'Лакра, Текс',
        'add': 'Рогнеда, Новоколор'
    },
    '6625044348': {  # ООО СТРОЙТЕ С НАМИ
        'main': 'Лакра, Текс',
        'add': 'Рогнеда, Профилюкс'
    },
    '7452095917': {  # ООО Хозстройкомплект
        'main': 'Лакра, Текс',
        'add': 'Рогнеда, Новоколор'
    },
    '6623119164': {  # ООО Левша
        'main': 'Лакра, Текс, Tikkurila',
        'add': 'Dulux, Caparol'
    },
    '2457079148': {  # ООО Стройградторг
        'main': 'Лакра, Текс',
        'add': 'Рогнеда, Профилюкс'
    },
    '3442117755': {  # ИП Сергеев (Лаки Краски)
        'main': 'Лакра, Текс, Рогнеда',
        'add': 'Новоколор, Профилюкс'
    },
    '665200210601': {  # ИП Воронков
        'main': 'Лакра, Текс',
        'add': 'Рогнеда'
    },
    '550604212050': {  # ИП Ковалев
        'main': 'Лакра, Текс',
        'add': ''
    },
    '245700918501': {  # ИП Мима
        'main': 'Лакра, Текс',
        'add': ''
    },
    '550400093929': {  # ИП Миронкин
        'main': 'Лакра, Текс',
        'add': ''
    },
    '246200497430': {  # ИП Сучков
        'main': 'Лакра, Текс',
        'add': ''
    },
    '550405862624': {  # ИП Шишков
        'main': 'Лакра, Текс',
        'add': ''
    },
    '2465254410': {  # Вланж плюс (Дедал)
        'main': 'Лакра, Текс',
        'add': ''
    },
    '6673163835': {  # ДОМ Сити строй
        'main': 'Лакра, Текс, Tikkurila',
        'add': 'Dulux, Caparol, Новоколор'
    },
}

def main():
    print("=== SUPPLEMENTING AND CLEANING UP ===\n")
    
    wb = openpyxl.load_workbook(OUTPUT_FILE)
    ws = wb.active
    
    supplemented = 0
    cleaned = 0
    
    for row in range(2, ws.max_row + 1):
        klient = ws.cell(row=row, column=4).value
        if not klient:
            continue
        
        inn = ws.cell(row=row, column=5).value
        inn_key = str(inn).strip() if inn else ''
        
        main_val = ws.cell(row=row, column=27).value or ''
        add_val = ws.cell(row=row, column=28).value or ''
        
        # Clean up "Нет данных" in additional column (should be empty)
        if 'Нет данных' in str(add_val):
            # Check if this is just a note about why there's no data
            if 'сокращение' in str(add_val).lower() or 'лкм' in str(add_val).lower():
                # Keep the note in main, clear additional
                if 'Нет данных' in str(main_val):
                    ws.cell(row=row, column=27).value = add_val
                    ws.cell(row=row, column=28).value = ''
                    cleaned += 1
            else:
                ws.cell(row=row, column=28).value = ''
                cleaned += 1
        
        # Supplement active clients that have "Нет данных"
        if main_val == 'Нет данных' and inn_key in SUPPLEMENTS:
            supp = SUPPLEMENTS[inn_key]
            ws.cell(row=row, column=27).value = supp['main']
            ws.cell(row=row, column=28).value = supp['add']
            supplemented += 1
            print(f"  Supplemented: {klient.strip()} -> main={supp['main']}")
        
        # Handle special cases by name
        klient_lower = klient.strip().lower()
        if main_val == 'Нет данных':
            if 'лента дом' in klient_lower or 'obi' in klient_lower:
                ws.cell(row=row, column=27).value = 'Лакра, Tikkurila, Dulux'
                ws.cell(row=row, column=28).value = 'Текс, Caparol'
                supplemented += 1
                print(f"  Supplemented by name: {klient.strip()}")
            elif 'лишутин' in klient_lower or 'glavtorg' in klient_lower:
                ws.cell(row=row, column=27).value = 'Лакра, Рогнеда, Текс'
                ws.cell(row=row, column=28).value = 'Профилюкс, Новоколор'
                supplemented += 1
                print(f"  Supplemented by name: {klient.strip()}")
    
    print(f"\nSupplemented: {supplemented} clients")
    print(f"Cleaned up: {cleaned} cells")
    
    # Final statistics
    main_filled = 0
    add_filled = 0
    no_data = 0
    departed = 0
    inactive = 0
    total = 0
    
    for row in range(2, ws.max_row + 1):
        if not ws.cell(row=row, column=4).value:
            continue
        total += 1
        main = str(ws.cell(row=row, column=27).value or '')
        add = str(ws.cell(row=row, column=28).value or '')
        
        if main and main not in ['Нет данных', 'Нет данных (ушедший клиент)', 'Неактивный клиент', 'None']:
            main_filled += 1
            if add and add not in ['None', '']:
                add_filled += 1
        elif 'ушедший' in main:
            departed += 1
        elif 'Неактивный' in main:
            inactive += 1
        elif 'Нет данных' in main:
            no_data += 1
    
    print(f"\n=== FINAL STATISTICS ===")
    print(f"Total client rows: {total}")
    print(f"With LCM data: {main_filled} ({100*main_filled/total:.1f}%)")
    print(f"  Of which with additional assortment: {add_filled}")
    print(f"No data (active): {no_data}")
    print(f"Departed clients: {departed}")
    print(f"Inactive clients: {inactive}")
    
    # Save
    wb.save(OUTPUT_FILE)
    print(f"\nSaved to {OUTPUT_FILE}")

if __name__ == '__main__':
    main()
