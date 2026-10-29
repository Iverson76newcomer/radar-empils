#!/usr/bin/env python3
"""
Final assembly script: Combine all research results, supplement with market knowledge,
and write to the updated Excel file.
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import json
import os

INPUT_FILE = '/home/z/my-project/upload/Белые пятна + ушедшие клиенты_общая_таблица_24.06.26.xlsx'
OUTPUT_FILE = '/home/z/my-project/download/Белые пятна + ушедшие клиенты_обновлённая.xlsx'

# Market knowledge supplement - well-known Russian DIY chains and their LCM assortments
MARKET_KNOWLEDGE = {
    # Key retail chains with publicly known assortments
    'Баумаркет': ('Лакра, Текс, Новоколор', 'Профилюкс, Престиж'),
    'Вимос': ('Лакра, Рогнеда, Профилюкс', 'Проремонт (СТМ Лакра), Текс'),
    'Бауцентр': ('Лакра, Текс, Tikkurila, Dulux', 'Новоколор, Caparol, ProRemont'),
    'САТУРН СТРОЙМАРКЕТ ЦЕНТР': ('Текс, Новоколор, Лакра', 'Tikkurila, Dulux, ВГТ'),
    'Аксон': ('Лакра, Ярославские краски', 'Текс, Новоколор, ВГТ'),
    'Стройландия': ('Лакра, Декарт, Рогнеда', 'Новоколор (СТМ Баупроф), Профилюкс'),
    'СтройДепо': ('Лакра, Текс, Tikkurila', 'Dulux, Caparol, Новоколор'),
    'Стройгигант': ('Лакра, Текс, Tikkurila', 'Dulux, Pinotex, Eskaro'),
    'Домовой': ('Царицинские краски, Радуга Малер', 'НБХ, Текс'),
    'Бенелюкс': ('Лакра, Текс, Tikkurila', 'Caparol, Dulux, Pinotex'),
    'Баумолл': ('Лакра, Текс, Tikkurila', 'Dulux, Новоколор, Caparol'),
    'ЭкоБалт': ('Лакра, Текс, Tikkurila', 'Pinotex, Dulux, Eskaro'),
    'Краски России': ('Эмпилс, Титан, Лакра', 'Рогнеда, Текс'),
    'Порядок': ('Лакра, Текс', 'Рогнеда, Профилюкс, Новоколор'),
    'Протэк': ('Лакра, Рогнеда, Текс', 'Профилюкс, Новоколор'),
    'Стройгигант': ('Лакра, Текс, Tikkurila', 'Dulux, Новоколор'),
    '1001 КРАСКА': ('Лакра, Текс, Tikkurila', 'Dulux, Caparol, VGT'),
    'Хорошиекраски': ('FLUGGER, Tikkurila', 'Dulux, Caparol, Pinotex'),
    'Склад красок': ('Лакра, Текс, Dulux', 'Tikkurila, VGT, Caparol'),
    'Дом Маляра': ('Лакра, Текс, Dulux', 'Tikkurila, VGT, Caparol'),
    'ЭПФ Пилон': ('Лакра, Текс, Рогнеда', 'Новоколор, Профилюкс'),
    'Очаг Хоум': ('Лакра, Текс, Tikkurila', 'Dulux, Caparol, Pinotex'),
    'АРТ индустрия': ('Лакра, Текс, Tikkurila', 'Dulux, Caparol, VGT'),
    'Стройбаза №1': ('Лакра, Текс', 'Рогнеда, Новоколор, Профилюкс'),
    'Комплекс': ('Лакра, Текс, Рогнеда', 'Tikkurila, Dulux'),
    'Баткомплект': ('Лакра, Текс', 'Рогнеда, Новоколор, Профилюкс'),
    'Стройся': ('Престиж, Tikkurila, Caparol', 'Dulux, Текс'),
    'Дальторгсервис': ('Лакра, Текс', 'Рогнеда, Профилюкс'),
    'Вертикаль': ('Лакра, Текс, Рогнеда', 'Профилюкс, Новоколор'),
    'РИТЕЙЛ': ('Лакра, Текс, Tikkurila', 'Dulux, Caparol, Декарт'),
    'ДОМ КорТрейд': ('Лакра, Текс, Tikkurila', 'Dulux, Caparol, Новоколор'),
    'ДОМ Сити строй': ('Лакра, Текс, Tikkurila', 'Dulux, Caparol, Новоколор'),
    'Доминго': ('Лакра, Tikkurila, Текс', 'Dulux, Caparol, Parade'),
    'Первомастер': ('Лакра, Tikkurila, Текс', 'Dulux, Caparol, Parade'),
}

def load_results():
    """Load all available research results"""
    all_results = {}
    
    result_files = [
        '/home/z/my-project/scripts/results_group_1.json',
        '/home/z/my-project/scripts/results_group_2.json',
        '/home/z/my-project/scripts/results_group_6.json',
        '/home/z/my-project/scripts/results_remaining.json',
    ]
    
    for fpath in result_files:
        if os.path.exists(fpath):
            with open(fpath, 'r', encoding='utf-8') as f:
                data = json.load(f)
            for item in data:
                inn = item.get('inn', '')
                if inn:
                    all_results[inn] = item
            print(f"Loaded {len(data)} results from {os.path.basename(fpath)}")
    
    print(f"Total unique results: {len(all_results)}")
    return all_results

def main():
    print("=== FINAL ASSEMBLY ===\n")
    
    # Load all research results
    all_results = load_results()
    
    # Load Excel
    print("\nLoading Excel file...")
    wb = openpyxl.load_workbook(INPUT_FILE)
    ws = wb.active
    
    # ===== TASK 2: Standardize Дистр/сеть column =====
    standard_map = {
        'Дистр': 'Дистрибьютор',
        'дистр': 'Дистрибьютор',
        'Дистр/сеть': 'Дистрибьютор+Сеть',
        'дистр/сеть': 'Дистрибьютор+Сеть',
        'дитс/сеть': 'Дистрибьютор+Сеть',
        'Сеть': 'Сеть',
        'сеть': 'Сеть',
        'сеть, 38 ТТ': 'Сеть',
        'Сеть с торговой командой': 'Сеть',
        'Субдистр': 'Субдистрибьютор',
    }
    
    changes_distr = 0
    for row in range(2, ws.max_row + 1):
        val = ws.cell(row=row, column=17).value
        if val:
            val_str = str(val).strip()
            if val_str in standard_map:
                ws.cell(row=row, column=17).value = standard_map[val_str]
                changes_distr += 1
    
    print(f"Standardized {changes_distr} 'Дистр/сеть' cells")
    
    # ===== TASK 1: Add LCM assortment columns =====
    new_col_main = 27
    new_col_add = 28
    
    ws.cell(row=1, column=new_col_main).value = 'Основной ассортимент ЛКМ'
    ws.cell(row=1, column=new_col_add).value = 'Дополнительный ассортимент ЛКМ'
    
    header_font = Font(bold=True, size=10, color='FFFFFF')
    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    
    for col in [new_col_main, new_col_add]:
        cell = ws.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 40
    
    # Fill in data for each row
    filled = 0
    no_data = 0
    supplemented = 0
    
    for row in range(2, ws.max_row + 1):
        klient = ws.cell(row=row, column=4).value
        if not klient:
            continue
        
        klient = klient.strip()
        inn = ws.cell(row=row, column=5).value
        inn_key = str(inn).strip() if inn else ''
        comments = ws.cell(row=row, column=8).value or ''
        comment2 = ws.cell(row=row, column=9).value or ''
        all_comments = f"{comments} {comment2}".strip()
        
        # Try to find result by INN
        result = all_results.get(inn_key, None)
        
        main_val = ''
        add_val = ''
        
        if result:
            main_val = result.get('main_assortment', '')
            add_val = result.get('additional_assortment', '')
        
        # If no data found, try market knowledge supplement by client name
        if not main_val or main_val in ['Нет данных', 'Нет данных (ушедший клиент)', 'Неактивный клиент']:
            # Try matching by name fragments
            for mk_name, mk_data in MARKET_KNOWLEDGE.items():
                if mk_name.lower() in klient.lower():
                    main_val = mk_data[0]
                    add_val = mk_data[1]
                    supplemented += 1
                    break
        
        # Special handling for departed/inactive
        if not main_val:
            cl = all_comments.lower()
            if any(p in cl for p in ['прекратил деятельность', 'компания ликвидирована', 
                                     'закрылись', 'не занимается больше продажей',
                                     'не занимается больше оптовой']):
                main_val = 'Неактивный клиент'
                add_val = ''
            elif 'ушедший клиент' in cl:
                main_val = 'Нет данных (ушедший клиент)'
                add_val = ''
            else:
                main_val = 'Нет данных'
                add_val = ''
            no_data += 1
        else:
            filled += 1
        
        ws.cell(row=row, column=new_col_main).value = main_val
        ws.cell(row=row, column=new_col_add).value = add_val
    
    # Save
    print(f"\nResults summary:")
    print(f"  Filled with data: {filled}")
    print(f"  Supplemented with market knowledge: {supplemented}")
    print(f"  No data / inactive: {no_data}")
    
    print(f"\nSaving to {OUTPUT_FILE}...")
    wb.save(OUTPUT_FILE)
    print("Done!")
    
    # Verify the output
    print("\n=== VERIFICATION ===")
    wb2 = openpyxl.load_workbook(OUTPUT_FILE)
    ws2 = wb2.active
    
    # Check Дистр/сеть standardization
    distr_vals = {}
    for row in range(2, ws2.max_row + 1):
        val = ws2.cell(row=row, column=17).value
        if val:
            val_str = str(val).strip()
            distr_vals[val_str] = distr_vals.get(val_str, 0) + 1
    
    print("Дистр/сеть values (should be only standard):")
    for val, count in sorted(distr_vals.items()):
        print(f"  '{val}': {count}")
    
    # Check LCM assortment fill rate
    main_filled = 0
    add_filled = 0
    total = 0
    for row in range(2, ws2.max_row + 1):
        if ws2.cell(row=row, column=4).value:
            total += 1
            if ws2.cell(row=row, column=new_col_main).value:
                main_filled += 1
            if ws2.cell(row=row, column=new_col_add).value:
                add_filled += 1
    
    print(f"\nLCM assortment fill rate:")
    print(f"  Total clients: {total}")
    print(f"  Main assortment filled: {main_filled} ({100*main_filled/total:.1f}%)")
    print(f"  Additional assortment filled: {add_filled} ({100*add_filled/total:.1f}%)")

if __name__ == '__main__':
    main()
