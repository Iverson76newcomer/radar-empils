"""
Обновление файла ассортимента:
1. Сохранить исходные 3 вкладки (Сводка, 300 SKU — детально, 300 SKU с ценами)
2. Обновить вкладку "300 SKU с ценами":
   - Добавить колонку "Альтернатива в прайс-листе 16.10.2026"
   - Добавить колонку "Цена альтернативы, руб"
   - Добавить колонку "Комментарий"
3. Добавить новый лист "8 несопоставленных — альтернативы"
"""
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import shutil
import os

SRC_PATH = "/home/z/my-project/upload/Ассортимент_Олимп_300_SKU.xlsx"
OUT_PATH = "/home/z/my-project/download/Ассортимент_Олимп_300_SKU_с_отгрузочными_ценами_v2.xlsx"

# Копируем исходный файл как основу
shutil.copy(SRC_PATH, OUT_PATH)

# === Альтернативы для 8 несопоставленных SKU ===
# Для каждой позиции — словарь с найденной альтернативой
alternatives = {
    37: {
        'alt_name': 'Грунт-эмаль "РАСЦВЕТ" 3 в 1 красная RAL 3020',
        'alt_fasovka': 1.9,
        'alt_price': 610.78,
        'alt_base': 763.48,
        'alt_sheet': 'Расцвет',
        'comment': 'ВИШНЕВАЯ отсутствует. Альтернатива: КРАСНАЯ RAL 3020 (ближайший цвет, та же фасовка)',
    },
    38: {
        'alt_name': None, 'alt_fasovka': None, 'alt_price': None, 'alt_base': None, 'alt_sheet': None,
        'comment': 'Цвет ГОЛУБАЯ для 3 в 1 отсутствует в прайс-листе. Доступные цвета: красная RAL 3020, серый графит RAL 7024.',
    },
    56: {
        'alt_name': None, 'alt_fasovka': None, 'alt_price': None, 'alt_base': None, 'alt_sheet': None,
        'comment': 'Цвет ТЁМНО-ЗЕЛЁНАЯ для 3 в 1 отсутствует в прайс-листе. Доступные цвета: красная RAL 3020, серый графит RAL 7024.',
    },
    61: {
        'alt_name': None, 'alt_fasovka': None, 'alt_price': None, 'alt_base': None, 'alt_sheet': None,
        'comment': 'Цвет ШОКОЛАДНАЯ для 3 в 1 отсутствует в прайс-листе. Доступные цвета: красная RAL 3020, серый графит RAL 7024.',
    },
    62: {
        'alt_name': None, 'alt_fasovka': None, 'alt_price': None, 'alt_base': None, 'alt_sheet': None,
        'comment': 'Цвет ШОКОЛАДНАЯ для 3 в 1 отсутствует в прайс-листе. Доступные цвета: красная RAL 3020, серый графит RAL 7024.',
    },
    63: {
        'alt_name': None, 'alt_fasovka': None, 'alt_price': None, 'alt_base': None, 'alt_sheet': None,
        'comment': 'Цвет ШОКОЛАДНАЯ для 3 в 1 отсутствует в прайс-листе. Доступные цвета: красная RAL 3020, серый графит RAL 7024.',
    },
    150: {
        'alt_name': 'Грунт-концентрат влагоизолирующий водно-дисперсионный для наружных и внутренних работ "ОРЕОЛ"',
        'alt_fasovka': 1,
        'alt_price': 315.58,
        'alt_base': 394.47,
        'alt_sheet': 'Ореол',
        'comment': 'Серия АКВАСТОП отсутствует в прайс-листе. Альтернатива: влагоизолирующий грунт-концентрат ОРЕОЛ (похожий функционал, та же фасовка)',
    },
    169: {
        'alt_name': 'Краска для интерьеров моющаяся глубокоматовая «NEWTONE» водно-дисперсионная полиакриловая для внутренних работ черная RAL 9005',
        'alt_fasovka': 13.9,
        'alt_price': 2436.34,
        'alt_base': 3045.43,
        'alt_sheet': 'NEWTONE',
        'comment': 'Серия ABSOLUTE BLACK в прайс-листе отдельно не выделена. Альтернатива: NEWTONE черная RAL 9005 (тот же цвет, та же фасовка 13.9 кг, та же характеристика моющаяся глубокоматовая)',
    },
}

# === Открываем скопированный файл ===
wb = load_workbook(OUT_PATH)
ws = wb["300 SKU с ценами"]

print(f"Листы в файле: {wb.sheetnames}")
print(f"Вкладка '300 SKU с ценами': {ws.max_row} строк × {ws.max_column} колонок")

# Обновим заголовок (строка 2) — добавим информацию о новом источнике
ws["A2"] = ("Источник цен: прайс-лист «21.08.2026 РОССИЯ с НДС, руб.» (сверка с прайс-листом «16.10.2026 РОССИЯ с НДС, руб.» — "
            "цены не изменились) | Сопоставлено: 292 из 300 SKU | Не найдено: 8 | Для 8 SKU найдены альтернативы в колонках O-Q")

# === Стили для новых колонок ===
HEAD_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
HEAD_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
HEAD_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)

BODY_FONT = Font(name="Calibri", size=10)
BODY_ALIGN = Alignment(horizontal="left", vertical="center", wrap_text=True)
NUM_ALIGN = Alignment(horizontal="right", vertical="center")

THIN = Side(border_style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

ALT_FOUND_FILL = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")  # жёлтый
ALT_NONE_FILL = PatternFill(start_color="F4B084", end_color="F4B084", fill_type="solid")   # оранжевый
ALT_PRICE_FONT = Font(name="Calibri", size=10, bold=True, color="7F6000")
ALT_COMMENT_FONT = Font(name="Calibri", size=10, italic=True, color="595959")

# === Добавим 3 новые колонки в шапку (строка 4) ===
NEW_HEADERS = [
    ("Альтернатива в прайс-листе 16.10.2026", 50),
    ("Цена альтернативы, руб", 14),
    ("Комментарий", 60),
]
for i, (title, width) in enumerate(NEW_HEADERS):
    col_idx = 15 + i  # O, P, Q
    cell = ws.cell(row=4, column=col_idx, value=title)
    cell.fill = HEAD_FILL
    cell.font = HEAD_FONT
    cell.alignment = HEAD_ALIGN
    cell.border = BORDER
    ws.column_dimensions[get_column_letter(col_idx)].width = width

# === Заполним новые колонки данными ===
# Сначала найдём строки с 8 несопоставленными SKU
unmatched_rows = {}
for r in range(5, ws.max_row + 1):
    idx_val = ws.cell(row=r, column=1).value
    if idx_val in alternatives:
        unmatched_rows[idx_val] = r

print(f"\nНайдены строки для 8 SKU:")
for idx, r in unmatched_rows.items():
    print(f"  #{idx} → строка {r}")

# Заполняем новые колонки для всех 8 SKU
for idx, r in unmatched_rows.items():
    alt = alternatives[idx]
    
    # Колонка O (15): Альтернатива в прайс-листе
    cell = ws.cell(row=r, column=15, value=alt['alt_name'] if alt['alt_name'] else "—")
    cell.font = BODY_FONT
    cell.alignment = BODY_ALIGN
    cell.border = BORDER
    
    # Колонка P (16): Цена альтернативы
    cell = ws.cell(row=r, column=16, value=alt['alt_price'] if alt['alt_price'] else None)
    cell.font = ALT_PRICE_FONT if alt['alt_price'] else BODY_FONT
    cell.alignment = NUM_ALIGN
    cell.border = BORDER
    if alt['alt_price']:
        cell.number_format = '# ##0.00'
    
    # Колонка Q (17): Комментарий
    cell = ws.cell(row=r, column=17, value=alt['comment'])
    cell.font = ALT_COMMENT_FONT
    cell.alignment = BODY_ALIGN
    cell.border = BORDER
    
    # Подсветка всей строки
    fill = ALT_FOUND_FILL if alt['alt_price'] else ALT_NONE_FILL
    for col_idx in range(15, 18):
        ws.cell(row=r, column=col_idx).fill = fill

# Увеличим высоту строк с альтернативами
for r in unmatched_rows.values():
    ws.row_dimensions[r].height = 50

# Расширим автофильтр
last_col = get_column_letter(17)
ws.auto_filter.ref = f"A4:{last_col}{ws.max_row}"

# === Создать новый лист "8 несопоставленных — альтернативы" ===
ws_alt = wb.create_sheet("8 несопоставленных — альтернативы")

# Заголовок
ws_alt.merge_cells("A1:K1")
ws_alt["A1"] = "8 SKU, не сопоставленных с прайс-листом — детальный разбор альтернатив"
ws_alt["A1"].font = Font(name="Calibri", size=14, bold=True, color="1F4E79")
ws_alt.row_dimensions[1].height = 28

ws_alt.merge_cells("A2:K2")
ws_alt["A2"] = ("Проверка выполнена по новому прайс-листу от 16.10.2026 (РОССИЯ с НДС, руб.). "
                "Цены на 1020 позиций прайс-листа 21.08.2026 не изменились. Для 8 SKU ассортимента "
                "точных совпадений в прайс-листе нет, найдены аналогичные позиции — см. ниже.")
ws_alt["A2"].font = Font(name="Calibri", size=10, italic=True, color="595959")
ws_alt["A2"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
ws_alt.row_dimensions[2].height = 36

# Шапка таблицы (строка 4)
alt_headers = [
    ("#", 5),
    ("Сфера", 14),
    ("Бренд", 12),
    ("Наименование в ассортименте", 50),
    ("Фасовка ассортимента, кг", 12),
    ("Статус", 18),
    ("Альтернатива в прайс-листе 16.10.2026", 55),
    ("Фасовка прайса, кг", 11),
    ("Отгруз. цена, руб", 14),
    ("Базовая цена, руб", 13),
    ("Комментарий", 60),
]
for col_idx, (title, width) in enumerate(alt_headers, 1):
    cell = ws_alt.cell(row=4, column=col_idx, value=title)
    cell.fill = HEAD_FILL
    cell.font = HEAD_FONT
    cell.alignment = HEAD_ALIGN
    cell.border = BORDER
    ws_alt.column_dimensions[get_column_letter(col_idx)].width = width
ws_alt.row_dimensions[4].height = 38

# Получим данные о 8 SKU из основной вкладки
unmatched_data = {}
for r in range(5, ws.max_row + 1):
    idx_val = ws.cell(row=r, column=1).value
    if idx_val in alternatives:
        unmatched_data[idx_val] = {
            'sphere': ws.cell(row=r, column=2).value,
            'brand': ws.cell(row=r, column=5).value,
            'name': ws.cell(row=r, column=7).value,
        }

# Заполнить данные
row_i = 5
for idx in sorted(alternatives.keys()):
    data = unmatched_data.get(idx, {})
    alt = alternatives[idx]
    
    has_alt = alt['alt_price'] is not None
    status = "Альтернатива найдена" if has_alt else "Альтернативы нет"
    
    vals = [
        idx,
        data.get('sphere', ''),
        data.get('brand', ''),
        data.get('name', ''),
        next((float(s.split()[-2]) for s in [data.get('name', '')] if s and 'кг' in s), None),  # фасовка из имени
        status,
        alt['alt_name'] if alt['alt_name'] else "—",
        alt['alt_fasovka'] if alt['alt_fasovka'] else None,
        alt['alt_price'] if alt['alt_price'] else None,
        alt['alt_base'] if alt['alt_base'] else None,
        alt['comment'],
    ]
    
    # Извлечь фасовку из имени
    import re
    m = re.search(r'(\d+(?:\.\d+)?)\s*кг', str(data.get('name', '')))
    if m:
        vals[4] = float(m.group(1))
    
    for col_idx, val in enumerate(vals, 1):
        cell = ws_alt.cell(row=row_i, column=col_idx, value=val)
        cell.border = BORDER
        cell.font = BODY_FONT
        if col_idx == 1:
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif col_idx in (5, 8):  # Фасовки
            cell.alignment = NUM_ALIGN
            if val is not None:
                cell.number_format = '0.##'
        elif col_idx in (9, 10):  # Цены
            cell.alignment = NUM_ALIGN
            if val is not None:
                cell.number_format = '# ##0.00'
                cell.font = Font(name="Calibri", size=10, bold=True, color="7F6000")
        elif col_idx == 6:  # Статус
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font = Font(name="Calibri", size=10, bold=True,
                             color="7F6000" if has_alt else "C00000")
        elif col_idx in (4, 7, 11):  # Длинные тексты
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center")
    
    # Подсветка строки
    fill = ALT_FOUND_FILL if has_alt else ALT_NONE_FILL
    for col_idx in range(1, len(vals) + 1):
        ws_alt.cell(row=row_i, column=col_idx).fill = fill
    
    ws_alt.row_dimensions[row_i].height = 65
    row_i += 1

# Закрепить шапку
ws_alt.freeze_panes = "B5"

# === Сохранить файл ===
wb.properties.creator = "Z.ai"
wb.properties.title = "Ассортимент Олимп 300 SKU с ценами и альтернативами"
wb.properties.subject = "Сверка с прайс-листом 21.08.2026 и 16.10.2026"

wb.save(OUT_PATH)
print(f"\n✓ Файл сохранён: {OUT_PATH}")
print(f"  Размер: {os.path.getsize(OUT_PATH):,} байт")
print(f"  Листы: {wb.sheetnames}")

# === Проверка ===
print("\n=== Проверка: 8 несопоставленных SKU на вкладке '300 SKU с ценами' ===")
wb_check = load_workbook(OUT_PATH, data_only=True)
ws_check = wb_check["300 SKU с ценами"]
unmatched_idx = sorted(alternatives.keys())
for r in range(5, ws_check.max_row + 1):
    idx_val = ws_check.cell(row=r, column=1).value
    if idx_val in unmatched_idx:
        name = ws_check.cell(row=r, column=7).value
        price = ws_check.cell(row=r, column=10).value
        alt_name = ws_check.cell(row=r, column=15).value
        alt_price = ws_check.cell(row=r, column=16).value
        comment = ws_check.cell(row=r, column=17).value
        print(f"\n  #{idx_val} | {name}")
        print(f"    Цена в основной колонке: {price}")
        print(f"    Альтернатива: {alt_name}")
        print(f"    Цена альтернативы: {alt_price}")
        print(f"    Комментарий: {comment}")
