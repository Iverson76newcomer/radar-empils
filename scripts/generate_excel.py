"""
Генерация итогового Excel-файла: ассортимент + колонка 'Отгрузочная цена, руб'.
"""
import json
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# === Загрузка результатов ===
with open('/home/z/my-project/scripts/results.json', 'r', encoding='utf-8') as f:
    results = json.load(f)

print(f"Загружено {len(results)} позиций")

# Открываем оригинальный файл ассортимента (чтобы сохранить листы)
src_path = "/home/z/my-project/upload/Ассортимент_Олимп_300_SKU.xlsx"
out_path = "/home/z/my-project/download/Ассортимент_Олимп_300_SKU_с_отгрузочными_ценами.xlsx"

# Создаём новый workbook (openpyxl плохо сохраняет стили исходного)
from openpyxl import Workbook
wb = Workbook()
ws = wb.active
ws.title = "300 SKU — с ценами"

# === Стили ===
HEAD_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
HEAD_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
HEAD_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)

TITLE_FONT = Font(name="Calibri", size=14, bold=True, color="1F4E79")
SUBTITLE_FONT = Font(name="Calibri", size=10, italic=True, color="595959")

BODY_FONT = Font(name="Calibri", size=10)
BODY_ALIGN = Alignment(horizontal="left", vertical="center", wrap_text=True)
NUM_ALIGN = Alignment(horizontal="right", vertical="center")
CENTER_ALIGN = Alignment(horizontal="center", vertical="center")

THIN = Side(border_style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

NOT_FOUND_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
MATCHED_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
PRICE_FONT = Font(name="Calibri", size=10, bold=True, color="1F4E79")
NOT_FOUND_FONT = Font(name="Calibri", size=10, italic=True, color="C00000")

# === Заголовок ===
ws.merge_cells("A1:N1")
ws["A1"] = "Стартовый ассортимент Олимп — 300 SKU с отгрузочными ценами"
ws["A1"].font = TITLE_FONT
ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
ws.row_dimensions[1].height = 28

ws.merge_cells("A2:N2")
matched_count = sum(1 for r in results if r.get('отгрузочная') is not None)
ws["A2"] = (f"Источник цен: прайс-лист «21.08.2026 РОССИЯ с НДС, руб.» | "
            f"Сопоставлено: {matched_count} из {len(results)} SKU | "
            f"Не найдено в прайс-листе: {len(results) - matched_count}")
ws["A2"].font = SUBTITLE_FONT
ws["A2"].alignment = Alignment(horizontal="left", vertical="center")
ws.row_dimensions[2].height = 18

# === Шапка (строка 4) ===
HEADERS = [
    ("#", 5),
    ("Сфера", 14),
    ("Назначение", 28),
    ("Основа", 16),
    ("Бренд", 14),
    ("Тип продукции", 22),
    ("Наименование", 60),
    ("Гр. код", 9),
    ("ABC", 6),
    ("Отгрузочная цена, руб", 14),
    ("Базовая цена, руб", 13),
    ("Фасовка в прайсе, кг", 11),
    ("Лист прайса", 13),
    ("Совпадение", 11),
]

HEADER_ROW = 4
for col_idx, (title, width) in enumerate(HEADERS, 1):
    cell = ws.cell(row=HEADER_ROW, column=col_idx, value=title)
    cell.fill = HEAD_FILL
    cell.font = HEAD_FONT
    cell.alignment = HEAD_ALIGN
    cell.border = BORDER
    ws.column_dimensions[get_column_letter(col_idx)].width = width

ws.row_dimensions[HEADER_ROW].height = 38

# === Данные ===
DATA_START_ROW = 5
row_idx = DATA_START_ROW

for r in results:
    matched = r.get('отгрузочная') is not None
    
    vals = [
        r['idx'],
        r.get('sphere', ''),
        r.get('purpose', ''),
        r.get('base', ''),
        r.get('brand', ''),
        r.get('type', ''),
        r.get('name', ''),
        r.get('grp_code', ''),
        r.get('abc', ''),
        r.get('отгрузочная') if matched else "не найдено в прайс-листе",
        r.get('базовая') if matched else None,
        r.get('фасовка_прайс') if matched else None,
        r.get('лист', '') if matched else '',
        r.get('match_type', '') if matched else '',
    ]
    
    for col_idx, val in enumerate(vals, 1):
        cell = ws.cell(row=row_idx, column=col_idx, value=val)
        cell.border = BORDER
        cell.font = BODY_FONT
        if col_idx in (1, 8, 9):  # #, Гр.код, ABC
            cell.alignment = CENTER_ALIGN
        elif col_idx in (10, 11, 12):  # Числа (цены, фасовка)
            cell.alignment = NUM_ALIGN
            if col_idx == 10:  # Отгрузочная цена
                if matched:
                    cell.font = PRICE_FONT
                    cell.number_format = '# ##0.00'
                else:
                    cell.font = NOT_FOUND_FONT
            elif col_idx == 11 and val is not None:
                cell.number_format = '# ##0.00'
            elif col_idx == 12 and val is not None:
                cell.number_format = '0.##'
        elif col_idx == 7:  # Наименование
            cell.alignment = BODY_ALIGN
        elif col_idx == 13:  # Лист прайса
            cell.alignment = CENTER_ALIGN
        elif col_idx == 14:  # Совпадение
            cell.alignment = CENTER_ALIGN
        else:
            cell.alignment = BODY_ALIGN
    
    # Подсветка строки
    fill = MATCHED_FILL if matched else NOT_FOUND_FILL
    for col_idx in range(1, len(HEADERS) + 1):
        if ws.cell(row=row_idx, column=col_idx).fill.start_color.rgb in (None, '00000000', 'FFFFFFFF'):
            ws.cell(row=row_idx, column=col_idx).fill = fill
        else:
            # Сохранить существующий fill ячейки (например, у заголовков)
            pass
    # Применяем fill ко всем ячейкам строки
    for col_idx in range(1, len(HEADERS) + 1):
        ws.cell(row=row_idx, column=col_idx).fill = fill
    
    row_idx += 1

# === Закрепить заголовок и первую колонку ===
ws.freeze_panes = "B5"

# === Автофильтр ===
last_col = get_column_letter(len(HEADERS))
last_row = DATA_START_ROW + len(results) - 1
ws.auto_filter.ref = f"A{HEADER_ROW}:{last_col}{last_row}"

# === Сводный лист ===
ws2 = wb.create_sheet("Сводка по сопоставлению")

ws2.merge_cells("A1:E1")
ws2["A1"] = "Сводка по сопоставлению с прайс-листом"
ws2["A1"].font = TITLE_FONT
ws2.row_dimensions[1].height = 24

# Блок 1: общая статистика
ws2["A3"] = "ОБЩАЯ СТАТИСТИКА"
ws2["A3"].font = Font(name="Calibri", size=12, bold=True, color="1F4E79")
ws2.merge_cells("A3:E3")

stats = [
    ("Всего SKU в ассортименте", len(results)),
    ("Сопоставлено с прайс-листом", matched_count),
    ("Не найдено в прайс-листе", len(results) - matched_count),
    ("Доля сопоставления, %", round(matched_count * 100 / len(results), 1)),
]
for i, (label, val) in enumerate(stats, 4):
    ws2.cell(row=i, column=1, value=label).font = Font(name="Calibri", size=10, bold=True)
    ws2.cell(row=i, column=2, value=val).font = Font(name="Calibri", size=10)
    ws2.cell(row=i, column=2).alignment = NUM_ALIGN

# Блок 2: по брендам
start_row = 9
ws2.cell(row=start_row, column=1, value="СОПОСТАВЛЕНИЕ ПО БРЕНДАМ").font = Font(name="Calibri", size=12, bold=True, color="1F4E79")
ws2.merge_cells(f"A{start_row}:E{start_row}")

headers2 = ["Бренд", "Всего SKU", "Сопоставлено", "Не сопоставлено", "Доля, %"]
for col_idx, h in enumerate(headers2, 1):
    cell = ws2.cell(row=start_row + 1, column=col_idx, value=h)
    cell.fill = HEAD_FILL
    cell.font = HEAD_FONT
    cell.alignment = HEAD_ALIGN
    cell.border = BORDER
ws2.row_dimensions[start_row + 1].height = 24

# Группировка
from collections import defaultdict
brand_stats = defaultdict(lambda: {'total': 0, 'matched': 0})
for r in results:
    b = r.get('brand', '')
    brand_stats[b]['total'] += 1
    if r.get('отгрузочная') is not None:
        brand_stats[b]['matched'] += 1

row_i = start_row + 2
for b in sorted(brand_stats.keys()):
    s = brand_stats[b]
    pct = round(s['matched'] * 100 / s['total'], 1) if s['total'] else 0
    vals = [b, s['total'], s['matched'], s['total'] - s['matched'], pct]
    for col_idx, val in enumerate(vals, 1):
        cell = ws2.cell(row=row_i, column=col_idx, value=val)
        cell.border = BORDER
        cell.font = BODY_FONT
        if col_idx == 1:
            cell.alignment = Alignment(horizontal="left", vertical="center")
        else:
            cell.alignment = NUM_ALIGN
        # Подсветка
        if s['matched'] == s['total']:
            cell.fill = MATCHED_FILL
        elif s['matched'] == 0:
            cell.fill = NOT_FOUND_FILL
    row_i += 1

# Итого
totals = [
    "ИТОГО",
    sum(s['total'] for s in brand_stats.values()),
    sum(s['matched'] for s in brand_stats.values()),
    sum(s['total'] - s['matched'] for s in brand_stats.values()),
    round(sum(s['matched'] for s in brand_stats.values()) * 100 / sum(s['total'] for s in brand_stats.values()), 1),
]
for col_idx, val in enumerate(totals, 1):
    cell = ws2.cell(row=row_i, column=col_idx, value=val)
    cell.font = Font(name="Calibri", size=10, bold=True)
    cell.fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
    cell.border = BORDER
    if col_idx == 1:
        cell.alignment = Alignment(horizontal="left", vertical="center")
    else:
        cell.alignment = NUM_ALIGN

# Ширины колонок на сводном листе
ws2.column_dimensions['A'].width = 22
ws2.column_dimensions['B'].width = 14
ws2.column_dimensions['C'].width = 14
ws2.column_dimensions['D'].width = 16
ws2.column_dimensions['E'].width = 10

# Блок 3: список несопоставленных
row_i += 3
ws2.cell(row=row_i, column=1, value="НЕ СОПОСТАВЛЕННЫЕ SKU (отсутствуют в прайс-листе)").font = Font(name="Calibri", size=12, bold=True, color="C00000")
ws2.merge_cells(f"A{row_i}:E{row_i}")
row_i += 1

unmatched_headers = ["#", "Бренд", "Наименование", "Причина"]
for col_idx, h in enumerate(unmatched_headers, 1):
    cell = ws2.cell(row=row_i, column=col_idx, value=h)
    cell.fill = HEAD_FILL
    cell.font = HEAD_FONT
    cell.alignment = HEAD_ALIGN
    cell.border = BORDER
ws2.row_dimensions[row_i].height = 24
row_i += 1

for r in results:
    if r.get('отгрузочная') is None:
        name = r.get('name', '')
        # Определим причину
        reason = "Цвет/серия не найден в прайс-листе"
        if 'ABSOLUTE BLACK' in name or 'АКВАСТОП' in name:
            reason = "Новинка / специальная серия отсутствует в прайс-листе"
        vals = [r['idx'], r.get('brand', ''), name, reason]
        for col_idx, val in enumerate(vals, 1):
            cell = ws2.cell(row=row_i, column=col_idx, value=val)
            cell.border = BORDER
            cell.font = BODY_FONT
            cell.fill = NOT_FOUND_FILL
            if col_idx == 1:
                cell.alignment = CENTER_ALIGN
            elif col_idx == 4:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            else:
                cell.alignment = BODY_ALIGN
        row_i += 1

# Увеличим ширину колонки D на сводном
ws2.column_dimensions['D'].width = 70
ws2.column_dimensions['E'].width = 50

# Перенесём сводный лист в начало
wb.move_sheet("Сводка по сопоставлению", offset=-1)

# Удалим дефолтный лист, если их несколько
if "Sheet" in wb.sheetnames:
    del wb["Sheet"]

# === Метаданные ===
wb.properties.creator = "Z.ai"
wb.properties.title = "Ассортимент Олимп 300 SKU с отгрузочными ценами"
wb.properties.subject = "Сопоставление с прайс-листом от 21.08.2026"

# Сохранение
os.makedirs(os.path.dirname(out_path), exist_ok=True)
wb.save(out_path)
print(f"\nФайл сохранён: {out_path}")

# Валидация: открываем заново
wb2 = load_workbook(out_path, data_only=True)
ws_check = wb2["300 SKU — с ценами"]
print(f"\n=== Проверка ===")
print(f"Листы: {wb2.sheetnames}")
print(f"Лист '300 SKU — с ценами': {ws_check.max_row} строк × {ws_check.max_column} колонок")
# Подсчёт сопоставленных
matched_check = 0
unmatched_check = 0
for row in ws_check.iter_rows(min_row=5, max_row=ws_check.max_row, values_only=True):
    if row[0] is None:
        continue
    if row[9] == "не найдено в прайс-листе" or row[9] is None:
        unmatched_check += 1
    else:
        matched_check += 1
print(f"Сопоставлено: {matched_check}")
print(f"Не сопоставлено: {unmatched_check}")

# Примеры
print("\n=== Первые 5 строк ===")
for i, row in enumerate(ws_check.iter_rows(min_row=5, max_row=9, values_only=True), 5):
    print(f"  Строка {i}: #{row[0]} | {row[4]} | {str(row[6])[:50]}... | отгруз={row[9]}")

print("\n=== Последние 3 несопоставленных ===")
cnt = 0
for row in ws_check.iter_rows(min_row=5, max_row=ws_check.max_row, values_only=True):
    if row[9] == "не найдено в прайс-листе":
        print(f"  #{row[0]} | {row[4]} | {row[6]}")
        cnt += 1
        if cnt >= 8:
            break
