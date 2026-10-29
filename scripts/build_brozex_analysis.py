"""Создание файла анализа статистики TITAN — БРОЗЭКС.
   Аналогично файлу TITAN_Рыгалов_анализ_статистики.xlsx
"""
import json
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Загружаем собранную статистику
with open('/home/z/my-project/scripts/brozex_stats.json', 'r', encoding='utf-8') as f:
    stats = json.load(f)

# Дополняем АКБ из итоговых строк файла-источника (это количество ТТ в классе)
akb_totals = {
    'A':             {'akb25': 3,  'akb26': 6},
    'B':             {'akb25': 8,  'akb26': 14},
    'C':             {'akb25': 12, 'akb26': 16},
    'D':             {'akb25': 34, 'akb26': 31},
    'E':             {'akb25': 13, 'akb26': 23},
    'F':             {'akb25': 1,  'akb26': 1},
    'G':             {'akb25': 5,  'akb26': 2},
    'J':             {'akb25': 5,  'akb26': 4},
    'Unclassified':  {'akb25': 0,  'akb26': 11},
}

# MML (рекомендованный) — только для ключевых классов B, C, D
mml = {'B': 23, 'C': 16, 'D': 6}

# Порядок классов и их описания
classes_order = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'J', 'Unclassified']
class_desc = {
    'A': 'Федеральная сеть / DIY-гипермаркет',
    'B': 'Региональная сеть (расширенная матрица)',
    'C': 'Супермаркет (топ-цвета + молотковая)',
    'D': 'Хозмаг (топ-6 цветов)',
    'E': 'Малая ТТ (ниже D)',
    'F': 'Очень малая ТТ',
    'G': 'Микро-ТТ',
    'J': 'Спец. канал (промышленные/проф. потребители)',
    'Unclassified': 'Прочие / не классифицированные ТТ',
}

# === Стилизация ===
ORANGE_FILL = PatternFill('solid', fgColor='F47629')
DARK_FILL = PatternFill('solid', fgColor='313131')
LIGHT_GRAY_FILL = PatternFill('solid', fgColor='F8F8F8')
WHITE_FONT = Font(name='Arial', size=11, bold=True, color='FFFFFF')
ORANGE_FONT = Font(name='Arial', size=11, bold=True, color='F47629')
RED_FONT = Font(name='Arial', size=11, bold=True, color='C0392B')
GREEN_FONT = Font(name='Arial', size=11, bold=True, color='16A34A')
NORMAL_FONT = Font(name='Arial', size=10)
TITLE_FONT = Font(name='Arial', size=14, bold=True, color='313131')
SUBTITLE_FONT = Font(name='Arial', size=10, italic=True, color='666666')
MML_FONT = Font(name='Arial', size=10, bold=True, color='1E3A5F')
THIN_BORDER = Border(
    left=Side(style='thin', color='D0D0D0'),
    right=Side(style='thin', color='D0D0D0'),
    top=Side(style='thin', color='D0D0D0'),
    bottom=Side(style='thin', color='D0D0D0')
)
EMPTY_FILL = PatternFill('solid', fgColor='FFFDE7')  # светло-жёлтый — для заполнения

CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)

# Создаём новую книгу
wb = Workbook()

# ============================================================
# ЛИСТ 1: Сводка по классам
# ============================================================
ws1 = wb.active
ws1.title = 'Сводка по классам'

# Заголовок (строки 1-2)
ws1.cell(row=1, column=1, value='TITAN — Статистика БРОЗЭКС').font = TITLE_FONT
ws1.merge_cells(start_row=1, start_column=1, end_row=1, end_column=12)
ws1.cell(row=1, column=1).alignment = Alignment(horizontal='left', vertical='center')

ws1.cell(row=2, column=1, value='Анализ 2025 vs 2026 + поля для замеров после программы').font = SUBTITLE_FONT
ws1.merge_cells(start_row=2, start_column=1, end_row=2, end_column=12)
ws1.cell(row=2, column=1).alignment = Alignment(horizontal='left', vertical='center')

ws1.cell(row=3, column=1, value='Источник: выгрузка БРОЗЭКС 09.09.2026 | Критерий: соблюдение рекомендованного MML TITAN в точках класса B, C, D').font = SUBTITLE_FONT
ws1.merge_cells(start_row=3, start_column=1, end_row=3, end_column=12)
ws1.cell(row=3, column=1).alignment = Alignment(horizontal='left', vertical='center')

ws1.row_dimensions[1].height = 22
ws1.row_dimensions[2].height = 18
ws1.row_dimensions[3].height = 18

# Заголовок таблицы (строка 4)
new_headers = ['Класс ТТ', 'Описание', 'MML (реком.)', 'SKU 2025', 'SKU 2026', 'Δ SKU',
               'АКБ 2025', 'АКБ 2026', 'Δ АКБ',
               'Продажи 2025, ₽', 'Продажи 2026, ₽', 'Δ продаж, %']
for ci, h in enumerate(new_headers, 1):
    c = ws1.cell(row=4, column=ci, value=h)
    c.font = WHITE_FONT; c.fill = DARK_FILL
    c.alignment = CENTER; c.border = THIN_BORDER
ws1.row_dimensions[4].height = 38

# Данные
tot_sku25 = tot_sku26 = 0
tot_akb25 = tot_akb26 = 0
tot_sales25 = tot_sales26 = 0.0

for ri, cls in enumerate(classes_order, 1):
    # База данных — ключ в JSON "Unknown", в выводе отображаем как "Unclassified"
    json_key = 'Unknown' if cls == 'Unclassified' else cls
    d = stats.get(json_key, {'sku25': 0, 'sku26': 0, 'sales25': 0, 'sales26': 0})
    mml_val = mml.get(cls, '—')
    sku25 = d['sku25']; sku26 = d['sku26']
    delta_sku = sku26 - sku25
    akb25 = akb_totals[cls]['akb25']; akb26 = akb_totals[cls]['akb26']
    delta_akb = akb26 - akb25
    sales25 = d['sales25']; sales26 = d['sales26']
    delta_sales = ((sales26 - sales25) / sales25 * 100) if sales25 else (100.0 if sales26 > 0 else 0)

    row_data = [cls, class_desc[cls], mml_val, sku25, sku26, delta_sku,
                akb25, akb26, delta_akb,
                f"{sales25:,.0f}", f"{sales26:,.0f}", f'{delta_sales:+.1f}%']
    for ci, v in enumerate(row_data, 1):
        c = ws1.cell(row=4 + ri, column=ci, value=v)
        c.font = NORMAL_FONT; c.border = THIN_BORDER
        c.alignment = CENTER if ci != 2 else LEFT
        if ci == 1:
            c.font = ORANGE_FONT
        elif ci == 3:
            c.font = MML_FONT if isinstance(mml_val, int) else NORMAL_FONT
        elif ci in (6, 9):
            val = v if isinstance(v, (int, float)) else 0
            if val > 0: c.font = GREEN_FONT
            elif val < 0: c.font = RED_FONT
        elif ci == 12:
            if delta_sales > 0: c.font = GREEN_FONT
            elif delta_sales < 0: c.font = RED_FONT

    tot_sku25 += sku25; tot_sku26 += sku26
    tot_akb25 += akb25; tot_akb26 += akb26
    tot_sales25 += sales25; tot_sales26 += sales26

# Итог
row_idx = 4 + len(classes_order) + 1
tot_delta_sales = ((tot_sales26 - tot_sales25) / tot_sales25 * 100) if tot_sales25 else 0
row_data = ['ИТОГО', 'Все классы', '—', tot_sku25, tot_sku26, tot_sku26 - tot_sku25,
            tot_akb25, tot_akb26, tot_akb26 - tot_akb25,
            f'{tot_sales25:,.0f}', f'{tot_sales26:,.0f}', f'{tot_delta_sales:+.1f}%']
for ci, v in enumerate(row_data, 1):
    c = ws1.cell(row=row_idx, column=ci, value=v)
    c.font = Font(name='Arial', size=10, bold=True, color='FFFFFF')
    c.fill = ORANGE_FILL; c.border = THIN_BORDER
    c.alignment = CENTER if ci != 2 else LEFT

# Ширины
widths = {'A': 12, 'B': 36, 'C': 12, 'D': 10, 'E': 10, 'F': 10,
          'G': 10, 'H': 10, 'I': 10, 'J': 16, 'K': 16, 'L': 12}
for col, w in widths.items():
    ws1.column_dimensions[col].width = w

ws1.freeze_panes = 'A5'

# ============================================================
# ЛИСТ 2: Замеры после программы
# ============================================================
ws2 = wb.create_sheet('Замеры после программы')

ws2.cell(row=1, column=1, value='ЗАМЕРЫ ПОСЛЕ ПРОГРАММЫ TITAN — жёлтые поля для заполнения').font = TITLE_FONT
ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=11)
ws2.cell(row=2, column=1, value='Жёлтые поля — для заполнения после завершения программы (план: 6 месяцев)').font = SUBTITLE_FONT
ws2.merge_cells(start_row=2, start_column=1, end_row=2, end_column=11)
ws2.row_dimensions[1].height = 22
ws2.row_dimensions[2].height = 18

new_headers2 = ['Класс ТТ', 'MML (реком.)', 'SKU до (2026)', 'SKU после',
                'Δ SKU', 'АКБ до (2026)', 'АКБ после', 'Δ АКБ',
                'Продажи до, ₽ (2026)', 'Продажи после, ₽', 'Δ продаж, %']
for ci, h in enumerate(new_headers2, 1):
    c = ws2.cell(row=4, column=ci, value=h)
    c.font = WHITE_FONT; c.fill = DARK_FILL
    c.alignment = CENTER; c.border = THIN_BORDER
ws2.row_dimensions[4].height = 38

for ri, cls in enumerate(classes_order, 1):
    json_key = 'Unknown' if cls == 'Unclassified' else cls
    d = stats.get(json_key, {'sku26': 0, 'sales26': 0})
    mml_val = mml.get(cls, '—')
    akb26 = akb_totals[cls]['akb26']
    row_data = [cls, mml_val, d['sku26'], '', '', akb26, '', '',
                f"{d['sales26']:,.0f}", '', '']
    for ci, v in enumerate(row_data, 1):
        c = ws2.cell(row=4 + ri, column=ci, value=v)
        c.font = NORMAL_FONT; c.border = THIN_BORDER
        c.alignment = CENTER
        if ci in (4, 5, 7, 8, 10, 11):
            c.fill = EMPTY_FILL
        if ci == 1:
            c.font = ORANGE_FONT
        elif ci == 2 and isinstance(mml_val, int):
            c.font = MML_FONT

# Итог
row_idx2 = 4 + len(classes_order) + 1
row_data2 = ['ИТОГО', '—', tot_sku26, '', '', tot_akb26, '', '',
             f'{tot_sales26:,.0f}', '', '']
for ci, v in enumerate(row_data2, 1):
    c = ws2.cell(row=row_idx2, column=ci, value=v)
    c.font = Font(name='Arial', size=10, bold=True, color='FFFFFF')
    c.fill = ORANGE_FILL; c.border = THIN_BORDER
    c.alignment = CENTER

widths2 = {'A': 12, 'B': 12, 'C': 14, 'D': 14, 'E': 10, 'F': 14, 'G': 14, 'H': 10,
           'I': 18, 'J': 18, 'K': 12}
for col, w in widths2.items():
    ws2.column_dimensions[col].width = w

ws2.freeze_panes = 'A5'

# ============================================================
# ЛИСТ 3: Тренды 2025→2026
# ============================================================
ws3 = wb.create_sheet('Тренды 2025→2026')

ws3.cell(row=1, column=1, value='ТРЕНДЫ 2025 → 2026 — анализ изменений').font = TITLE_FONT
ws3.merge_cells(start_row=1, start_column=1, end_row=1, end_column=12)
ws3.row_dimensions[1].height = 22

new_headers3 = ['Класс', 'SKU 2025', 'SKU 2026', 'Δ SKU', 'Тренд SKU',
                'АКБ 2025', 'АКБ 2026', 'Δ АКБ', 'Тренд АКБ',
                'Продажи 2025, ₽', 'Продажи 2026, ₽', 'Тренд продаж']
for ci, h in enumerate(new_headers3, 1):
    c = ws3.cell(row=3, column=ci, value=h)
    c.font = WHITE_FONT; c.fill = DARK_FILL
    c.alignment = CENTER; c.border = THIN_BORDER
ws3.row_dimensions[3].height = 38

for ri, cls in enumerate(classes_order, 1):
    json_key = 'Unknown' if cls == 'Unclassified' else cls
    d = stats.get(json_key, {'sku25': 0, 'sku26': 0, 'sales25': 0, 'sales26': 0})
    sku25 = d['sku25']; sku26 = d['sku26']
    delta_sku = sku26 - sku25
    akb25 = akb_totals[cls]['akb25']; akb26 = akb_totals[cls]['akb26']
    delta_akb = akb26 - akb25
    sales25 = d['sales25']; sales26 = d['sales26']
    delta_sales = ((sales26 - sales25) / sales25 * 100) if sales25 else (100.0 if sales26 > 0 else 0)

    trend_sku = '🟢 Рост' if delta_sku > 0 else ('🔴 Снижение' if delta_sku < 0 else '⚪ Стабильно')
    trend_akb = '🟢 Рост' if delta_akb > 0 else ('🔴 Снижение' if delta_akb < 0 else '⚪ Стабильно')
    if delta_sales > 10:
        trend_sales = '🟢 Рост'
    elif delta_sales < -10:
        trend_sales = '🔴 Снижение'
    else:
        trend_sales = '🟡 Стабильно'

    row_data = [cls, sku25, sku26, delta_sku, trend_sku,
                akb25, akb26, delta_akb, trend_akb,
                f"{sales25:,.0f}", f"{sales26:,.0f}", f'{delta_sales:+.1f}% {trend_sales}']
    for ci, v in enumerate(row_data, 1):
        c = ws3.cell(row=3 + ri, column=ci, value=v)
        c.font = NORMAL_FONT; c.border = THIN_BORDER
        c.alignment = CENTER
        if ci == 1:
            c.font = ORANGE_FONT
        elif ci in (4, 8):
            val = v if isinstance(v, (int, float)) else 0
            if val > 0: c.font = GREEN_FONT
            elif val < 0: c.font = RED_FONT
        elif ci == 12:
            if 'Рост' in v and '🟢' in v:
                c.font = GREEN_FONT
            elif 'Снижение' in v:
                c.font = RED_FONT

# Итог
row_idx3 = 3 + len(classes_order) + 1
tot_delta_akb = tot_akb26 - tot_akb25
tot_delta_sales = ((tot_sales26 - tot_sales25) / tot_sales25 * 100) if tot_sales25 else 0
trend_akb_tot = '🟢 Рост' if tot_delta_akb > 0 else ('🔴 Снижение' if tot_delta_akb < 0 else '⚪ Стабильно')
if tot_delta_sales > 10:
    trend_sales_tot = '🟢 Рост'
elif tot_delta_sales < -10:
    trend_sales_tot = '🔴 Снижение'
else:
    trend_sales_tot = '🟡 Стабильно'
row_data3 = ['ИТОГО', tot_sku25, tot_sku26, tot_sku26 - tot_sku25, '—',
             tot_akb25, tot_akb26, tot_delta_akb, trend_akb_tot,
             f'{tot_sales25:,.0f}', f'{tot_sales26:,.0f}', f'{tot_delta_sales:+.1f}% {trend_sales_tot}']
for ci, v in enumerate(row_data3, 1):
    c = ws3.cell(row=row_idx3, column=ci, value=v)
    c.font = Font(name='Arial', size=10, bold=True, color='FFFFFF')
    c.fill = ORANGE_FILL; c.border = THIN_BORDER
    c.alignment = CENTER

widths3 = {'A': 12, 'B': 10, 'C': 10, 'D': 10, 'E': 14, 'F': 10, 'G': 10, 'H': 10, 'I': 14,
           'J': 16, 'K': 16, 'L': 22}
for col, w in widths3.items():
    ws3.column_dimensions[col].width = w

ws3.freeze_panes = 'A4'

# ============================================================
# ЛИСТЫ 4+: Детализация по классам B, C, D (где задан MML)
# ============================================================
for cls in ['B', 'C', 'D']:
    ws = wb.create_sheet(f'Детализация класс {cls}')
    ws.cell(row=1, column=1, value=f'ДЕТАЛИЗАЦИЯ ПО SKU — КЛАСС {cls}').font = TITLE_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)
    ws.cell(row=2, column=1, value=f'MML (рекомендованный): {mml[cls]} SKU | АКБ = 1, если SKU был в продажах периода, иначе —').font = SUBTITLE_FONT
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=5)
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 18

    headers = ['SKU', 'АКБ 2025', 'Продажи 2025, ₽', 'АКБ 2026', 'Продажи 2026, ₽']
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=4, column=ci, value=h)
        c.font = WHITE_FONT; c.fill = DARK_FILL
        c.alignment = CENTER; c.border = THIN_BORDER
    ws.row_dimensions[4].height = 32

    d = stats[cls]
    sku_data = d['sku_data']
    sorted_skus = sorted(sku_data.keys())

    for ri, sku_name in enumerate(sorted_skus, 1):
        sd = sku_data[sku_name]
        akb25_val = 1 if (sd['akb25'] > 0 or sd['sales25'] > 0) else '—'
        akb26_val = 1 if (sd['akb26'] > 0 or sd['sales26'] > 0) else '—'
        sales25_val = f"{sd['sales25']:,.0f}" if sd['sales25'] > 0 else '—'
        sales26_val = f"{sd['sales26']:,.0f}" if sd['sales26'] > 0 else '—'
        row_data = [sku_name, akb25_val, sales25_val, akb26_val, sales26_val]
        for ci, v in enumerate(row_data, 1):
            c = ws.cell(row=4 + ri, column=ci, value=v)
            c.font = NORMAL_FONT; c.border = THIN_BORDER
            if ci == 1:
                c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            else:
                c.alignment = CENTER

    # Ширины
    ws.column_dimensions['A'].width = 75
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 12
    ws.column_dimensions['E'].width = 18
    ws.freeze_panes = 'A5'

# Сохраняем
out_path = '/home/z/my-project/download/TITAN_БРОЗЭКС_анализ_статистики.xlsx'
wb.save(out_path)
print(f'OK saved: {out_path}')
print(f'\n=== ИТОГИ ===')
print(f'Всего SKU 2025: {tot_sku25}, 2026: {tot_sku26}')
print(f'Всего ТТ (АКБ) 2025: {tot_akb25}, 2026: {tot_akb26}')
print(f'Всего продажи 2025: {tot_sales25:,.0f} ₽, 2026: {tot_sales26:,.0f} ₽')
print(f'Δ продаж: {tot_delta_sales:+.1f}%')
