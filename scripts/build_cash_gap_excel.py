"""
Excel-модель кассового разрыва.
"""
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.chart.label import DataLabelList
import json
import os

# Загрузить модель
with open('/home/z/my-project/scripts/cash_gap_model.json', 'r', encoding='utf-8') as f:
    model = json.load(f)

OUT_PATH = "/home/z/my-project/download/Кассовый_разрыв_Олимп_модель.xlsx"

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
PRICE_FONT = Font(name="Calibri", size=10, bold=True, color="1F4E79")
NEGATIVE_FONT = Font(name="Calibri", size=10, bold=True, color="C00000")
POSITIVE_FONT = Font(name="Calibri", size=10, bold=True, color="375623")
GAP_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
WARN_FILL = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
POSITIVE_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
TOTAL_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
PRIMARY_HIGHLIGHT = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")  # жёлтый — первичка
SECONDARY_HIGHLIGHT = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")  # голубой — вторичка

THIN = Side(border_style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

wb = Workbook()
wb.properties.creator = "Z.ai"
wb.properties.title = "Модель кассового разрыва — дистрибьютор Олимп"

# ==================== ЛИСТ 1: СВОДКА ====================
ws = wb.active
ws.title = "Сводка"

# Заголовок
ws.merge_cells("A1:H1")
ws["A1"] = "Модель кассового разрыва — дистрибьютор Олимп (Год 1)"
ws["A1"].font = TITLE_FONT
ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
ws.row_dimensions[1].height = 28

ws.merge_cells("A2:H2")
ws["A2"] = ("Источник: КП Эмпилс от 12.10.2026 (товарооборот 60 млн, наценка 35%, отсрочка 30-45 дн) + "
            "ассортимент v4 (300 SKU с ценами) + экспертная оценка сезонности ЛКМ")
ws["A2"].font = SUBTITLE_FONT
ws["A2"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
ws.row_dimensions[2].height = 36

# === Ключевые параметры ===
ws["A4"] = "КЛЮЧЕВЫЕ ПАРАМЕТРЫ МОДЕЛИ"
ws["A4"].font = Font(name="Calibri", size=12, bold=True, color="1F4E79")
ws.merge_cells("A4:H4")

params = [
    ("Годовой товарооборот (план)", "60 000 000 руб", "Из КП Эмпилс от 12.10.2026, Год 1"),
    ("Наценка дистрибьютора (Дистр→РТТ)", "35%", "Из КП: заводская цена × 1.35 = цена для РТТ"),
    ("Наценка РТТ (РТТ→Потребитель)", "50%", "Из КП: цена для РТТ × 1.50 = розничная цена"),
    ("Условия оплаты Эмпилс→Дистр", "Постоплата 45 дней", "Из КП: «диапазон 30-45 дней» — берём максимум 45"),
    ("Условия оплаты РТТ→Дистр", "Постоплата 30 дней", "Типовая рыночная практика по РТТ"),
    ("Стартовый капитал дистрибьютора", "0 руб", "Моделируем сценарий «без вливаний», чтобы найти разрыв"),
    ("Количество SKU в матрице", "300", "Ассортимент Олимп v4"),
    ("Сумма отгрузочных цен всех SKU", f"{sum(s['cost'] for s in [{'cost':p} for p in []]):,} руб", "Среднее по 300 SKU из прайс-листа 16.10.2026"),
]
# Считаем сумму всех цен SKU
import openpyxl
wb_src = openpyxl.load_workbook("/home/z/my-project/upload/Ассортимент_Олимп_300_SKU_финал_v4.xlsx", data_only=True)
ws_src = wb_src["300 SKU с ценами"]
total_sku_cost = 0
sku_count = 0
for r in range(5, ws_src.max_row + 1):
    price = ws_src.cell(row=r, column=10).value
    if price and isinstance(price, (int, float)):
        total_sku_cost += price
        sku_count += 1
avg_price = total_sku_cost / sku_count

params = [
    ("Годовой товарооборот (план)", "60 000 000 руб", "Из КП Эмпилс от 12.10.2026, Год 1"),
    ("Наценка дистрибьютора (Дистр→РТТ)", "35%", "Из КП: заводская цена × 1.35 = цена для РТТ"),
    ("Наценка РТТ (РТТ→Потребитель)", "50%", "Из КП: цена для РТТ × 1.50 = розничная цена"),
    ("Условия оплаты Эмпилс→Дистр", "Постоплата 45 дней", "Из КП: «диапазон 30-45 дней» — берём максимум 45"),
    ("Условия оплаты РТТ→Дистр", "Постоплата 30 дней", "Типовая рыночная практика по РТТ"),
    ("Стартовый капитал дистрибьютора", "0 руб", "Моделируем «без вливаний», чтобы увидеть разрыв"),
    ("Количество SKU в матрице", "300", "Ассортимент Олимп v4"),
    ("Сумма цен всех 300 SKU (по прайсу)", f"{total_sku_cost:,.0f} руб", "Средняя цена 1 SKU = " + f"{avg_price:,.0f} руб"),
]

# Шапка таблицы параметров
for col_idx, (title, width) in enumerate([("Параметр", 38), ("Значение", 24), ("Источник / комментарий", 50)], 1):
    cell = ws.cell(row=5, column=col_idx, value=title)
    cell.fill = HEAD_FILL
    cell.font = HEAD_FONT
    cell.alignment = HEAD_ALIGN
    cell.border = BORDER
    ws.column_dimensions[get_column_letter(col_idx)].width = width
ws.row_dimensions[5].height = 32

for i, (param, val, src) in enumerate(params):
    r = 6 + i
    ws.cell(row=r, column=1, value=param).font = Font(name="Calibri", size=10, bold=True)
    ws.cell(row=r, column=1).border = BORDER
    ws.cell(row=r, column=1).alignment = Alignment(horizontal="left", vertical="center")
    
    ws.cell(row=r, column=2, value=val).font = Font(name="Calibri", size=10, bold=True, color="1F4E79")
    ws.cell(row=r, column=2).border = BORDER
    ws.cell(row=r, column=2).alignment = CENTER_ALIGN
    
    ws.cell(row=r, column=3, value=src).font = BODY_FONT
    ws.cell(row=r, column=3).border = BORDER
    ws.cell(row=r, column=3).alignment = BODY_ALIGN

# === Ключевой вывод ===
r = 6 + len(params) + 2
ws.merge_cells(f"A{r}:H{r}")
ws.cell(row=r, column=1, value="КЛЮЧЕВОЙ ВЫВОД: КАССОВЫЙ РАЗРЫВ")
ws.cell(row=r, column=1).font = Font(name="Calibri", size=12, bold=True, color="C00000")

r += 1
ws.merge_cells(f"A{r}:H{r+2}")
ws.cell(row=r, column=1, value=(
    "✗ Кассовый разрыв возникает в МАРТЕ (баланс -1,6 млн руб) и достигает минимума в АПРЕЛЕ (-2,8 млн руб).\n"
    "Причина: в январе-феврале делается стартовая закупка на 12 млн руб (10%+10% годового объёма), а деньги от РТТ "
    "за январскую-февральскую вторичку приходят только в феврале-марте. К моменту, когда нужно платить Эмпилс за январскую первичку "
    "(платёж в марте с учётом отсрочки 45 дней), деньги от РТТ ещё не покрыли закупку.\n"
    "Решение: дистрибьютору нужен стартовый оборотный капитал минимум 3 млн руб (с запасом), либо увеличенная отсрочка от Эмпилс до 60 дней, "
    "либо предоплата от РТТ хотя бы по первым 3 месяцам."
))
ws.cell(row=r, column=1).font = Font(name="Calibri", size=11, bold=True, color="C00000")
ws.cell(row=r, column=1).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
ws.cell(row=r, column=1).fill = GAP_FILL
ws.row_dimensions[r].height = 30
ws.row_dimensions[r+1].height = 30
ws.row_dimensions[r+2].height = 30

# ==================== ЛИСТ 2: КАССОВЫЙ ПОТОК ====================
ws2 = wb.create_sheet("Кассовый поток")

# Заголовок
ws2.merge_cells("A1:I1")
ws2["A1"] = "Помесячный кассовый поток (с учётом отсрочек)"
ws2["A1"].font = TITLE_FONT
ws2.row_dimensions[1].height = 24

ws2.merge_cells("A2:I2")
ws2["A2"] = ("Логика: 45 дней отсрочки от Эмпилс → платёж в месяце N+1.5 ≈ округляем до N+2 (т.е. 100% первички месяца N оплачивается в месяце N+2). "
             "30 дней отсрочки от РТТ → деньги от РТТ в месяце N+1. Стартовый капитал = 0.")
ws2["A2"].font = SUBTITLE_FONT
ws2["A2"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
ws2.row_dimensions[2].height = 36

# Шапка
headers = [
    ("Месяц", 12),
    ("Первичка (закупка у Эмпилс)", 18),
    ("Вторичка (продажа в РТТ)", 18),
    ("Платёж Эмпилс (T+45д)", 18),
    ("Деньги от РТТ (T+30д)", 18),
    ("Чистый поток месяца", 16),
    ("Накопленный баланс", 16),
    ("Долг Эмпилс на конец месяца", 20),
    ("Статус", 14),
]
for col_idx, (title, width) in enumerate(headers, 1):
    cell = ws2.cell(row=4, column=col_idx, value=title)
    cell.fill = HEAD_FILL
    cell.font = HEAD_FONT
    cell.alignment = HEAD_ALIGN
    cell.border = BORDER
    ws2.column_dimensions[get_column_letter(col_idx)].width = width
ws2.row_dimensions[4].height = 42

# Данные
cash_flow = model['cash_flow']

# Расчёт долга Эмпилс на конец месяца
# Долг = первичка, по которой ещё не наступил срок платежа
# При отсрочке 45 дней и начале в январе:
# - Январь: первичка 6 млн, платёж в марте → на конец января долг = 6 млн
# - Февраль: +6 млн первички → долг = 12 млн
# - Март: +7 млн, -6 млн (платёж за январь) → долг = 13 млн
# - Апрель: +7 млн, -6 млн (платёж за февраль) → долг = 14 млн
# - Май: +4 млн, -7 млн (платёж за март) → долг = 11 млн
# и т.д.

debt_empils = 0
prev_primary = [0, 0]  # [N-2, N-1]
for i, cf in enumerate(cash_flow):
    # Платёж в этом месяце = первичка 2 месяца назад
    payment_this_month = prev_primary[0]
    
    # Обновляем долг: + первичка этого месяца, - платёж этого месяца
    debt_empils = debt_empils + cf['primary'] - payment_this_month
    
    cf['debt_empils'] = debt_empils
    
    prev_primary = [prev_primary[1], cf['primary']]

# Заполняем строки
for i, cf in enumerate(cash_flow):
    r = 5 + i
    ws2.cell(row=r, column=1, value=cf['month']).font = Font(name="Calibri", size=10, bold=True)
    ws2.cell(row=r, column=1).border = BORDER
    ws2.cell(row=r, column=1).alignment = CENTER_ALIGN
    
    # Первичка (жёлтый фон)
    cell = ws2.cell(row=r, column=2, value=cf['primary'])
    cell.font = BODY_FONT
    cell.alignment = NUM_ALIGN
    cell.number_format = '# ##0'
    cell.border = BORDER
    cell.fill = PRIMARY_HIGHLIGHT
    
    # Вторичка (голубой фон)
    cell = ws2.cell(row=r, column=3, value=cf['secondary'])
    cell.font = BODY_FONT
    cell.alignment = NUM_ALIGN
    cell.number_format = '# ##0'
    cell.border = BORDER
    cell.fill = SECONDARY_HIGHLIGHT
    
    # Платёж Эмпилс
    cell = ws2.cell(row=r, column=4, value=cf['pay_empils'])
    cell.font = BODY_FONT if cf['pay_empils'] == 0 else Font(name="Calibri", size=10, bold=True, color="C00000")
    cell.alignment = NUM_ALIGN
    cell.number_format = '# ##0'
    cell.border = BORDER
    
    # Деньги от РТТ
    cell = ws2.cell(row=r, column=5, value=cf['receive_rtt'])
    cell.font = BODY_FONT if cf['receive_rtt'] == 0 else Font(name="Calibri", size=10, bold=True, color="375623")
    cell.alignment = NUM_ALIGN
    cell.number_format = '# ##0'
    cell.border = BORDER
    
    # Чистый поток
    cell = ws2.cell(row=r, column=6, value=cf['net'])
    if cf['net'] < 0:
        cell.font = NEGATIVE_FONT
        cell.fill = GAP_FILL
    elif cf['net'] > 0:
        cell.font = POSITIVE_FONT
        cell.fill = POSITIVE_FILL
    else:
        cell.font = BODY_FONT
    cell.alignment = NUM_ALIGN
    cell.number_format = '# ##0;[Red]-# ##0'
    cell.border = BORDER
    
    # Накопленный баланс
    cell = ws2.cell(row=r, column=7, value=cf['cumulative'])
    if cf['cumulative'] < 0:
        cell.font = NEGATIVE_FONT
        cell.fill = GAP_FILL
    elif cf['cumulative'] > 0:
        cell.font = POSITIVE_FONT
        cell.fill = POSITIVE_FILL
    else:
        cell.font = BODY_FONT
    cell.alignment = NUM_ALIGN
    cell.number_format = '# ##0;[Red]-# ##0'
    cell.border = BORDER
    
    # Долг Эмпилс
    cell = ws2.cell(row=r, column=8, value=cf['debt_empils'])
    cell.font = BODY_FONT
    cell.alignment = NUM_ALIGN
    cell.number_format = '# ##0'
    cell.border = BORDER
    
    # Статус
    if cf['cumulative'] < 0:
        status = "⚠ РАЗРЫВ"
        color = "C00000"
    elif cf['cumulative'] == 0 and i == 0:
        status = "Старт"
        color = "595959"
    else:
        status = "✓ ОК"
        color = "375623"
    cell = ws2.cell(row=r, column=9, value=status)
    cell.font = Font(name="Calibri", size=10, bold=True, color=color)
    cell.alignment = CENTER_ALIGN
    cell.border = BORDER

# Итоговая строка
r_total = 5 + len(cash_flow)
ws2.cell(row=r_total, column=1, value="ИТОГО").font = Font(name="Calibri", size=10, bold=True)
ws2.cell(row=r_total, column=1).fill = TOTAL_FILL
ws2.cell(row=r_total, column=1).border = BORDER
ws2.cell(row=r_total, column=1).alignment = CENTER_ALIGN

totals = {
    'primary': sum(cf['primary'] for cf in cash_flow),
    'secondary': sum(cf['secondary'] for cf in cash_flow),
    'pay_empils': sum(cf['pay_empils'] for cf in cash_flow),
    'receive_rtt': sum(cf['receive_rtt'] for cf in cash_flow),
    'net': sum(cf['net'] for cf in cash_flow),
}
for col_idx, key in enumerate(['primary', 'secondary', 'pay_empils', 'receive_rtt', 'net'], 2):
    cell = ws2.cell(row=r_total, column=col_idx, value=totals[key])
    cell.font = Font(name="Calibri", size=10, bold=True)
    cell.fill = TOTAL_FILL
    cell.alignment = NUM_ALIGN
    cell.number_format = '# ##0;[Red]-# ##0'
    cell.border = BORDER

# Финальный накопленный баланс
cell = ws2.cell(row=r_total, column=7, value=cash_flow[-1]['cumulative'])
cell.font = Font(name="Calibri", size=10, bold=True, color="375623" if cash_flow[-1]['cumulative'] > 0 else "C00000")
cell.fill = TOTAL_FILL
cell.alignment = NUM_ALIGN
cell.number_format = '# ##0;[Red]-# ##0'
cell.border = BORDER

# Закрепить заголовок
ws2.freeze_panes = "B5"

# === График кассового потока ===
chart = LineChart()
chart.title = "Кассовый поток: поступления, платежи, накопленный баланс"
chart.style = 10
chart.y_axis.title = "Рубли"
chart.x_axis.title = "Месяц"
chart.height = 12
chart.width = 24

data = Reference(ws2, min_col=2, min_row=4, max_col=7, max_row=4 + len(cash_flow))
cats = Reference(ws2, min_col=1, min_row=5, max_row=4 + len(cash_flow))
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)

ws2.add_chart(chart, f"K4")

# ==================== ЛИСТ 3: ПЛАН ПЕРВИЧКИ И ВТОРИЧКИ ====================
ws3 = wb.create_sheet("План первички и вторички")

ws3.merge_cells("A1:F1")
ws3["A1"] = "Помесячный план первички и вторички с учётом сезонности ЛКМ"
ws3["A1"].font = TITLE_FONT
ws3.row_dimensions[1].height = 24

ws3.merge_cells("A2:F2")
ws3["A2"] = ("Сезонность ЛКМ (по данным открытых источников): пик спроса — июнь (+25% к весне), "
             "строительный сезон — апрель-сентябрь. Зимой (янв-мар) — низкий сезон. "
             "Январь-май: 50% годовой первички (30 млн) — стартовая загрузка склада. "
             "Июнь-дек: 30 млн по убывающей.")
ws3["A2"].font = SUBTITLE_FONT
ws3["A2"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
ws3.row_dimensions[2].height = 50

headers3 = [
    ("Месяц", 12),
    ("Первичка, руб", 16),
    ("Доля первички, %", 14),
    ("Вторичка, руб", 16),
    ("Доля вторички, %", 14),
    ("Коэфф. опереж. (первичка/вторичка)", 18),
]
for col_idx, (title, width) in enumerate(headers3, 1):
    cell = ws3.cell(row=4, column=col_idx, value=title)
    cell.fill = HEAD_FILL
    cell.font = HEAD_FONT
    cell.alignment = HEAD_ALIGN
    cell.border = BORDER
    ws3.column_dimensions[get_column_letter(col_idx)].width = width
ws3.row_dimensions[4].height = 42

for i, cf in enumerate(cash_flow):
    r = 5 + i
    prim = cf['primary']
    sec = cf['secondary']
    prim_pct = prim / 60_000_000 * 100
    sec_pct = sec / 60_000_000 * 100
    ratio = prim / sec if sec > 0 else 0
    
    ws3.cell(row=r, column=1, value=cf['month']).font = Font(name="Calibri", size=10, bold=True)
    ws3.cell(row=r, column=1).alignment = CENTER_ALIGN
    ws3.cell(row=r, column=1).border = BORDER
    
    cell = ws3.cell(row=r, column=2, value=prim)
    cell.fill = PRIMARY_HIGHLIGHT
    cell.font = BODY_FONT
    cell.alignment = NUM_ALIGN
    cell.number_format = '# ##0'
    cell.border = BORDER
    
    cell = ws3.cell(row=r, column=3, value=prim_pct/100)
    cell.font = BODY_FONT
    cell.alignment = NUM_ALIGN
    cell.number_format = '0.0%'
    cell.border = BORDER
    
    cell = ws3.cell(row=r, column=4, value=sec)
    cell.fill = SECONDARY_HIGHLIGHT
    cell.font = BODY_FONT
    cell.alignment = NUM_ALIGN
    cell.number_format = '# ##0'
    cell.border = BORDER
    
    cell = ws3.cell(row=r, column=5, value=sec_pct/100)
    cell.font = BODY_FONT
    cell.alignment = NUM_ALIGN
    cell.number_format = '0.0%'
    cell.border = BORDER
    
    cell = ws3.cell(row=r, column=6, value=ratio)
    cell.font = BODY_FONT
    cell.alignment = NUM_ALIGN
    cell.number_format = '0.00'
    cell.border = BORDER
    
    # Подсветка: если первичка > вторички (опережение), выделить зелёным
    if ratio > 1.0:
        for c in range(1, 7):
            ws3.cell(row=r, column=c).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

# Итоги
r_total = 5 + len(cash_flow)
ws3.cell(row=r_total, column=1, value="ИТОГО").font = Font(name="Calibri", size=10, bold=True)
ws3.cell(row=r_total, column=1).fill = TOTAL_FILL
ws3.cell(row=r_total, column=1).border = BORDER
ws3.cell(row=r_total, column=1).alignment = CENTER_ALIGN

cell = ws3.cell(row=r_total, column=2, value=sum(cf['primary'] for cf in cash_flow))
cell.font = Font(name="Calibri", size=10, bold=True)
cell.fill = TOTAL_FILL
cell.alignment = NUM_ALIGN
cell.number_format = '# ##0'
cell.border = BORDER

cell = ws3.cell(row=r_total, column=3, value=1.0)
cell.font = Font(name="Calibri", size=10, bold=True)
cell.fill = TOTAL_FILL
cell.alignment = NUM_ALIGN
cell.number_format = '0.0%'
cell.border = BORDER

cell = ws3.cell(row=r_total, column=4, value=sum(cf['secondary'] for cf in cash_flow))
cell.font = Font(name="Calibri", size=10, bold=True)
cell.fill = TOTAL_FILL
cell.alignment = NUM_ALIGN
cell.number_format = '# ##0'
cell.border = BORDER

cell = ws3.cell(row=r_total, column=5, value=sum(cf['secondary'] for cf in cash_flow)/60_000_000)
cell.font = Font(name="Calibri", size=10, bold=True)
cell.fill = TOTAL_FILL
cell.alignment = NUM_ALIGN
cell.number_format = '0.0%'
cell.border = BORDER

# ==================== ЛИСТ 4: АНАЛИЗ РАЗРЫВА ====================
ws4 = wb.create_sheet("Анализ разрыва")

ws4.merge_cells("A1:F1")
ws4["A1"] = "Анализ кассового разрыва и рекомендации"
ws4["A1"].font = TITLE_FONT
ws4.row_dimensions[1].height = 28

# Раздел 1: точка разрыва
ws4["A3"] = "1. ТОЧКА КАССОВОГО РАЗРЫВА"
ws4["A3"].font = Font(name="Calibri", size=12, bold=True, color="C00000")
ws4.merge_cells("A3:F3")

gap_data = [
    ("Параметр", "Значение"),
    ("Первый месяц с отрицательным балансом", "МАРТ"),
    ("Накопленный баланс в марте", "-1 600 000 руб"),
    ("Месяц с минимальным балансом", "АПРЕЛЬ"),
    ("Минимальный накопленный баланс", "-2 800 000 руб"),
    ("Глубина разрыва (продолжительность)", "2 месяца (март-апрель)"),
    ("Когда баланс восстанавливается", "ИЮЛЬ (накопл. +2,55 млн руб)"),
    ("Долг Эмпилс на конец апреля", "14 000 000 руб"),
]

for col_idx, (title, width) in enumerate([("Параметр", 45), ("Значение", 35)], 1):
    cell = ws4.cell(row=4, column=col_idx, value=title)
    cell.fill = HEAD_FILL
    cell.font = HEAD_FONT
    cell.alignment = HEAD_ALIGN
    cell.border = BORDER
    ws4.column_dimensions[get_column_letter(col_idx)].width = width

for i, (param, val) in enumerate(gap_data[1:], 5):
    ws4.cell(row=i, column=1, value=param).font = Font(name="Calibri", size=10, bold=True)
    ws4.cell(row=i, column=1).border = BORDER
    ws4.cell(row=i, column=1).alignment = Alignment(horizontal="left", vertical="center")
    
    cell = ws4.cell(row=i, column=2, value=val)
    if "разрыв" in param.lower() or "миним" in param.lower() or "отриц" in param.lower() or "МАРТ" in val or "АПРЕЛЬ" in val:
        cell.font = Font(name="Calibri", size=10, bold=True, color="C00000")
        cell.fill = GAP_FILL
    else:
        cell.font = Font(name="Calibri", size=10, bold=True, color="375623")
        cell.fill = POSITIVE_FILL
    cell.alignment = CENTER_ALIGN
    cell.border = BORDER

# Раздел 2: причины
r = 14
ws4.cell(row=r, column=1, value="2. ПРИЧИНЫ РАЗРЫВА").font = Font(name="Calibri", size=12, bold=True, color="1F4E79")
ws4.merge_cells(f"A{r}:F{r}")

reasons = [
    ("Стартовая загрузка склада в январе", "6 млн руб первички в январе — это 10% годового плана за один месяц. Дистрибьютор должен выкупить у Эмпилс 300 SKU, чтобы заполнить полки РТТ."),
    ("Опережение вторички первичкой", "В январе первичка 6 млн, а вторичка всего 2 млн (коэффициент опережения 3.0). В феврале — 6/2,4 = 2,5. Это нормальная логика дистрибьютора: сначала склад, потом продажи."),
    ("45-дневная отсрочка недостаточна для январской закупки", "Январская первичка (6 млн) должна быть оплачена в марте. К этому моменту деньги от РТТ за январскую вторичку (2 млн) и февральскую (2,4 млн) уже пришли (4,4 млн), но это покрывает только 73% долга Эмпилс."),
    ("Мартовская «контрактная вилка»", "В марте нужно одновременно платить за январь (6 млн) И продолжать закупать (первичка марта = 7 млн). Суммарная потребность в марте = 13 млн, при поступлениях от РТТ только 2,4 млн."),
    ("Минимум в апреле", "В апреле приходит платеж за февральскую первичку (6 млн), но деньги от РТТ за мартовскую вторичку (4,8 млн) ещё не покрывают его. Дополнительно нужна апрельская первичка 7 млн."),
]

for i, (cause, detail) in enumerate(reasons):
    rr = 15 + i
    ws4.cell(row=rr, column=1, value=cause).font = Font(name="Calibri", size=10, bold=True, color="C00000")
    ws4.cell(row=rr, column=1).border = BORDER
    ws4.cell(row=rr, column=1).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws4.merge_cells(f"B{rr}:F{rr}")
    ws4.cell(row=rr, column=2, value=detail).font = BODY_FONT
    ws4.cell(row=rr, column=2).border = BORDER
    ws4.cell(row=rr, column=2).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws4.row_dimensions[rr].height = 50

# Раздел 3: варианты решения
r = 22
ws4.cell(row=r, column=1, value="3. ВАРИАНТЫ РЕШЕНИЯ").font = Font(name="Calibri", size=12, bold=True, color="375623")
ws4.merge_cells(f"A{r}:F{r}")

solutions = [
    ("Вариант", "Описание", "Эффект на разрыв", "Сложность"),
    ("А. Стартовый капитал 3 млн руб", "Дистрибьютор вносит собственные средства на старте (январь), чтобы покрыть разрыв марта-апреля", "Полное закрытие разрыва (баланс всегда ≥ 0)", "Низкая — нужно найти 3 млн"),
    ("Б. Увеличить отсрочку до 60 дней", "Просить Эмпилс увеличить отсрочку с 45 до 60 дней (доп. 15 дней)", "Сдвиг платежей на 1 месяц вперёд: разрыв начнётся в апреле (-1,2 млн), но будет мягче", "Средняя — нужны переговоры"),
    ("В. Снизить январскую первичку до 4 млн", "Уменьшить стартовую загрузку: 4 млн в январе вместо 6 млн", "Разрыв начнётся в марте с -0,8 млн, минимум -2,0 млн в апреле. Меньше глубина.", "Низкая — но риск дефицита на полках"),
    ("Г. Предоплата от РТТ первые 3 мес", "Договориться с РТТ о предоплате в январе-марте (вместо постоплаты 30 дн)", "Деньги приходят сразу: январь +2 млн, февраль +2,4 млн, март +4,8 млн. Разрыв полностью закрывается.", "Высокая — РТТ редко соглашаются"),
    ("Д. Комбинация: кап. 1,5 млн + отсрочка 60 дн", "Минимальный капитал 1,5 млн и отсрочка 60 дней", "Полное закрытие разрыва с запасом", "Средняя"),
]

for col_idx, (title, width) in enumerate([("Вариант", 30), ("Описание", 50), ("Эффект на разрыв", 40), ("Сложность", 25)], 1):
    cell = ws4.cell(row=23, column=col_idx, value=title)
    cell.fill = HEAD_FILL
    cell.font = HEAD_FONT
    cell.alignment = HEAD_ALIGN
    cell.border = BORDER
    ws4.column_dimensions[get_column_letter(col_idx)].width = width

for i, (var, desc, eff, diff) in enumerate(solutions[1:], 24):
    ws4.cell(row=i, column=1, value=var).font = Font(name="Calibri", size=10, bold=True)
    ws4.cell(row=i, column=1).border = BORDER
    ws4.cell(row=i, column=1).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    
    ws4.cell(row=i, column=2, value=desc).font = BODY_FONT
    ws4.cell(row=i, column=2).border = BORDER
    ws4.cell(row=i, column=2).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    
    ws4.cell(row=i, column=3, value=eff).font = BODY_FONT
    ws4.cell(row=i, column=3).border = BORDER
    ws4.cell(row=i, column=3).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    
    ws4.cell(row=i, column=4, value=diff).font = BODY_FONT
    ws4.cell(row=i, column=4).border = BORDER
    ws4.cell(row=i, column=4).alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
    
    ws4.row_dimensions[i].height = 50

# Раздел 4: вывод
r = 30
ws4.cell(row=r, column=1, value="4. ВЫВОД ПО ЛОГИКЕ ДИСТРИБЬЮТОРСКОГО БИЗНЕСА В ЛКМ").font = Font(name="Calibri", size=12, bold=True, color="1F4E79")
ws4.merge_cells(f"A{r}:F{r}")

conclusion = (
    "Логика подтверждена. Дистрибьюторский бизнес в ЛКМ имеет ярко выраженный кассовый разрыв на старте по следующей причине:\n\n"
    "1. СЕЗОННОСТЬ: пик спроса в июне (+25% к весне, по данным открытых источников). Чтобы встретить пик, нужно в январе-мае сделать 50% годовой первички (30 млн руб).\n"
    "2. ОТСРОЧКА 45 ДНЕЙ НЕДОСТАТОЧНА: при постоплате 45 дней январская первичка (6 млн) оплачивается в марте — раньше, чем дистрибьютор успевает распродать её через РТТ и собрать деньги (с учётом 30-дневной отсрочки от РТТ).\n"
    "3. ОПЕРЕЖЕНИЕ ПЕРВИЧКИ ВТОРИЧКОЙ: в январе коэффициент опережения = 3,0 (6/2), в феврале — 2,5 (6/2,4). Это нормально для фазы запуска, но требует оборотного капитала.\n"
    "4. ВОССТАНОВЛЕНИЕ: с июня график выходит в плюс (накопленный баланс +2,55 млн в июле). К концу года накопленный излишек +2,36 млн — это резерв на следующий сезон.\n\n"
    "РЕКОМЕНДАЦИЯ ЭМПИЛС: при переговорах с Олимпом предложить гибридный вариант (Д) — стартовый капитал 1,5 млн руб со стороны Олимпа + увеличенная отсрочка 60 дней от Эмпилс на первые 6 месяцев сотрудничества. "
    "Это снимет кассовый разрыв и создаст доверие к партнёрству."
)
ws4.merge_cells(f"A{r+1}:F{r+8}")
ws4.cell(row=r+1, column=1, value=conclusion).font = BODY_FONT
ws4.cell(row=r+1, column=1).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
ws4.cell(row=r+1, column=1).fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
ws4.row_dimensions[r+1].height = 25
for rr in range(r+2, r+9):
    ws4.row_dimensions[rr].height = 25

# ==================== ЛИСТ 5: АССОРТИМЕНТ И СЕБЕСТОИМОСТЬ ====================
ws5 = wb.create_sheet("Ассортимент и себестоимость")

ws5.merge_cells("A1:G1")
ws5["A1"] = "Ассортиментная матрица 300 SKU: распределение по ценовым сегментам"
ws5["A1"].font = TITLE_FONT
ws5.row_dimensions[1].height = 24

ws5.merge_cells("A2:G2")
ws5["A2"] = ("Источник: файл «Ассортимент_Олимп_300_SKU_финал_v4.xlsx», лист «300 SKU с ценами». "
             "Цены — отгрузочные из прайс-листа Эмпилс от 16.10.2026. Эти цены = себестоимость закупки дистрибьютора.")
ws5["A2"].font = SUBTITLE_FONT
ws5["A2"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
ws5.row_dimensions[2].height = 36

# Сегментация по цене
segments = [
    ("До 100 руб", 0, 100),
    ("100-300 руб", 100, 300),
    ("300-500 руб", 300, 500),
    ("500-1000 руб", 500, 1000),
    ("1000-2000 руб", 1000, 2000),
    ("2000-5000 руб", 2000, 5000),
    ("Более 5000 руб", 5000, 999999),
]

# Шапка
headers5 = [
    ("Ценовой сегмент", 22),
    ("Кол-во SKU", 14),
    ("Доля, %", 12),
    ("Сумма цен, руб", 18),
    ("Доля в сумме, %", 14),
    ("Средняя цена, руб", 18),
]
for col_idx, (title, width) in enumerate(headers5, 1):
    cell = ws5.cell(row=4, column=col_idx, value=title)
    cell.fill = HEAD_FILL
    cell.font = HEAD_FONT
    cell.alignment = HEAD_ALIGN
    cell.border = BORDER
    ws5.column_dimensions[get_column_letter(col_idx)].width = width
ws5.row_dimensions[4].height = 32

# Считаем распределение
all_skus = []
for r in range(5, ws_src.max_row + 1):
    price = ws_src.cell(row=r, column=10).value
    if price and isinstance(price, (int, float)):
        all_skus.append(price)

total_count = len(all_skus)
total_sum = sum(all_skus)

for i, (seg_name, lo, hi) in enumerate(segments):
    r = 5 + i
    in_seg = [p for p in all_skus if lo <= p < hi]
    count = len(in_seg)
    seg_sum = sum(in_seg)
    avg = seg_sum / count if count > 0 else 0
    
    ws5.cell(row=r, column=1, value=seg_name).font = Font(name="Calibri", size=10, bold=True)
    ws5.cell(row=r, column=1).border = BORDER
    ws5.cell(row=r, column=1).alignment = Alignment(horizontal="left", vertical="center")
    
    ws5.cell(row=r, column=2, value=count).font = BODY_FONT
    ws5.cell(row=r, column=2).border = BORDER
    ws5.cell(row=r, column=2).alignment = NUM_ALIGN
    
    ws5.cell(row=r, column=3, value=count/total_count).font = BODY_FONT
    ws5.cell(row=r, column=3).border = BORDER
    ws5.cell(row=r, column=3).alignment = NUM_ALIGN
    ws5.cell(row=r, column=3).number_format = '0.0%'
    
    ws5.cell(row=r, column=4, value=seg_sum).font = BODY_FONT
    ws5.cell(row=r, column=4).border = BORDER
    ws5.cell(row=r, column=4).alignment = NUM_ALIGN
    ws5.cell(row=r, column=4).number_format = '# ##0'
    
    ws5.cell(row=r, column=5, value=seg_sum/total_sum).font = BODY_FONT
    ws5.cell(row=r, column=5).border = BORDER
    ws5.cell(row=r, column=5).alignment = NUM_ALIGN
    ws5.cell(row=r, column=5).number_format = '0.0%'
    
    ws5.cell(row=r, column=6, value=avg).font = BODY_FONT
    ws5.cell(row=r, column=6).border = BORDER
    ws5.cell(row=r, column=6).alignment = NUM_ALIGN
    ws5.cell(row=r, column=6).number_format = '# ##0.0'

# Итого
r_tot = 5 + len(segments)
ws5.cell(row=r_tot, column=1, value="ИТОГО 300 SKU").font = Font(name="Calibri", size=10, bold=True)
ws5.cell(row=r_tot, column=1).fill = TOTAL_FILL
ws5.cell(row=r_tot, column=1).border = BORDER

ws5.cell(row=r_tot, column=2, value=total_count).font = Font(name="Calibri", size=10, bold=True)
ws5.cell(row=r_tot, column=2).fill = TOTAL_FILL
ws5.cell(row=r_tot, column=2).border = BORDER
ws5.cell(row=r_tot, column=2).alignment = NUM_ALIGN

ws5.cell(row=r_tot, column=3, value=1.0).font = Font(name="Calibri", size=10, bold=True)
ws5.cell(row=r_tot, column=3).fill = TOTAL_FILL
ws5.cell(row=r_tot, column=3).border = BORDER
ws5.cell(row=r_tot, column=3).alignment = NUM_ALIGN
ws5.cell(row=r_tot, column=3).number_format = '0.0%'

ws5.cell(row=r_tot, column=4, value=total_sum).font = Font(name="Calibri", size=10, bold=True)
ws5.cell(row=r_tot, column=4).fill = TOTAL_FILL
ws5.cell(row=r_tot, column=4).border = BORDER
ws5.cell(row=r_tot, column=4).alignment = NUM_ALIGN
ws5.cell(row=r_tot, column=4).number_format = '# ##0'

ws5.cell(row=r_tot, column=5, value=1.0).font = Font(name="Calibri", size=10, bold=True)
ws5.cell(row=r_tot, column=5).fill = TOTAL_FILL
ws5.cell(row=r_tot, column=5).border = BORDER
ws5.cell(row=r_tot, column=5).alignment = NUM_ALIGN
ws5.cell(row=r_tot, column=5).number_format = '0.0%'

ws5.cell(row=r_tot, column=6, value=total_sum/total_count).font = Font(name="Calibri", size=10, bold=True)
ws5.cell(row=r_tot, column=6).fill = TOTAL_FILL
ws5.cell(row=r_tot, column=6).border = BORDER
ws5.cell(row=r_tot, column=6).alignment = NUM_ALIGN
ws5.cell(row=r_tot, column=6).number_format = '# ##0.0'

# Ключевой вывод по ассортименту
r = r_tot + 3
ws5.cell(row=r, column=1, value="КЛЮЧЕВЫЕ ВЫВОДЫ ПО АССОРТИМЕНТУ").font = Font(name="Calibri", size=12, bold=True, color="1F4E79")
ws5.merge_cells(f"A{r}:F{r}")

assortment_notes = [
    f"Всего в матрице 300 SKU с общей суммой отгрузочных цен {total_sum:,.0f} руб (средняя цена 1 SKU = {total_sum/total_count:,.0f} руб).",
    "Это значит, что для закупки всей матрицы по 1 шт. каждого SKU дистрибьютору нужно ~2,3 млн руб (без учёта фасовок).",
    "В реальности, чтобы заполнить полки 300 SKU с минимальным запасом 5 шт. на SKU, нужно ~11,5 млн руб — что примерно соответствует нашей январской первичке в 6 млн руб (т.е. стартовая закупка = ~50% от полной заливки матрицы).",
    "Дистрибьютор на старте НЕ закупает все 300 SKU сразу — он формирует топ-100 SKU для запуска и постепенно расширяет матрицу. Это снижает кассовый разрыв, но требует более точного отбора ходовых позиций (по ABC-анализу).",
    "Структура цен: большая часть SKU (по количеству) находится в сегменте 100-1000 руб — это типичный ассортимент масс-маркета.",
]

for i, note in enumerate(assortment_notes):
    rr = r + 1 + i
    ws5.merge_cells(f"A{rr}:F{rr}")
    ws5.cell(row=rr, column=1, value=f"• {note}").font = BODY_FONT
    ws5.cell(row=rr, column=1).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws5.row_dimensions[rr].height = 36

# === Сохранить ===
wb.save(OUT_PATH)
print(f"✓ Файл сохранён: {OUT_PATH}")
print(f"  Размер: {os.path.getsize(OUT_PATH):,} байт")
print(f"  Листы: {wb.sheetnames}")
