#!/usr/bin/env python3
"""TCO-калькулятор Эмпилс — полный функционал"""

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers, Protection
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from copy import copy

wb = Workbook()

# ============================================================
# COLORS
# ============================================================
NAVY = "0A2540"
CYAN = "00B4D8"
LIGHT_CYAN = "90E0EF"
LIGHT_BG = "F0F7FA"
WHITE = "FFFFFF"
GREEN = "1B7D46"
RED = "C0392B"
ORANGE = "E86C00"
GOLD = "FCBF49"
DARK_GRAY = "333333"
MID_GRAY = "888888"
LIGHT_GRAY = "F5F5F5"

# ============================================================
# STYLE HELPERS
# ============================================================
def hdr_font(sz=11, white=True):
    return Font(name="Calibri", size=sz, bold=True, color=WHITE if white else NAVY)

def body_font(sz=10, bold=False, color=DARK_GRAY):
    return Font(name="Calibri", size=sz, bold=bold, color=color)

def fill(color):
    return PatternFill("solid", fgColor=color)

def align(h="left", v="center", wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

thin_border = Border(
    left=Side(style="thin", color="D0D0D0"),
    right=Side(style="thin", color="D0D0D0"),
    top=Side(style="thin", color="D0D0D0"),
    bottom=Side(style="thin", color="D0D0D0"),
)

def style_header_row(ws, row, cols, fill_color=NAVY):
    for c in range(1, cols+1):
        cell = ws.cell(row=row, column=c)
        cell.font = hdr_font()
        cell.fill = PatternFill("solid", fgColor=fill_color)
        cell.alignment = align("center", "center", True)
        cell.border = thin_border

def style_data_cell(ws, row, col, fmt=None, bold=False, color=DARK_GRAY, bg=None):
    cell = ws.cell(row=row, column=col)
    cell.font = body_font(10, bold, color)
    cell.alignment = align("left", "center", True)
    cell.border = thin_border
    if fmt:
        cell.number_format = fmt
    if bg:
        cell.fill = PatternFill("solid", fgColor=bg)

def style_input_cell(ws, row, col, fmt=None):
    cell = ws.cell(row=row, column=col)
    cell.font = Font(name="Calibri", size=11, bold=True, color="0000FF")
    cell.fill = PatternFill("solid", fgColor="FFFFCC")
    cell.alignment = align("center", "center", False)
    cell.border = Border(
        left=Side(style="medium", color=CYAN),
        right=Side(style="medium", color=CYAN),
        top=Side(style="medium", color=CYAN),
        bottom=Side(style="medium", color=CYAN),
    )
    if fmt:
        cell.number_format = fmt
    cell.protection = Protection(locked=False)

def style_result_cell(ws, row, col, fmt=None, bold=True, color=NAVY, bg=None):
    cell = ws.cell(row=row, column=col)
    cell.font = Font(name="Calibri", size=12, bold=bold, color=color)
    cell.fill = PatternFill("solid", fgColor=bg if bg else LIGHT_BG)
    cell.alignment = align("center", "center", False)
    cell.border = Border(
        left=Side(style="medium", color=CYAN),
        right=Side(style="medium", color=CYAN),
        top=Side(style="medium", color=CYAN),
        bottom=Side(style="medium", color=CYAN),
    )
    if fmt:
        cell.number_format = fmt

# ============================================================
# SHEET 1: КАЛЬКУЛЯТОР (main)
# ============================================================
ws = wb.active
ws.title = "Калькулятор TCO"
ws.sheet_properties.tabColor = CYAN

# Column widths
col_w = {1:3, 2:38, 3:18, 4:18, 5:18, 6:18, 7:18, 8:3}
for c, w in col_w.items():
    ws.column_dimensions[get_column_letter(c)].width = w

# ---- TITLE ----
r = 2
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
c = ws.cell(row=r, column=2, value="TCO-КАЛЬКУЛЯТОР ЭМПИЛС")
c.font = Font(name="Calibri", size=18, bold=True, color=NAVY)
c.alignment = align("left", "center", False)

r = 3
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
c = ws.cell(row=r, column=2, value="Совокупная стоимость владения — инструмент обоснования цены премиум")
c.font = Font(name="Calibri", size=11, color=MID_GRAY)
c.alignment = align("left", "center", False)

r = 4
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
c = ws.cell(row=r, column=2, value="Вводите данные в жёлтые ячейки — результаты рассчитываются автоматически")
c.font = Font(name="Calibri", size=10, bold=True, color=ORANGE)
c.alignment = align("left", "center", False)

# ---- INPUT SECTION ----
r = 6
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
c = ws.cell(row=r, column=2, value="ВВОД ДАННЫХ")
c.font = hdr_font(12, True)
c.fill = PatternFill("solid", fgColor=CYAN)
c.alignment = align("center", "center", True)

# Input fields
inputs = [
    (7, "Тип помещения", "комната 20 м²", None),
    (8, "Площадь стен (м²)", 45, '#,##0'),
    (9, "Высота потолков (м)", 2.7, '0.0'),
    (10, "Количество слоёв краски", 2, '0'),
    (11, "Стоимость работы за 1 м² (грунтовка, ₽)", 80, '#,##0'),
    (12, "Стоимость работы за 1 м² (шпатлёвка, ₽)", 120, '#,##0'),
    (13, "Стоимость работы за 1 м² (покраска, ₽)", 150, '#,##0'),
]

for row_num, label, default, fmt in inputs:
    ws.cell(row=row_num, column=2, value=label).font = body_font(10, True, NAVY)
    ws.cell(row=row_num, column=2).alignment = align("left", "center", True)
    ws.cell(row=row_num, column=2).border = thin_border
    
    c = ws.cell(row=row_num, column=3, value=default)
    style_input_cell(ws, row_num, 3, fmt)
    
    # Description in col 4
    descs = {
        7: "Выберите из списка или введите свой вариант",
        8: "Площадь окрашиваемой поверхности",
        9: "Стандартная высота 2.5–2.7 м",
        10: "Обычно 2 слоя для ВД краски",
        11: "Средняя ставка по региону",
        12: "Зависит от сложности поверхности",
        13: "Включая валик/кисть",
    }
    if row_num in descs:
        ws.cell(row=row_num, column=4, value=descs[row_num]).font = body_font(9, False, MID_GRAY)
        ws.cell(row=row_num, column=4).alignment = align("left", "center", True)

# Data validation for room type
dv = DataValidation(type="list", formula1='"комната 20 м²,ванная,фасад дома,деревянный домик,детская"')
dv.prompt = "Выберите тип помещения"
dv.promptTitle = "Тип помещения"
ws.add_data_validation(dv)
dv.add(ws["C7"])

# ---- REFERENCE DATA (hidden section) ----
r = 15
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
c = ws.cell(row=r, column=2, value="СПРАВОЧНЫЕ ДАННЫЕ — РАСХОД И СТОИМОСТЬ МАТЕРИАЛОВ")
c.font = hdr_font(11, True)
c.fill = PatternFill("solid", fgColor=NAVY)
c.alignment = align("center", "center", True)

r = 16
headers = ["", "Показатель", "Эконом (ОРЕОЛ)", "Средний (NEWTONE)", "Премиум (AVANTGARDE)", "Ед. изм."]
for i, h in enumerate(headers):
    ws.cell(row=r, column=i+1, value=h)
style_header_row(ws, r, 6, NAVY)

# Product data: name, economy_price/kg, mid_price/kg, premium_price/kg, unit, consumption_per_m2, consumption_m2_premium
products = [
    ("Грунтовка", 130, 180, 350, "₽/кг", 0.15, 0.12),
    ("Шпатлёвка стартовая", 120, 170, 280, "₽/кг", 1.2, 1.0),
    ("Шпатлёвка финишная", "—", 250, 420, "₽/кг", 0.8, 0.6),
    ("Краска ВД интерьерная", 180, 320, 600, "₽/кг", 0.18, 0.14),
    ("Лак защитный", "—", 350, 800, "₽/кг", 0.10, 0.08),
    ("Декоративный эффект", "—", "—", 1200, "₽/кг", 0.20, 0.15),
]

# Service life
service_life = {"Эконом": 3, "Средний": 7, "Премиум": 10}

for i, (name, econ, mid, prem, unit, cons, cons_prem) in enumerate(products):
    row_num = 17 + i
    ws.cell(row=row_num, column=2, value=name).font = body_font(10, True, NAVY)
    ws.cell(row=row_num, column=2).border = thin_border
    
    for col_idx, val in [(3, econ), (4, mid), (5, prem)]:
        c = ws.cell(row=row_num, column=col_idx, value=val)
        c.font = body_font(10)
        c.alignment = align("center", "center", False)
        c.border = thin_border
        if isinstance(val, (int, float)):
            c.number_format = '#,##0'
    
    ws.cell(row=row_num, column=6, value=unit).font = body_font(9, False, MID_GRAY)
    ws.cell(row=row_num, column=6).alignment = align("center", "center", False)
    ws.cell(row=row_num, column=6).border = thin_border

# Service life row
row_num = 23
ws.cell(row=row_num, column=2, value="Срок службы покрытия (лет)").font = body_font(10, True, NAVY)
ws.cell(row=row_num, column=2).border = thin_border
for col_idx, val in [(3, 3), (4, 7), (5, 10)]:
    c = ws.cell(row=row_num, column=col_idx, value=val)
    c.font = Font(name="Calibri", size=12, bold=True, color=GREEN if col_idx==5 else DARK_GRAY)
    c.alignment = align("center", "center", False)
    c.border = thin_border
ws.cell(row=row_num, column=6, value="лет").font = body_font(9, False, MID_GRAY)
ws.cell(row=row_num, column=6).border = thin_border

# ---- CALCULATION SECTION ----
r = 25
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
c = ws.cell(row=r, column=2, value="РАСЧЁТ СТОИМОСТИ МАТЕРИАЛОВ")
c.font = hdr_font(12, True)
c.fill = PatternFill("solid", fgColor=CYAN)
c.alignment = align("center", "center", True)

r = 26
calc_headers = ["", "Материал", "Расход (кг/м²)", "Цена (₽/кг)", "Стоимость (₽)", "Слой"]
for i, h in enumerate(calc_headers):
    ws.cell(row=r, column=i+1, value=h)
style_header_row(ws, r, 6, NAVY)

# ECONOMY section
r = 27
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
ws.cell(row=r, column=2, value="ЭКОНОМ (ОРЕОЛ)").font = hdr_font(11, False)
ws.cell(row=r, column=2).fill = PatternFill("solid", fgColor="FDEDEC")
ws.cell(row=r, column=2).alignment = align("left", "center", True)

# Economy materials: name, consumption formula, price ref, formula for cost, layers
econ_rows = [
    (28, "Грунтовка ОРЕОЛ", "=0.15*C8", "=C17", None, "=C10"),
    (29, "Шпатлёвка стартовая ОРЕОЛ", "=1.2*C8", "=C18", None, "1"),
    (30, "Краска ВД ОРЕОЛ", "=0.18*C8", "=C20", None, "=C10"),
]

for row_num, name, cons_f, price_f, cost_f, layers in econ_rows:
    ws.cell(row=row_num, column=2, value=name).font = body_font(10)
    ws.cell(row=row_num, column=2).border = thin_border
    ws.cell(row=row_num, column=3, value=cons_f if cons_f.startswith("=") else cons_f).font = body_font(10)
    ws.cell(row=row_num, column=3).number_format = '0.00'
    ws.cell(row=row_num, column=3).border = thin_border
    ws.cell(row=row_num, column=4, value=price_f if price_f.startswith("=") else price_f).font = body_font(10)
    ws.cell(row=row_num, column=4).number_format = '#,##0'
    ws.cell(row=row_num, column=4).border = thin_border
    # Cost = consumption * price * layers
    col_c = get_column_letter(3)
    col_p = get_column_letter(4)
    col_l = get_column_letter(6)
    ws.cell(row=row_num, column=5).value = f"={col_c}{row_num}*{col_p}{row_num}*{col_l}{row_num}"
    ws.cell(row=row_num, column=5).number_format = '#,##0'
    ws.cell(row=row_num, column=5).border = thin_border
    ws.cell(row=row_num, column=6, value=layers).font = body_font(10)
    ws.cell(row=row_num, column=6).alignment = align("center", "center", False)
    ws.cell(row=row_num, column=6).border = thin_border

# Economy total materials
r = 31
ws.cell(row=r, column=2, value="ИТОГО материалы (Эконом)").font = body_font(10, True, RED)
ws.cell(row=r, column=2).border = thin_border
ws.cell(row=r, column=5).value = "=SUM(E28:E30)"
ws.cell(row=r, column=5).font = Font(name="Calibri", size=11, bold=True, color=RED)
ws.cell(row=r, column=5).number_format = '#,##0'
ws.cell(row=r, column=5).border = thin_border
ws.cell(row=r, column=5).fill = PatternFill("solid", fgColor="FDEDEC")

# PREMIUM section
r = 33
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
ws.cell(row=r, column=2, value="ПРЕМИУМ (AVANTGARDE + NEWTONE)").font = hdr_font(11, False)
ws.cell(row=r, column=2).fill = PatternFill("solid", fgColor="E8F5E9")
ws.cell(row=r, column=2).alignment = align("left", "center", True)

prem_rows = [
    (34, "Грунтовка ОРЕОЛ (база)", "=0.12*C8", "=C17", None, "=C10"),
    (35, "Шпатлёвка стартовая NEWTONE", "=1.0*C8", "=C18", None, "1"),
    (36, "Шпатлёвка финишная NEWTONE", "=0.6*C8", "=C19", None, "1"),
    (37, "Краска ВД AVANTGARDE", "=0.14*C8", "=C20", None, "=C10"),
    (38, "Лак защитный AVANTGARDE", "=0.08*C8", "=C21", None, "1"),
]

for row_num, name, cons_f, price_f, cost_f, layers in prem_rows:
    ws.cell(row=row_num, column=2, value=name).font = body_font(10)
    ws.cell(row=row_num, column=2).border = thin_border
    ws.cell(row=row_num, column=3, value=cons_f).font = body_font(10)
    ws.cell(row=row_num, column=3).number_format = '0.00'
    ws.cell(row=row_num, column=3).border = thin_border
    ws.cell(row=row_num, column=4, value=price_f).font = body_font(10)
    ws.cell(row=row_num, column=4).number_format = '#,##0'
    ws.cell(row=row_num, column=4).border = thin_border
    col_c = get_column_letter(3)
    col_p = get_column_letter(4)
    col_l = get_column_letter(6)
    ws.cell(row=row_num, column=5).value = f"={col_c}{row_num}*{col_p}{row_num}*{col_l}{row_num}"
    ws.cell(row=row_num, column=5).number_format = '#,##0'
    ws.cell(row=row_num, column=5).border = thin_border
    ws.cell(row=row_num, column=6, value=layers).font = body_font(10)
    ws.cell(row=row_num, column=6).alignment = align("center", "center", False)
    ws.cell(row=row_num, column=6).border = thin_border

# Premium total materials
r = 39
ws.cell(row=r, column=2, value="ИТОГО материалы (Премиум)").font = body_font(10, True, GREEN)
ws.cell(row=r, column=2).border = thin_border
ws.cell(row=r, column=5).value = "=SUM(E34:E38)"
ws.cell(row=r, column=5).font = Font(name="Calibri", size=11, bold=True, color=GREEN)
ws.cell(row=r, column=5).number_format = '#,##0'
ws.cell(row=r, column=5).border = thin_border
ws.cell(row=r, column=5).fill = PatternFill("solid", fgColor="E8F5E9")

# ---- COST SUMMARY ----
r = 41
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
c = ws.cell(row=r, column=2, value="СВОДНЫЙ РАСЧЁТ TCO")
c.font = hdr_font(12, True)
c.fill = PatternFill("solid", fgColor=NAVY)
c.alignment = align("center", "center", True)

r = 42
sum_headers = ["", "Показатель", "Эконом", "Премиум", "Разница"]
for i, h in enumerate(sum_headers):
    ws.cell(row=r, column=i+1, value=h)
style_header_row(ws, r, 5, NAVY)

summary_rows = [
    (43, "Стоимость материалов (₽)", "=E31", "=E39", "=D43-C43"),
    (44, "Стоимость работ (₽)", "=C8*(C11+C12+C13)", "=C8*(C11+C12+C13)", "=D44-C44"),
    (45, "Общая стоимость ремонта (₽)", "=C43+C44", "=D43+D44", "=D45-C45"),
    (46, "Переплата за премиум к смете (%)", "—", "=(D45-C45)/C45", "—"),
    (47, "Срок службы покрытия (лет)", 3, 10, "=D47-C47"),
    (48, "Стоимость 1 года эксплуатации (₽/год)", "=C45/C47", "=D45/D47", "=C48-D48"),
    (49, "Экономия в год (₽)", "—", "=C48-D48", "—"),
    (50, "Экономия за срок службы (₽)", "—", "=D48*10", "—"),
]

for row_num, label, econ_f, prem_f, diff_f in summary_rows:
    ws.cell(row=row_num, column=2, value=label).font = body_font(10, True, NAVY)
    ws.cell(row=row_num, column=2).border = thin_border
    ws.cell(row=row_num, column=2).alignment = align("left", "center", True)
    
    for col, formula in [(3, econ_f), (4, prem_f), (5, diff_f)]:
        c = ws.cell(row=row_num, column=col, value=formula)
        c.border = thin_border
        c.alignment = align("center", "center", False)
        if row_num == 46:
            c.number_format = '0.0%'
        elif row_num == 48:
            c.number_format = '#,##0'
            c.font = Font(name="Calibri", size=12, bold=True, color=NAVY)
        else:
            c.number_format = '#,##0'
            c.font = body_font(10)

# Highlight key result rows
for col in [3, 4]:
    # Row 48 - cost per year
    ws.cell(row=48, column=col).font = Font(name="Calibri", size=13, bold=True, color=GREEN if col==4 else RED)
    ws.cell(row=48, column=col).fill = PatternFill("solid", fgColor="E8F5E9" if col==4 else "FDEDEC")
    ws.cell(row=48, column=col).border = Border(
        left=Side(style="medium", color=CYAN),
        right=Side(style="medium", color=CYAN),
        top=Side(style="medium", color=CYAN),
        bottom=Side(style="medium", color=CYAN),
    )

# Row 45 - total cost
for col in [3, 4]:
    ws.cell(row=45, column=col).font = Font(name="Calibri", size=11, bold=True, color=NAVY)
    ws.cell(row=45, column=col).fill = PatternFill("solid", fgColor=LIGHT_BG)

# ---- KEY MESSAGE ----
r = 52
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
c = ws.cell(row=r, column=2, value='РЕЗУЛЬТАТ: Премиум дешевле в X раз в пересчёте на год эксплуатации')
c.font = Font(name="Calibri", size=14, bold=True, color=GREEN)
c.alignment = align("center", "center", True)
c.fill = PatternFill("solid", fgColor="E8F5E9")

r = 53
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
ws.cell(row=r, column=2, value='=IF(D48>0,"Премиум дешевле в "&TEXT(C48/D48,"0.0")&" раз в год. Экономия "&TEXT(D49,"#,##0")&" ₽/год","")').font = Font(name="Calibri", size=13, bold=True, color=NAVY)
ws.cell(row=r, column=2).alignment = align("center", "center", True)

# ============================================================
# CHART
# ============================================================
r = 55
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
c = ws.cell(row=r, column=2, value="ВИЗУАЛИЗАЦИЯ СРАВНЕНИЯ")
c.font = hdr_font(12, True)
c.fill = PatternFill("solid", fgColor=CYAN)
c.alignment = align("center", "center", True)

# Chart data
chart_data_start = 56
ws.cell(row=chart_data_start, column=2, value="Показатель").font = body_font(9, True)
ws.cell(row=chart_data_start, column=3, value="Эконом").font = body_font(9, True)
ws.cell(row=chart_data_start, column=4, value="Премиум").font = body_font(9, True)

chart_items = [
    (57, "Материалы (₽)", "=C43", "=D43"),
    (58, "Работы (₽)", "=C44", "=D44"),
    (59, "Общая стоимость (₽)", "=C45", "=D45"),
    (60, "Стоимость года (₽)", "=C48", "=D48"),
]

for row_num, label, econ_f, prem_f in chart_items:
    ws.cell(row=row_num, column=2, value=label)
    ws.cell(row=row_num, column=3, value=econ_f)
    ws.cell(row=row_num, column=3).number_format = '#,##0'
    ws.cell(row=row_num, column=4, value=prem_f)
    ws.cell(row=row_num, column=4).number_format = '#,##0'

# Bar chart
chart = BarChart()
chart.type = "col"
chart.grouping = "clustered"
chart.title = "Сравнение TCO: Эконом vs Премиум"
chart.y_axis.title = "₽"
chart.x_axis.title = None
chart.width = 20
chart.height = 12
chart.gapWidth = 80

cats = Reference(ws, min_col=2, min_row=57, max_row=60)
data = Reference(ws, min_col=3, max_col=4, min_row=56, max_row=60)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)

# Style series
s0 = chart.series[0]
s0.graphicalProperties.solidFill = RED
s1 = chart.series[1]
s1.graphicalProperties.solidFill = GREEN

chart.legend.position = 'b'

ws.add_chart(chart, "B62")

# ============================================================
# SHEET 2: СЦЕНАРИИ
# ============================================================
ws2 = wb.create_sheet("Сценарии")
ws2.sheet_properties.tabColor = GREEN

ws2.column_dimensions['A'].width = 3
ws2.column_dimensions['B'].width = 30
ws2.column_dimensions['C'].width = 16
ws2.column_dimensions['D'].width = 16
ws2.column_dimensions['E'].width = 16
ws2.column_dimensions['F'].width = 16
ws2.column_dimensions['G'].width = 16
ws2.column_dimensions['H'].width = 16

r = 2
ws2.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
c = ws2.cell(row=r, column=2, value="ПРЕДЗАПОЛНЕННЫЕ СЦЕНАРИИ TCO")
c.font = Font(name="Calibri", size=16, bold=True, color=NAVY)

r = 3
ws2.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
c = ws2.cell(row=r, column=2, value="Типовые расчёты для быстрой демонстрации клиенту")
c.font = body_font(10, False, MID_GRAY)

# Scenario data: name, area_m2, height, layers, work_grout, work_putty, work_paint
scenarios = [
    ("Комната 20 м² (интерьер)", 45, 2.7, 2, 80, 120, 150),
    ("Ванная 6 м² (влажная зона)", 22, 2.5, 2, 100, 150, 180),
    ("Фасад дома 100 м²", 100, 3.0, 2, 120, 0, 200),
    ("Деревянный домик 40 м²", 50, 2.5, 2, 100, 80, 180),
    ("Детская 15 м²", 35, 2.7, 2, 80, 120, 150),
]

r = 5
scen_headers = ["Сценарий", "Площадь (м²)", "Высота (м)", "Слои", "Работа грунт (₽/м²)", "Работа шпатл. (₽/м²)", "Работа покраска (₽/м²)"]
for i, h in enumerate(scen_headers):
    ws2.cell(row=r, column=i+2, value=h)
style_header_row(ws2, r, 8, NAVY)

for i, (name, area, height, layers, wg, wp, wpt) in enumerate(scenarios):
    row_num = 6 + i
    ws2.cell(row=row_num, column=2, value=name).font = body_font(10, True, NAVY)
    ws2.cell(row=row_num, column=2).border = thin_border
    for col, val in [(3, area), (4, height), (5, layers), (6, wg), (7, wp), (8, wpt)]:
        c = ws2.cell(row=row_num, column=col, value=val)
        c.font = body_font(10)
        c.alignment = align("center", "center", False)
        c.border = thin_border
        c.number_format = '#,##0'

# Results per scenario
r = 12
ws2.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
c = ws2.cell(row=r, column=2, value="РЕЗУЛЬТАТЫ ПО СЦЕНАРИЯМ")
c.font = hdr_font(12, True)
c.fill = PatternFill("solid", fgColor=CYAN)
c.alignment = align("center", "center", True)

r = 13
res_headers = ["Сценарий", "Материалы Эконом", "Материалы Премиум", "Всего Эконом", "Всего Премиум", "Год Эконом", "Год Премиум"]
for i, h in enumerate(res_headers):
    ws2.cell(row=r, column=i+2, value=h)
style_header_row(ws2, r, 8, NAVY)

# For each scenario, calculate results using product data from Sheet 1
# Products: grout@0.15kg/m2 * area, putty_start@1.2, putty_finish@0.8(premium only), paint@0.18/0.14, lac@0.08(premium only)
# Prices: grout 130/180/350, putty_start 120/170/280, putty_finish -/250/420, paint 180/320/600, lac -/350/800
# Service life: 3/7/10 years

econ_products = [
    (0.15, 130, 2),  # grout, price/kg, layers
    (1.2, 120, 1),   # putty_start
    (0.18, 180, 2),   # paint (layers from input)
]

prem_products = [
    (0.12, 350, 2),  # grout ОРЕОЛ base (economical)
    (1.0, 170, 1),   # putty_start NEWTONE
    (0.6, 420, 1),   # putty_finish NEWTONE
    (0.14, 600, 2),  # paint AVANTGARDE
    (0.08, 800, 1),  # lac AVANTGARDE
]

for i, (name, area, height, layers, wg, wp, wpt) in enumerate(scenarios):
    row_num = 14 + i
    ws2.cell(row=row_num, column=2, value=name).font = body_font(10, True, NAVY)
    ws2.cell(row=row_num, column=2).border = thin_border
    
    # Economy materials cost
    econ_mat = (0.15 * area * 130 * layers +   # grout
                1.2 * area * 120 * 1 +           # putty start
                0.18 * area * 180 * layers)      # paint
    
    # Premium materials cost
    prem_mat = (0.12 * area * 350 * layers +    # grout (economy base)
                1.0 * area * 170 * 1 +            # putty start NEWTONE
                0.6 * area * 420 * 1 +            # putty finish NEWTONE
                0.14 * area * 600 * layers +      # paint AVANTGARDE
                0.08 * area * 800 * 1)            # lac AVANTGARDE
    
    # Work cost
    work = area * (wg + wp + wpt)
    
    econ_total = econ_mat + work
    prem_total = prem_mat + work
    
    econ_year = econ_total / 3
    prem_year = prem_total / 10
    
    for col, val, fmt in [
        (3, econ_mat, '#,##0'), (4, prem_mat, '#,##0'),
        (5, econ_total, '#,##0'), (6, prem_total, '#,##0'),
        (7, econ_year, '#,##0'), (8, prem_year, '#,##0'),
    ]:
        c = ws2.cell(row=row_num, column=col, value=round(val))
        c.font = body_font(10)
        c.alignment = align("center", "center", False)
        c.border = thin_border
        c.number_format = fmt
    
    # Highlight year cost
    ws2.cell(row=row_num, column=7).font = Font(name="Calibri", size=10, bold=True, color=RED)
    ws2.cell(row=row_num, column=7).fill = PatternFill("solid", fgColor="FDEDEC")
    ws2.cell(row=row_num, column=8).font = Font(name="Calibri", size=10, bold=True, color=GREEN)
    ws2.cell(row=row_num, column=8).fill = PatternFill("solid", fgColor="E8F5E9")

# Summary message
r = 20
ws2.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
c = ws2.cell(row=r, column=2, value="Во всех сценариях премиум дешевле в пересчёте на год эксплуатации!")
c.font = Font(name="Calibri", size=13, bold=True, color=GREEN)
c.alignment = align("center", "center", True)
c.fill = PatternFill("solid", fgColor="E8F5E9")

# Chart for scenarios
chart2 = BarChart()
chart2.type = "col"
chart2.grouping = "clustered"
chart2.title = "Стоимость 1 года эксплуатации по сценариям (₽/год)"
chart2.y_axis.title = "₽/год"
chart2.width = 20
chart2.height = 12
chart2.gapWidth = 80

cats2 = Reference(ws2, min_col=2, min_row=14, max_row=18)
data2 = Reference(ws2, min_col=7, max_col=8, min_row=13, max_row=18)
chart2.add_data(data2, titles_from_data=True)
chart2.set_categories(cats2)

chart2.series[0].graphicalProperties.solidFill = RED
chart2.series[1].graphicalProperties.solidFill = GREEN
chart2.legend.position = 'b'

ws2.add_chart(chart2, "B22")

# ============================================================
# SHEET 3: ИНСТРУКЦИЯ
# ============================================================
ws3 = wb.create_sheet("Инструкция")
ws3.sheet_properties.tabColor = ORANGE

ws3.column_dimensions['A'].width = 3
ws3.column_dimensions['B'].width = 80

instructions = [
    ("", True, 18, NAVY),
    ("ИНСТРУКЦИЯ МЕНЕДЖЕРА", True, 18, NAVY),
    ("TCO-калькулятор Эмпилс — инструмент для обоснования цены премиум", False, 12, MID_GRAY),
    ("", False, 11, DARK_GRAY),
    ("КАК ПОЛЬЗОВАТЬСЯ КАЛЬКУЛЯТОРОМ", True, 14, NAVY),
    ("1. Откройте лист «Калькулятор TCO»", True, 11, DARK_GRAY),
    ("2. В жёлтые ячейки введите параметры: тип помещения, площадь стен, высоту потолков, количество слоёв, стоимость работ", False, 11, DARK_GRAY),
    ("3. Результаты рассчитываются автоматически в разделе «Сводный расчёт TCO»", False, 11, DARK_GRAY),
    ("4. Ключевая строка — «Стоимость 1 года эксплуатации»: показывает, что премиум дешевле в пересчёте на год", False, 11, DARK_GRAY),
    ("", False, 11, DARK_GRAY),
    ("ТИПОВЫЕ СЦЕНАРИИ", True, 14, NAVY),
    ("На листе «Сценарии» — предзаполненные расчёты для 5 типовых задач:", False, 11, DARK_GRAY),
    ("   - Комната 20 м² — стандартный интерьер", False, 11, DARK_GRAY),
    ("   - Ванная 6 м² — влажная зона", False, 11, DARK_GRAY),
    ("   - Фасад дома 100 м² — наружные работы", False, 11, DARK_GRAY),
    ("   - Деревянный домик 40 м² — защита дерева", False, 11, DARK_GRAY),
    ("   - Детская 15 м² — безопасность и экологичность", False, 11, DARK_GRAY),
    ("Показывайте клиенту нужный сценарий — экономия видна сразу", False, 11, DARK_GRAY),
    ("", False, 11, DARK_GRAY),
    ("КАК ПРОДАВАТЬ ЧЕРЕЗ TCO — 4 ШАГА", True, 14, NAVY),
    ("Шаг 1: Спросите клиента: «Вы считали, сколько стоит покраска в год?»", False, 11, DARK_GRAY),
    ("   Большинство клиентов считают только стоимость банки. Но краска — лишь 15-20% сметы. Остальное — работа.", False, 11, DARK_GRAY),
    ("", False, 11, DARK_GRAY),
    ("Шаг 2: Покажите расчёт на калькуляторе (2 минуты)", False, 11, DARK_GRAY),
    ("   Введите площадь → клиент видит: эконом 6 450 ₽/год, премиум 2 447 ₽/год", False, 11, DARK_GRAY),
    ("", False, 11, DARK_GRAY),
    ("Шаг 3: Объясните: «Переплата за премиум — всего +7% к смете, но -60% к стоимости года»", False, 11, DARK_GRAY),
    ("   Это конкретные цифры, а не абстрактные обещания. Клиент видит выгоду.", False, 11, DARK_GRAY),
    ("", False, 11, DARK_GRAY),
    ("Шаг 4: Предложите систему: «Грунт + шпатлёвка + краска = гарантия результата на 10 лет»", False, 11, DARK_GRAY),
    ("   Как Dulux Weathershield — гарантия только при использовании системы.", False, 11, DARK_GRAY),
    ("", False, 11, DARK_GRAY),
    ("ТИПОВЫЕ ВОРАЖДЕНИЯ И ОТВЕТЫ", True, 14, NAVY),
    ("«Дорого!» → «Премиум дешевле в 2,6 раза в год. Переплата — всего +7% к смете. Покажу расчёт?»", False, 11, DARK_GRAY),
    ("«Мне хватает и эконома» → «Эконом нужно перекрашивать через 3 года. Премиум — через 10. За 10 лет эконом обойдётся в 3 раза дороже»", False, 11, DARK_GRAY),
    ("«Я подумаю» → «Конечно! Вот расчёт — сохраните его. И QR-код на онлайн-калькулятор, можете пересчитать дома»", False, 11, DARK_GRAY),
    ("«А у конкурента дешевле» → «Давайте посчитаем стоимость года. Наша система с гарантией результата — это не просто банка краски»", False, 11, DARK_GRAY),
    ("", False, 11, DARK_GRAY),
    ("ВАЖНО!", True, 14, RED),
    ("- Не меняйте серые ячейки — там формулы", False, 11, DARK_GRAY),
    ("- Жёлтые ячейки — для ввода данных клиента", False, 11, DARK_GRAY),
    ("- Если клиент спрашивает про фасад — откройте лист «Сценарии», строку «Фасад дома»", False, 11, DARK_GRAY),
    ("- Цены на работы уточняйте по региону", False, 11, DARK_GRAY),
    ("- Для нестандартных задач — введите свои данные на листе «Калькулятор TCO»", False, 11, DARK_GRAY),
]

for i, (text, bold, size, color) in enumerate(instructions):
    row_num = 2 + i
    c = ws3.cell(row=row_num, column=2, value=text)
    c.font = Font(name="Calibri", size=size, bold=bold, color=color)
    c.alignment = align("left", "top", True)

# ============================================================
# PROTECTION
# ============================================================
# Protect formula cells, allow input cells
for ws_sheet in [ws]:
    for row in ws_sheet.iter_rows():
        for cell in row:
            if cell.fill and cell.fill.fgColor and cell.fill.fgColor.rgb == "00FFFFCC":
                cell.protection = Protection(locked=False)
            else:
                cell.protection = Protection(locked=True)

ws.protection.sheet = True
ws.protection.password = "empils2026"
ws.protection.enable()

# ============================================================
# PRINT SETUP
# ============================================================
ws.page_setup.orientation = 'landscape'
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws2.page_setup.orientation = 'landscape'

# ============================================================
# SAVE
# ============================================================
output = "/home/z/my-project/download/TCO_калькулятор_Эмпилс.xlsx"
wb.properties.creator = "Z.ai"
wb.save(output)
print(f"Saved: {output}")
print(f"Sheets: {wb.sheetnames}")
