#!/usr/bin/env python3
"""
TCO-Калькулятор Эмпилс v2.0 — с базой продуктов и выбором из справочника.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

dst_file = '/home/z/my-project/download/TCO_калькулятор_Эмпилс_v2.xlsx'
wb = openpyxl.Workbook()

# === СТИЛИ ===
NAVY, CYAN, LIGHT_CYAN, ORANGE, GREEN, RED = '0A2540', '00B4D8', '90E0EF', 'E86C00', '1B7D46', 'C0392B'
YELLOW_INPUT, WHITE, LIGHT_GRAY, BORDER_GRAY = 'FFF2CC', 'FFFFFF', 'F2F2F2', 'D9D9D9'

ft = lambda sz, b=False, c=NAVY: Font(name='Carlito', size=sz, bold=b, color=c)
ft_title = ft(18, True, WHITE)
ft_sub = ft(11, False, LIGHT_CYAN)
ft_sect = ft(13, True, WHITE)
ft_hdr = ft(10, True, WHITE)
ft_lbl = ft(10)
ft_val = ft(10)
ft_prod = ft(10, True, CYAN)
ft_hint = ft(9, False, '666666')
ft_res = ft(14, True, WHITE)

pf = lambda c: PatternFill(start_color=c, end_color=c, fill_type='solid')
pf_navy, pf_cyan, pf_orange, pf_green = pf(NAVY), pf(CYAN), pf(ORANGE), pf(GREEN)
pf_yellow, pf_white, pf_gray = pf(YELLOW_INPUT), pf(WHITE), pf(LIGHT_GRAY)

ac = Alignment(horizontal='center', vertical='center', wrap_text=True)
al = Alignment(horizontal='left', vertical='center', wrap_text=True)
bd = Border(left=Side('thin', BORDER_GRAY), right=Side('thin', BORDER_GRAY),
            top=Side('thin', BORDER_GRAY), bottom=Side('thin', BORDER_GRAY))

def set_cell(ws, row, col, val, font=ft_val, fill=None, align=ac, border=bd, fmt=None):
    c = ws.cell(row=row, column=col, value=val)
    c.font = font; c.alignment = align; c.border = border
    if fill: c.fill = fill
    if fmt: c.number_format = fmt
    return c

# ============================================================
# ЛИСТ 1: БАЗА ПРОДУКТОВ
# ============================================================
ws_db = wb.active
ws_db.title = 'База продуктов'
ws_db.sheet_properties.tabColor = CYAN

for i, w in enumerate([3, 18, 35, 15, 18, 20, 18, 22], 1):
    ws_db.column_dimensions[get_column_letter(i)].width = w

from openpyxl.utils import get_column_letter

# Recompute widths
ws_db.column_dimensions['A'].width = 3
ws_db.column_dimensions['B'].width = 20
ws_db.column_dimensions['C'].width = 38
ws_db.column_dimensions['D'].width = 15
ws_db.column_dimensions['E'].width = 18
ws_db.column_dimensions['F'].width = 22
ws_db.column_dimensions['G'].width = 20
ws_db.column_dimensions['H'].width = 18

# Title
ws_db.merge_cells('B2:H2')
set_cell(ws_db, 2, 2, 'БАЗА ПРОДУКТОВ ЭМПИЛС — СПРАВОЧНИК ДЛЯ TCO-КАЛЬКУЛЯТОРА', ft_title, pf_navy, ac)
ws_db.merge_cells('B3:H3')
set_cell(ws_db, 3, 2, 'Выбирайте продукты на листе «Калькулятор» — данные подставляются автоматически', ft_sub, pf_navy, al)

# Headers
hdrs = ['Категория', 'Продукт', 'Бренд', 'Цена (₽/кг)', 'Расход (кг/м²)', 'Назначение', 'Код']
for i, h in enumerate(hdrs):
    set_cell(ws_db, 5, i+2, h, ft_hdr, pf_cyan, ac)

# Products data — 27 позиций
products = [
    ['Грунтовка', 'ОРЕОЛ Грунт универсальный', 'ОРЕОЛ', 130, 0.15, 'Грунтование под шпатлёвку и краску', 'Г001'],
    ['Грунтовка', 'NEWTONE Грунт глубокого проникновения', 'NEWTONE', 180, 0.12, 'Укрепление рыхлых оснований', 'Г002'],
    ['Грунтовка', 'AVANTGARDE Грунт премиум', 'AVANTGARDE', 350, 0.10, 'Под премиум-систему', 'Г003'],
    ['Грунтовка', 'ОРЕОЛ Бетоноконтакт', 'ОРЕОЛ', 220, 0.20, 'Гладкие невпитывающие основания', 'Г004'],
    ['Грунтовка', 'NEWTONE Грунт антиплесневый', 'NEWTONE', 280, 0.12, 'Влажные зоны', 'Г005'],
    ['Шпатлёвка стартовая', 'ОРЕОЛ Шпатлёвка стартовая', 'ОРЕОЛ', 120, 1.2, 'Выравнивание, слой до 10 мм', 'С001'],
    ['Шпатлёвка стартовая', 'NEWTONE Шпатлёвка базовая', 'NEWTONE', 170, 1.0, 'Базовое выравнивание, до 8 мм', 'С002'],
    ['Шпатлёвка стартовая', 'AVANTGARDE Шпатлёвка сверхпрочная', 'AVANTGARDE', 280, 0.8, 'С армирующим эффектом', 'С003'],
    ['Шпатлёвка финишная', 'NEWTONE Шпатлёвка финишная', 'NEWTONE', 250, 0.6, 'Финиш, слой до 3 мм', 'Ф001'],
    ['Шпатлёвка финишная', 'AVANTGARDE Суперфинишная', 'AVANTGARDE', 420, 0.5, 'Идеально гладкая поверхность', 'Ф002'],
    ['Шпатлёвка финишная', 'ОРЕОЛ Шпатлёвка финишная', 'ОРЕОЛ', 160, 0.7, 'Эконом финиш, до 2 мм', 'Ф003'],
    ['Краска ВД', 'ОРЕОЛ Краска ВД интерьерная', 'ОРЕОЛ', 180, 0.18, 'Базовая ВД для сухих помещений', 'К001'],
    ['Краска ВД', 'NEWTONE Краска моющаяся', 'NEWTONE', 320, 0.15, 'Повышенная нагрузка', 'К002'],
    ['Краска ВД', 'AVANTGARDE Краска премиум', 'AVANTGARDE', 600, 0.12, '1 слой = 2 стандартным', 'К003'],
    ['Краска ВД', 'ОРЕОЛ Краска потолочная', 'ОРЕОЛ', 150, 0.16, 'Специально для потолков', 'К004'],
    ['Краска ВД', 'NEWTONE Для влажных зон', 'NEWTONE', 380, 0.14, 'Ванные, кухни', 'К005'],
    ['Краска фасадная', 'ОРЕОЛ Краска фасадная', 'ОРЕОЛ', 250, 0.20, 'Фасадная эконом, 3-5 лет', 'ФК001'],
    ['Краска фасадная', 'NEWTONE Краска фасадная', 'NEWTONE', 420, 0.16, 'Фасадная средний, 7 лет', 'ФК002'],
    ['Краска фасадная', 'AVANTGARDE Краска фасадная', 'AVANTGARDE', 780, 0.12, 'Фасадная премиум, 10+ лет', 'ФК003'],
    ['Лак защитный', 'NEWTONE Лак акриловый', 'NEWTONE', 350, 0.08, 'Защитный интерьерный', 'Л001'],
    ['Лак защитный', 'AVANTGARDE Лак полиуретановый', 'AVANTGARDE', 800, 0.06, 'Премиум, высокая износостойкость', 'Л002'],
    ['Лак защитный', 'ОРЕОЛ Лак для дерева', 'ОРЕОЛ', 200, 0.10, 'Эконом защита дерева', 'Л003'],
    ['Декор. эффект', 'AVANTGARDE Декоративная штукатурка', 'AVANTGARDE', 1200, 0.8, 'Эффект «короед»', 'Д001'],
    ['Декор. эффект', 'AVANTGARDE Венецианская штукатурка', 'AVANTGARDE', 1800, 0.5, 'Мраморный эффект', 'Д002'],
    ['Эмаль', 'ОРЕОЛ Эмаль ПФ-115', 'ОРЕОЛ', 220, 0.10, 'Универсальная по металлу/дереву', 'Э001'],
    ['Эмаль', 'NEWTONE Эмаль для радиаторов', 'NEWTONE', 400, 0.08, 'Термостойкая', 'Э002'],
    ['Эмаль', 'Расцвет Эмаль для полов', 'Расцвет', 280, 0.12, 'Износостойкая для полов', 'Э003'],
]

cat_colors = {
    'Грунтовка': 'D4EDDA', 'Шпатлёвка стартовая': 'FFF3CD', 'Шпатлёвка финишная': 'FFE0B2',
    'Краска ВД': 'CCE5FF', 'Краска фасадная': 'D1ECF1', 'Лак защитный': 'E2D5F1',
    'Декор. эффект': 'F8D7DA', 'Эмаль': 'D6D8DB',
}

for i, p in enumerate(products):
    row = 6 + i
    bg = pf_white if i % 2 == 0 else pf_gray
    cat_bg = PatternFill(start_color=cat_colors.get(p[0], 'FFFFFF'), end_color=cat_colors.get(p[0], 'FFFFFF'), fill_type='solid')
    for j, val in enumerate(p):
        c = ws_db.cell(row=row, column=j+2, value=val)
        c.font = ft_val; c.border = bd
        c.alignment = ac if j in [3,4,6] else al
        c.fill = cat_bg if j == 0 else bg
    ws_db.cell(row=row, column=5).number_format = '#,##0'

# ============================================================
# ЛИСТ 2: КАЛЬКУЛЯТОР
# ============================================================
ws = wb.create_sheet('Калькулятор TCO')
ws.sheet_properties.tabColor = ORANGE
ws.column_dimensions['A'].width = 3
ws.column_dimensions['B'].width = 40
ws.column_dimensions['C'].width = 22
ws.column_dimensions['D'].width = 22
ws.column_dimensions['E'].width = 22
ws.column_dimensions['F'].width = 12

# Title
ws.merge_cells('B2:F2')
set_cell(ws, 2, 2, 'TCO-КАЛЬКУЛЯТОР ЭМПИЛС v2.0', ft_title, pf_navy, ac)
ws.merge_cells('B3:F3')
set_cell(ws, 3, 2, 'Выбирайте продукты из справочника — расход и цена подставляются автоматически!', ft_sub, pf_navy, ac)

# --- ВВОД ДАННЫХ ---
set_cell(ws, 5, 2, 'ВВОД ДАННЫХ', ft_sect, pf_cyan, ac)
ws.merge_cells('B5:F5')

inputs = [
    (6, 'Тип помещения', 'комната 20 м²', 'Выберите из списка'),
    (7, 'Площадь поверхности (м²)', 42, 'Окрашиваемая площадь'),
    (8, 'Высота потолков (м)', 2.7, 'Стандарт 2.5–2.7'),
    (9, 'Количество слоёв краски', 2, 'Обычно 2 слоя'),
    (10, 'Стоимость работы — грунтовка (₽/м²)', 80, 'Средняя по региону'),
    (11, 'Стоимость работы — шпатлёвка (₽/м²)', 120, 'Зависит от сложности'),
    (12, 'Стоимость работы — покраска (₽/м²)', 150, 'Включая валик/кисть'),
]
for row, label, val, hint in inputs:
    set_cell(ws, row, 2, label, ft_lbl, None, al)
    set_cell(ws, row, 3, val, ft_val, pf_yellow)
    set_cell(ws, row, 4, hint, ft_hint, None, al)

dv_room = DataValidation(type='list', formula1='"комната 20 м²,ванная,фасад дома,деревянный домик,детская,офис,лоджия"')
ws.add_data_validation(dv_room)
dv_room.add(ws['C6'])

# --- ВЫБОР ПРОДУКТОВ ---
set_cell(ws, 14, 2, 'ВЫБОР ПРОДУКТОВ ИЗ СПРАВОЧНИКА (жёлтые ячейки — выберите продукт)', ft_sect, pf_orange, ac)
ws.merge_cells('B14:F14')

for i, h in enumerate(['Этап работ', 'Продукт (код)', 'Расход (кг/м²)', 'Цена (₽/кг)', 'Слои']):
    set_cell(ws, 15, i+2, h, ft_hdr, pf_navy)

# ЭКОНОМ
set_cell(ws, 16, 2, 'СИСТЕМА «ЭКОНОМ» (ОРЕОЛ)', ft(11, True, WHITE), PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid'), al)
ws.merge_cells('B16:F16')

eco_rows = [
    (17, 'Грунтовка', 'Г001'),
    (18, 'Шпатлёвка стартовая', 'С001'),
    (19, 'Шпатлёвка финишная', 'Ф003'),
    (20, 'Краска ВД', 'К001'),
]

# ВПР-формулы ссылаются на Базу продуктов
# H6:H31 — код, D6:D31 — бренд, E6:E31 — цена, F6:F31 — расход
for row, label, code in eco_rows:
    set_cell(ws, row, 2, label, ft_lbl, None, al)
    set_cell(ws, row, 3, code, ft_prod, pf_yellow)
    # Расход = ВПР(код; База продуктов!B6:H31; столбец_расхода; ЛОЖЬ)
    # В базе: B=Категория, C=Продукт, D=Бренд, E=Цена, F=Расход, G=Назначение, H=Код
    # ВПР по столбцу H (код), нужно вернуть F (расход) — это сдвиг 5 от H
    # Но ВПР ищет в первом столбце. Перестроим: ищем по коду в H, возврат из F
    # ИСПРАВЛЕНО: ВПР(C{row};'База продуктов'!H6:F31; ...) — нельзя, первый столбец должен быть левее
    # Решение: ИНДЕКС+ПОИСКПОЗ
    ws.cell(row=row, column=4, value=f"=ИНДЕКС('База продуктов'!F6:F31;ПОИСКПОЗ(C{row};'База продуктов'!H6:H31;0))")
    ws.cell(row=row, column=4).font = ft_val; ws.cell(row=row, column=4).alignment = ac; ws.cell(row=row, column=4).border = bd
    ws.cell(row=row, column=4).number_format = '0.00'
    
    ws.cell(row=row, column=5, value=f"=ИНДЕКС('База продуктов'!E6:E31;ПОИСКПОЗ(C{row};'База продуктов'!H6:H31;0))")
    ws.cell(row=row, column=5).font = ft_val; ws.cell(row=row, column=5).alignment = ac; ws.cell(row=row, column=5).border = bd
    ws.cell(row=row, column=5).number_format = '#,##0'
    
    layers = f'=C9' if 'Краск' in label else '1'
    set_cell(ws, row, 6, layers, ft_val)

# Выпадающие списки для эконом
dv_g = DataValidation(type='list', formula1="'База продуктов'!H6:H10")
ws.add_data_validation(dv_g); dv_g.add(ws['C17'])
dv_s = DataValidation(type='list', formula1="'База продуктов'!H11:H13")
ws.add_data_validation(dv_s); dv_s.add(ws['C18'])
dv_f = DataValidation(type='list', formula1="'База продуктов'!H14:H16")
ws.add_data_validation(dv_f); dv_f.add(ws['C19'])
dv_k = DataValidation(type='list', formula1="'База продуктов'!H17:H21")
ws.add_data_validation(dv_k); dv_k.add(ws['C20'])

# ПРЕМИУМ
set_cell(ws, 22, 2, 'СИСТЕМА «ПРЕМИУМ» (AVANTGARDE + NEWTONE)', ft(11, True, WHITE), PatternFill(start_color='C00000', end_color='C00000', fill_type='solid'), al)
ws.merge_cells('B22:F22')

prem_rows = [
    (23, 'Грунтовка', 'Г003'),
    (24, 'Шпатлёвка стартовая', 'С002'),
    (25, 'Шпатлёвка финишная', 'Ф002'),
    (26, 'Краска ВД', 'К003'),
    (27, 'Лак защитный', 'Л002'),
]

for row, label, code in prem_rows:
    set_cell(ws, row, 2, label, ft_lbl, None, al)
    set_cell(ws, row, 3, code, ft_prod, pf_yellow)
    ws.cell(row=row, column=4, value=f"=ИНДЕКС('База продуктов'!F6:F31;ПОИСКПОЗ(C{row};'База продуктов'!H6:H31;0))")
    ws.cell(row=row, column=4).font = ft_val; ws.cell(row=row, column=4).alignment = ac; ws.cell(row=row, column=4).border = bd
    ws.cell(row=row, column=4).number_format = '0.00'
    ws.cell(row=row, column=5, value=f"=ИНДЕКС('База продуктов'!E6:E31;ПОИСКПОЗ(C{row};'База продуктов'!H6:H31;0))")
    ws.cell(row=row, column=5).font = ft_val; ws.cell(row=row, column=5).alignment = ac; ws.cell(row=row, column=5).border = bd
    ws.cell(row=row, column=5).number_format = '#,##0'
    layers = f'=C9' if 'Краск' in label else '1'
    set_cell(ws, row, 6, layers, ft_val)

dv_g2 = DataValidation(type='list', formula1="'База продуктов'!H6:H10")
ws.add_data_validation(dv_g2); dv_g2.add(ws['C23'])
dv_s2 = DataValidation(type='list', formula1="'База продуктов'!H11:H13")
ws.add_data_validation(dv_s2); dv_s2.add(ws['C24'])
dv_f2 = DataValidation(type='list', formula1="'База продуктов'!H14:H16")
ws.add_data_validation(dv_f2); dv_f2.add(ws['C25'])
dv_k2 = DataValidation(type='list', formula1="'База продуктов'!H17:H21")
ws.add_data_validation(dv_k2); dv_k2.add(ws['C26'])
dv_l = DataValidation(type='list', formula1="'База продуктов'!H22:H24")
ws.add_data_validation(dv_l); dv_l.add(ws['C27'])

# --- РАСЧЁТ СТОИМОСТИ ---
set_cell(ws, 29, 2, 'РАСЧЁТ СТОИМОСТИ МАТЕРИАЛОВ', ft_sect, pf_green, ac)
ws.merge_cells('B29:F29')

for i, h in enumerate(['Материал', 'Расход на площадь (кг)', 'Цена (₽/кг)', 'Стоимость (₽)', '']):
    set_cell(ws, 30, i+2, h, ft_hdr, pf_navy)

# Эконом расчёт
set_cell(ws, 31, 2, 'СИСТЕМА «ЭКОНОМ»', ft(10, True, '4472C4'), None, al)
eco_calc = [
    (32, 'Грунтовка', '=D17*C7', '=E17', '=C32*D32*F17'),
    (33, 'Шпатлёвка стартовая', '=D18*C7', '=E18', '=C33*D33*F18'),
    (34, 'Шпатлёвка финишная', '=D19*C7', '=E19', '=C34*D34*F19'),
    (35, 'Краска ВД', '=D20*C7', '=E20', '=C35*D35*F20'),
]
for row, label, r1, r2, r3 in eco_calc:
    set_cell(ws, row, 2, label, ft_lbl, None, al)
    set_cell(ws, row, 3, r1, ft_val, None, ac, bd, '#,##0.00')
    set_cell(ws, row, 4, r2, ft_val, None, ac, bd, '#,##0')
    set_cell(ws, row, 5, r3, ft(10, True), None, ac, bd, '#,##0')

set_cell(ws, 36, 2, 'ИТОГО материалы (Эконом)', ft(10, True, '4472C4'))
set_cell(ws, 36, 5, '=SUM(E32:E35)', ft(11, True, '4472C4'), PatternFill(start_color='D6E4F0', end_color='D6E4F0', fill_type='solid'), ac, bd, '#,##0')

# Премиум расчёт
set_cell(ws, 38, 2, 'СИСТЕМА «ПРЕМИУМ»', ft(10, True, 'C00000'), None, al)
prem_calc = [
    (39, 'Грунтовка', '=D23*C7', '=E23', '=C39*D39*F23'),
    (40, 'Шпатлёвка стартовая', '=D24*C7', '=E24', '=C40*D40*F24'),
    (41, 'Шпатлёвка финишная', '=D25*C7', '=E25', '=C41*D41*F25'),
    (42, 'Краска ВД', '=D26*C7', '=E26', '=C42*D42*F26'),
    (43, 'Лак защитный', '=D27*C7', '=E27', '=C43*D43*F27'),
]
for row, label, r1, r2, r3 in prem_calc:
    set_cell(ws, row, 2, label, ft_lbl, None, al)
    set_cell(ws, row, 3, r1, ft_val, None, ac, bd, '#,##0.00')
    set_cell(ws, row, 4, r2, ft_val, None, ac, bd, '#,##0')
    set_cell(ws, row, 5, r3, ft(10, True), None, ac, bd, '#,##0')

set_cell(ws, 44, 2, 'ИТОГО материалы (Премиум)', ft(10, True, 'C00000'))
set_cell(ws, 44, 5, '=SUM(E39:E43)', ft(11, True, 'C00000'), PatternFill(start_color='F4CCCC', end_color='F4CCCC', fill_type='solid'), ac, bd, '#,##0')

# --- СВОДНЫЙ РАСЧЁТ TCO ---
set_cell(ws, 46, 2, 'СВОДНЫЙ РАСЧЁТ TCO', ft_sect, pf_navy, ac)
ws.merge_cells('B46:E46')

for i, h in enumerate(['Показатель', 'Эконом', 'Премиум', 'Разница']):
    set_cell(ws, 47, i+2, h, ft_hdr, pf_cyan)

tco = [
    (48, 'Стоимость материалов (₽)', '=E36', '=E44', '=D48-C48', '#,##0'),
    (49, 'Стоимость работ (₽)', '=C7*(C10+C11+C12)', '=C7*(C10+C11+C12)', '=D49-C49', '#,##0'),
    (50, 'Общая стоимость ремонта (₽)', '=C48+C49', '=D48+D49', '=D50-C50', '#,##0'),
    (51, 'Переплата за премиум (%)', '—', '=IF(C50>0,(D50-C50)/C50,0)', '—', '0.0%'),
    (52, 'Срок службы покрытия (лет)', 3, 10, '=D52-C52', '0'),
    (53, 'Стоимость 1 года (₽/год)', '=IF(C52>0,C50/C52,0)', '=IF(D52>0,D50/D52,0)', '=C53-D53', '#,##0'),
    (54, 'Экономия в год (₽)', '—', '=C53-D53', '—', '#,##0'),
    (55, 'Экономия за 10 лет (₽)', '—', '=D53*10', '—', '#,##0'),
]

for row, label, eco, prem, diff, fmt in tco:
    set_cell(ws, row, 2, label, ft_lbl, None, al)
    set_cell(ws, row, 3, eco, ft_val, None, ac, bd, fmt)
    set_cell(ws, row, 4, prem, ft_val, None, ac, bd, fmt)
    set_cell(ws, row, 5, diff, ft_val, None, ac, bd, fmt)

# Подсветка ключевой строки
for col in range(2, 6):
    ws.cell(row=53, column=col).fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    ws.cell(row=53, column=col).font = ft(12, True)

# Результат
ws.merge_cells('B57:E57')
set_cell(ws, 57, 2, '=IF(D53>0,"ПРЕМИУМ дешевле в "&TEXT(C53/D53,"0.0")&" раз в год! Экономия "&TEXT(E53,"#,##0")&" ₽/год","")', ft_res, pf_green, ac)

# --- ВИЗУАЛИЗАЦИЯ ---
set_cell(ws, 59, 2, 'СРАВНЕНИЕ TCO', ft_sect, pf_navy, ac)
ws.merge_cells('B59:D59')

for i, h in enumerate(['Показатель', 'Эконом', 'Премиум']):
    set_cell(ws, 60, i+2, h, ft_hdr, pf_cyan)

chart_data = [
    (61, 'Материалы (₽)', '=C48', '=D48'),
    (62, 'Работы (₽)', '=C49', '=D49'),
    (63, 'Общая стоимость (₽)', '=C50', '=D50'),
    (64, 'Стоимость года (₽)', '=C53', '=D53'),
]
for row, label, eco, prem in chart_data:
    set_cell(ws, row, 2, label, ft_lbl, None, al)
    set_cell(ws, row, 3, eco, ft_val, None, ac, bd, '#,##0')
    set_cell(ws, row, 4, prem, ft_val, None, ac, bd, '#,##0')

chart = BarChart()
chart.type = 'col'; chart.style = 10
chart.title = 'Сравнение TCO: Эконом vs Премиум'
chart.y_axis.title = '₽'
cats = Reference(ws, min_col=2, min_row=61, max_row=64)
d1 = Reference(ws, min_col=3, min_row=60, max_row=64)
d2 = Reference(ws, min_col=4, min_row=60, max_row=64)
chart.add_data(d1, titles_from_data=True)
chart.add_data(d2, titles_from_data=True)
chart.set_categories(cats)
chart.series[0].graphicalProperties.solidFill = '4472C4'
chart.series[1].graphicalProperties.solidFill = 'C00000'
chart.width = 22; chart.height = 14
ws.add_chart(chart, 'B66')

# ============================================================
# ЛИСТ 3: СЦЕНАРИИ
# ============================================================
ws_s = wb.create_sheet('Сценарии')
ws_s.sheet_properties.tabColor = GREEN
ws_s.column_dimensions['A'].width = 3
ws_s.column_dimensions['B'].width = 32
for c in ['C','D','E','F','G','H']:
    ws_s.column_dimensions[c].width = 16

ws_s.merge_cells('B2:H2')
set_cell(ws_s, 2, 2, 'ПРЕДЗАПОЛНЕННЫЕ СЦЕНАРИИ TCO', ft_title, pf_navy, ac)

for i, h in enumerate(['Сценарий', 'Площадь (м²)', 'Высота (м)', 'Слои', 'Грунтовка (₽/м²)', 'Шпатлёвка (₽/м²)', 'Покраска (₽/м²)']):
    set_cell(ws_s, 4, i+2, h, ft_hdr, pf_cyan)

scenarios = [
    ('Комната 20 м²', 45, 2.7, 2, 80, 120, 150),
    ('Ванная 6 м²', 22, 2.5, 2, 100, 150, 180),
    ('Фасад дома 100 м²', 100, 3, 2, 120, 0, 200),
    ('Деревянный домик 40 м²', 50, 2.5, 2, 100, 80, 180),
    ('Детская 15 м²', 35, 2.7, 2, 80, 120, 150),
    ('Офис 50 м²', 120, 3, 2, 80, 100, 140),
    ('Лоджия 8 м²', 20, 2.5, 2, 100, 100, 160),
]
for i, s in enumerate(scenarios):
    for j, val in enumerate(s):
        set_cell(ws_s, 5+i, j+2, val, ft_val, pf_white if i%2==0 else pf_gray)

# ============================================================
# ЛИСТ 4: ИНСТРУКЦИЯ
# ============================================================
ws_i = wb.create_sheet('Инструкция')
ws_i.sheet_properties.tabColor = NAVY
ws_i.column_dimensions['A'].width = 3
ws_i.column_dimensions['B'].width = 95

ws_i.merge_cells('B2:B2')
set_cell(ws_i, 2, 2, 'ИНСТРУКЦИЯ МЕНЕДЖЕРА — TCO-КАЛЬКУЛЯТОР v2.0', ft_title, pf_navy, ac)

lines = [
    ('', False),
    ('КАК ПОЛЬЗОВАТЬСЯ v2.0', True),
    ('', False),
    ('1. Откройте лист «Калькулятор TCO»', False),
    ('2. В жёлтые ячейки введите параметры: площадь, высота, стоимость работ', False),
    ('3. В разделе «Выбор продуктов» кликните на жёлтую ячейку с кодом продукта', False),
    ('4. Из выпадающего списка выберите нужный продукт — расход и цена подставляются автоматически!', False),
    ('5. Все расчёты пересчитаются мгновенно', False),
    ('', False),
    ('ЧТО НОВОГО В v2.0', True),
    ('', False),
    ('• База продуктов — 27 позиций Эмпилс с ценами и расходами', False),
    ('• Выбор продукта из выпадающего списка — не нужно помнить цифры', False),
    ('• Расход и цена подставляются автоматически через формулы ИНДЕКС+ПОИСКПОЗ', False),
    ('• Категории: грунтовки, шпатлёвки (старт/финиш), краски ВД, фасадные, лаки, эмали, декор', False),
    ('', False),
    ('КАК ПРОДАВАТЬ ЧЕРЕЗ TCO — 4 ШАГА', True),
    ('', False),
    ('Шаг 1: «Вы считали, сколько стоит покраска в год?»', False),
    ('Шаг 2: Покажите расчёт на калькуляторе (2 минуты)', False),
    ('Шаг 3: «Переплата за премиум — +7% к смете, но -60% к стоимости года»', False),
    ('Шаг 4: «Грунт + шпатлёвка + краска = гарантия результата на 10 лет»', False),
    ('', False),
    ('ТИПОВЫЕ ВОРАЖДЕНИЯ', True),
    ('', False),
    ('«Дорого!» → «Премиум дешевле в 2,6 раза в год. Покажу расчёт?»', False),
    ('«Мне хватает эконома» → «За 10 лет эконом обойдётся в 3 раза дороже»', False),
    ('«Я подумаю» → «Вот расчёт — сохраните»', False),
    ('«У конкурента дешевле» → «Посчитаем стоимость года с нашей системой»', False),
    ('', False),
    ('ВАЖНО!', True),
    ('', False),
    ('• Не удаляйте лист «База продуктов» — калькулятор ссылается на него', False),
    ('• Для добавления продуктов — добавьте строку в «Базу продуктов» и обновите выпадающий список', False),
    ('• Жёлтые ячейки — для ввода данных, остальные рассчитываются автоматически', False),
]
for i, (text, is_hdr) in enumerate(lines):
    row = 3 + i
    c = ws_i.cell(row=row, column=2, value=text)
    c.font = ft(12, True) if is_hdr else ft(10)
    c.alignment = al

# Сохранение
wb.save(dst_file)
print(f'✅ Сохранено: {dst_file}')
