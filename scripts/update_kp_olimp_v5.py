"""
Updated KP v5:
- Remove logic chain at bottom
- Add section 4.1: Наценки (первичная 35% + вторичная 50%) before ETP section
- Add section 7: Поддержка маркетинга after Assortment
- Renumber: 1-Naselenie, 2-Potencial, 3-Tekushiy rynok, 3.1-Gap, 4-Empils 8%, 4.1-Nacenki, 5-ETP, 6-Prognoz, 7-Assortiment, 8-Podderzhka
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
import os

OUT = '/home/z/my-project/download/КП_Олимп_коммерческое_предложение.docx'

# ── Key numbers ─────────────────────────────────────────────────────
market_rub_bln = 3.3
actual_bln = 1.1
gap_bln = 2.7
empils_8pct_mln = 312
year1 = 60; year2 = 120; year3 = 312
margin_pct = 0.35
earn1 = year1 * margin_pct; earn2 = year2 * margin_pct; earn3 = year3 * margin_pct

regions = [
    {'name': 'Новосибирская область', 'pop': 2786500, 'cities': 'Новосибирск', 'city_pop': 1633000},
    {'name': 'Кемеровская область',  'pop': 2527200, 'cities': 'Кемерово', 'city_pop': 560000},
    {'name': 'Алтайский край',        'pop': 2099200, 'cities': 'Барнаул', 'city_pop': 640000},
    {'name': 'Томская область',       'pop': 1075000,  'cities': 'Томск', 'city_pop': 578000},
    {'name': 'Республика Алтай',     'pop': 210100,   'cities': 'Горно-Алтайск', 'city_pop': 64000},
]
total_pop = sum(r['pop'] for r in regions)
total_city_pop = sum(r['city_pop'] for r in regions)
per_capita_kg = 14.59; per_capita_rub = 1123; avg_price_per_ton = 77000
market_tons = total_city_pop * per_capita_kg / 1000
actual_mln = int(actual_bln * 1000)
actual_tons = actual_mln * 1e6 / avg_price_per_ton
gap_mln = int(gap_bln * 1000); gap_tons = market_tons - actual_tons
gap_pct = (1 - actual_mln / (market_rub_bln * 1000)) * 100
empils_tons = empils_8pct_mln * 1e6 / avg_price_per_ton
akb_target = 1200; etp_covers = 120; etp_count = 4; etp_total_rtt = etp_count * etp_covers
etp_revenue = 40e6; distributor_tp_revenue = 20e6; assortment = 300

distributors = [
    {'name': 'Сибхозторг/Стройторг', 'portfolio': 'Рогнеда, Кушнаренко, Декарт, ЭСКАРО, НБХ', 'territory': 'Новосиб. обл., Алт. край'},
    {'name': 'ООО «Олимп»', 'portfolio': 'АВС Фарбен, Радуга, НБХ, ЭСКАРО, ДЕКАРТ', 'territory': 'Алт. край, Новосиб., Кемерово, Томск'},
    {'name': 'ООО «АЛЬТЕРРА»', 'portfolio': 'Лакра, Рогнеда, Мефферт', 'territory': 'Новосиб. обл., Алт. край'},
    {'name': 'Стройпарк 1ГМ', 'portfolio': 'Тиккурила, Рогнеда, Декарт', 'territory': 'г. Томск'},
    {'name': 'ООО «Нисса»', 'portfolio': 'Лакра, Континенталь', 'territory': 'Алтайский край'},
    {'name': 'ООО «Колорлон»', 'portfolio': 'Рогнеда, Лакра, Декарт, Мефферт', 'territory': 'Новосиб. обл.'},
    {'name': 'ООО «АНА»', 'portfolio': 'Тиккурила', 'territory': 'Омская обл.'},
    {'name': 'ИП Шишков С.В.', 'portfolio': 'ТД Радуга, АВС Фарбен, Рогнеда, ВИТ, НБХ', 'territory': 'Омская обл.'},
    {'name': 'ИП Миронкин Е.В.', 'portfolio': 'Лакра', 'territory': 'Омская обл.'},
    {'name': 'ИП Ковалев А.А.', 'portfolio': 'Кушнаренко, Новый Дом', 'territory': 'Омская обл.'},
]

# ── Create document ───────────────────────────────────────────────────
doc = Document()
for section in doc.sections:
    section.top_margin = Cm(1.3); section.bottom_margin = Cm(1.3)
    section.left_margin = Cm(1.8); section.right_margin = Cm(1.8)

INK = RGBColor(0x2F, 0x3A, 0x48); MUTE = RGBColor(0x8A, 0x95, 0xA3)
RED = RGBColor(0xC0, 0x00, 0x00); SAND = RGBColor(0xC9, 0xA7, 0x7C)
BLUE = RGBColor(0x6B, 0x8F, 0xA8); GREEN = RGBColor(0x2E, 0xCC, 0x71)
WHITE = RGBColor(0xFF, 0xFF, 0xFF); DARK_GREEN = RGBColor(0x1A, 0x7D, 0x44)

def shade_cell(cell, color_hex):
    shading = cell._element.get_or_add_tcPr()
    shd = shading.makeelement(qn('w:shd'), {qn('w:fill'): color_hex, qn('w:val'): 'clear'})
    shading.append(shd)

def set_cell(cell, text, size=8, bold=False, color=INK, align=WD_ALIGN_PARAGRAPH.LEFT):
    cell.text = ''; p = cell.paragraphs[0]; p.alignment = align
    run = p.add_run(text)
    run.font.size = Pt(size); run.font.bold = bold; run.font.color.rgb = color; run.font.name = 'Inter'

def add_heading(text, color=RED, size=11, space_before=10, space_after=6):
    p = doc.add_paragraph(); p.space_before = Pt(space_before); p.space_after = Pt(space_after)
    run = p.add_run(text)
    run.font.size = Pt(size); run.font.bold = True; run.font.color.rgb = color; run.font.name = 'Inter'
    return p

def add_body(text, bold=False, color=INK, size=9, space_after=4):
    p = doc.add_paragraph(); p.space_after = Pt(space_after)
    run = p.add_run(text)
    run.font.size = Pt(size); run.font.bold = bold; run.font.color.rgb = color; run.font.name = 'Inter'
    return p

def add_bullet(text, color=INK, size=9, space_after=2):
    p = doc.add_paragraph(); p.space_after = Pt(space_after)
    p.paragraph_format.left_indent = Cm(0.5)
    run = p.add_run(f'·  {text}')
    run.font.size = Pt(size); run.font.color.rgb = color; run.font.name = 'Inter'
    return p

# ════════════════════════════════════════════════════════════════════════
# HERO BLOCK
# ════════════════════════════════════════════════════════════════════════
title = doc.add_paragraph(); title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run('Коммерческое предложение ЗАО «Эмпилс»')
run.font.size = Pt(16); run.font.bold = True; run.font.color.rgb = INK; run.font.name = 'Inter'

subtitle = doc.add_paragraph(); subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = subtitle.add_run('Партнёр: Олимп · Новосибирская, Томская, Кемеровская обл., Алтайский край, Республика Алтай')
run.font.size = Pt(9); run.font.italic = True; run.font.color.rgb = MUTE; run.font.name = 'Inter'

# Hero table
hero = doc.add_table(rows=2, cols=4); hero.style = 'Table Grid'; hero.alignment = WD_TABLE_ALIGNMENT.CENTER
hero_data = [
    {'label': 'ПОТЕНЦИАЛ РЫНКА', 'value': f'{market_rub_bln} млрд ₽', 'sub': f'{market_tons:,.0f} т'.replace(',', ' '), 'color': '6B8FA8', 'text_color': BLUE},
    {'label': 'ФАКТИЧЕСКИЙ ТО', 'value': f'{actual_bln} млрд ₽', 'sub': f'{actual_tons:,.0f} т'.replace(',', ' '), 'color': 'C9A77C', 'text_color': SAND},
    {'label': 'НЕДОЗАГРУЗКА', 'value': f'{gap_bln} млрд ₽', 'sub': f'{gap_pct:.0f}% рынка', 'color': 'C00000', 'text_color': RED},
    {'label': 'ДОЛЯ ЭМПИЛС 8%', 'value': f'{empils_8pct_mln} млн ₽', 'sub': 'цель на 3 года', 'color': '2ECC71', 'text_color': DARK_GREEN},
]
for j, hd in enumerate(hero_data):
    cell = hero.rows[0].cells[j]; cell.text = ''; p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(hd['label']); run.font.size = Pt(7); run.font.bold = True; run.font.color.rgb = WHITE; run.font.name = 'Inter'
    shade_cell(cell, hd['color'])
    cell = hero.rows[1].cells[j]; cell.text = ''
    p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(hd['value']); run.font.size = Pt(14); run.font.bold = True; run.font.color.rgb = hd['text_color']; run.font.name = 'Roboto Mono'
    p2 = cell.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = p2.add_run(hd['sub']); run2.font.size = Pt(7); run2.font.color.rgb = MUTE; run2.font.name = 'Inter'
    shade_cell(cell, 'F5F2EC')

doc.add_paragraph()

# Revenue forecast + distributor earnings
add_heading('ТОВАРООБОРОТ И ЗАРАБОТОК ПАРТНЁРА ПО ГОДАМ', color=DARK_GREEN, size=11)
forecast = doc.add_table(rows=5, cols=4); forecast.style = 'Table Grid'; forecast.alignment = WD_TABLE_ALIGNMENT.CENTER
for j, h in enumerate(['Год', 'Товарооборот', 'Наценка 35% (доход партнёра)', 'Комментарий']):
    set_cell(forecast.rows[0].cells[j], h, size=9, bold=True, color=WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)
    shade_cell(forecast.rows[0].cells[j], '2F3A48')
fc_rows = [
    ['Год 1', f'{year1} млн ₽', f'{earn1:.0f} млн ₽', '4 ЭТП + ТП дистрибьютора, 300 SKU'],
    ['Год 2', f'{year2} млн ₽', f'{earn2:.0f} млн ₽', 'Расширение АКБ до 800+ РТТ'],
    ['Год 3', f'{year3} млн ₽', f'{earn3:.0f} млн ₽', 'Полное покрытие 8% потенциала'],
    ['ИТОГО', f'{year1+year2+year3} млн ₽', f'{earn1+earn2+earn3:.0f} млн ₽', 'За 3 года'],
]
for i, row_data in enumerate(fc_rows):
    row = forecast.rows[i + 1]; is_total = i == 3
    for j, val in enumerate(row_data):
        set_cell(row.cells[j], val, size=9, bold=is_total, color=(RED if is_total else INK),
                 align=(WD_ALIGN_PARAGRAPH.LEFT if j == 3 else WD_ALIGN_PARAGRAPH.CENTER))
        if is_total: shade_cell(row.cells[j], 'FFF5E6')
        elif i % 2 == 0: shade_cell(row.cells[j], 'F5F2EC')

p = doc.add_paragraph()
run = p.add_run(f'Цель: {year3} млн ₽ товарооборота в Год 3 = 8% потенциала региона. Заработок партнёра: {earn3:.0f} млн ₽/год при наценке 35%.')
run.font.size = Pt(9); run.font.bold = True; run.font.color.rgb = DARK_GREEN; run.font.name = 'Inter'

divider = doc.add_paragraph(); divider.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = divider.add_run('—' * 60); run.font.color.rgb = MUTE; run.font.size = Pt(8)
add_body('↓ ДЕТАЛИЗАЦИЯ РАСЧЁТА ↓', bold=True, color=MUTE, size=9)

# ════════════════════════════════════════════════════════════════════════
# DETAIL SECTION
# ════════════════════════════════════════════════════════════════════════

# 1. Population
add_heading('1. ПОТЕНЦИАЛ РЕГИОНА — НАСЕЛЕНИЕ')
table = doc.add_table(rows=len(regions) + 2, cols=4); table.style = 'Table Grid'; table.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, h in enumerate(['Регион', 'Население (чел.)', 'Ключевые города', 'Города (чел.)']):
    set_cell(table.rows[0].cells[i], h, size=8, bold=True, color=WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)
    shade_cell(table.rows[0].cells[i], '2F3A48')
for i, r in enumerate(regions):
    row = table.rows[i + 1]
    set_cell(row.cells[0], r['name'], size=8)
    set_cell(row.cells[1], f'{r["pop"]:,}'.replace(',', ' '), size=8, align=WD_ALIGN_PARAGRAPH.RIGHT)
    set_cell(row.cells[2], r['cities'], size=8)
    set_cell(row.cells[3], f'{r["city_pop"]:,}'.replace(',', ' '), size=8, align=WD_ALIGN_PARAGRAPH.RIGHT)
    if i % 2 == 0:
        for j in range(4): shade_cell(row.cells[j], 'F5F2EC')
set_cell(table.rows[-1].cells[0], 'ИТОГО', size=8, bold=True)
set_cell(table.rows[-1].cells[1], f'{total_pop:,}'.replace(',', ' '), size=8, bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
set_cell(table.rows[-1].cells[2], '5 городов', size=8, bold=True)
set_cell(table.rows[-1].cells[3], f'{total_city_pop:,}'.replace(',', ' '), size=8, bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
for j in range(4): shade_cell(table.rows[-1].cells[j], 'E8E3D9')
add_body(f'База расчёта: городское население — {total_city_pop:,} чел.'.replace(',', ' '), bold=True)

# 2. Market potential
add_heading('2. ПОТЕНЦИАЛ РЫНКА ЛКМ (по городскому населению)')
mt = doc.add_table(rows=5, cols=3); mt.style = 'Table Grid'; mt.alignment = WD_TABLE_ALIGNMENT.CENTER
m_data = [
    ['Параметр', 'В рублях', 'В тоннах'],
    ['Потребление ЛКМ на душу', f'~{per_capita_rub} ₽/чел.', f'~{per_capita_kg} кг/чел.'],
    ['Расчёт', f'{total_city_pop:,} × {per_capita_rub} ₽'.replace(',', ' '), f'{total_city_pop:,} × {per_capita_kg} кг'.replace(',', ' ')],
    ['Средняя цена', f'~{avg_price_per_ton//1000} 000 ₽/т', '—'],
    ['ПОТЕНЦИАЛ РЫНКА ЛКМ', f'{market_rub_bln} млрд ₽', f'{market_tons:,.0f} т'.replace(',', ' ')],
]
for i, rd in enumerate(m_data):
    row = mt.rows[i]
    for j, val in enumerate(rd):
        is_h = i == 0; is_t = i == 4
        set_cell(row.cells[j], val, size=8, bold=(is_h or is_t),
                 color=(WHITE if is_h else (RED if is_t else INK)),
                 align=(WD_ALIGN_PARAGRAPH.CENTER if j > 0 else WD_ALIGN_PARAGRAPH.LEFT))
        if is_h: shade_cell(row.cells[j], '2F3A48')
        elif is_t: shade_cell(row.cells[j], 'FFF5E6')
        elif i % 2 == 1: shade_cell(row.cells[j], 'F5F2EC')

# 3. Current market
add_heading('3. ТЕКУЩИЙ РЫНОК — ФАКТИЧЕСКИЕ ДИСТРИБЬЮТОРЫ')
dt = doc.add_table(rows=len(distributors) + 2, cols=3); dt.style = 'Table Grid'; dt.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, h in enumerate(['Дистрибьютор', 'Портфель конкурентов', 'Территория']):
    set_cell(dt.rows[0].cells[i], h, size=8, bold=True, color=WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)
    shade_cell(dt.rows[0].cells[i], '2F3A48')
for i, d in enumerate(distributors):
    row = dt.rows[i + 1]
    set_cell(row.cells[0], d['name'], size=7.5, bold=True)
    set_cell(row.cells[1], d['portfolio'], size=7.5)
    set_cell(row.cells[2], d['territory'], size=7.5)
    if i % 2 == 0:
        for j in range(3): shade_cell(row.cells[j], 'F5F2EC')
set_cell(dt.rows[-1].cells[0], 'ИТОГО (10 дистрибьюторов)', size=8, bold=True)
set_cell(dt.rows[-1].cells[1], f'Совокупный ТО: {actual_bln} млрд ₽', size=8, bold=True, color=RED)
set_cell(dt.rows[-1].cells[2], f'~{actual_tons:,.0f} т'.replace(',', ' '), size=8, bold=True, color=RED)
for j in range(3): shade_cell(dt.rows[-1].cells[j], 'FFF5E6')

# 3.1 Gap
add_heading('НЕДОЗАГРУЗКА РЕГИОНА', color=RED, size=10, space_before=8, space_after=4)
gt = doc.add_table(rows=4, cols=4); gt.style = 'Table Grid'; gt.alignment = WD_TABLE_ALIGNMENT.CENTER
g_data = [
    ['Показатель', 'Млрд ₽', 'Тонн', 'Доля'],
    ['Теоретический потенциал', f'{market_rub_bln}', f'{market_tons:,.0f}'.replace(',', ' '), '100%'],
    ['Фактический ТО дистрибьюторов', f'{actual_bln}', f'{actual_tons:,.0f}'.replace(',', ' '), f'{actual_bln/market_rub_bln*100:.0f}%'],
    ['НЕДОЗАГРУЗКА', f'{gap_bln}', f'{gap_tons:,.0f}'.replace(',', ' '), f'{gap_pct:.0f}%'],
]
for i, rd in enumerate(g_data):
    row = gt.rows[i]
    for j, val in enumerate(rd):
        is_h = i == 0; is_g = i == 3
        set_cell(row.cells[j], val, size=8, bold=(is_h or is_g),
                 color=(WHITE if is_h else (RED if is_g else INK)),
                 align=(WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.CENTER))
        if is_h: shade_cell(row.cells[j], '2F3A48')
        elif is_g: shade_cell(row.cells[j], 'FFE0E0')
        elif i % 2 == 1: shade_cell(row.cells[j], 'F5F2EC')
p = doc.add_paragraph()
run = p.add_run(f'⚠ Регион недозагружен на {gap_pct:.0f}% — потенциал {gap_bln} млрд ₽ не освоен')
run.font.size = Pt(9); run.font.bold = True; run.font.color.rgb = RED; run.font.name = 'Inter'

# 4. Empils 8%
add_heading('4. ДОЛЯ ЭМПИЛС — 8% РЫНКА')
et = doc.add_table(rows=3, cols=3); et.style = 'Table Grid'; et.alignment = WD_TABLE_ALIGNMENT.CENTER
e_data = [
    ['Показатель', 'В рублях', 'В тоннах'],
    ['Потенциал рынка (100%)', f'{market_rub_bln} млрд ₽', f'{market_tons:,.0f} т'.replace(',', ' ')],
    ['Доля Эмпилс (8%)', f'{empils_8pct_mln} млн ₽', f'{empils_tons:,.0f} т'.replace(',', ' ')],
]
for i, rd in enumerate(e_data):
    row = et.rows[i]
    for j, val in enumerate(rd):
        is_h = i == 0; is_e = i == 2
        set_cell(row.cells[j], val, size=8, bold=(is_h or is_e),
                 color=(WHITE if is_h else (RED if is_e else INK)),
                 align=(WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.CENTER))
        if is_h: shade_cell(row.cells[j], '2F3A48')
        elif is_e: shade_cell(row.cells[j], 'FFF5E6')
        else: shade_cell(row.cells[j], 'F5F2EC')
add_body(f'{empils_8pct_mln} млн ₽ — целевой товарооборот Эмпилс, достигаемый за 3 года. Покрытие АКБ: {akb_target} РТТ.', bold=True, color=DARK_GREEN)

# ── 4.1 НАЦЕНКИ (NEW — before ETP section) ──────────────────────────
add_heading('4.1. ЦЕНООБРАЗОВАНИЕ — НАЦЕНКИ', color=INK, size=10, space_before=8, space_after=4)

nc_table = doc.add_table(rows=4, cols=4); nc_table.style = 'Table Grid'; nc_table.alignment = WD_TABLE_ALIGNMENT.CENTER
nc_data = [
    ['Уровень', 'Цепочка', 'Наценка', 'Описание'],
    ['Первичная', 'Завод → Дистрибьютор', '35%', 'Базовая наценка дистрибьютора при закупке у Эмпилс'],
    ['Вторичная', 'Дистрибьютор → РТТ', '50%', 'Наценка при продаже в розничную торговую точку'],
    ['Итого', 'Завод → РТТ', '~117%', 'Совокупная наценка от заводской цены до цены на полке'],
]
for i, rd in enumerate(nc_data):
    row = nc_table.rows[i]
    for j, val in enumerate(rd):
        is_h = i == 0; is_t = i == 3
        set_cell(row.cells[j], val, size=8, bold=(is_h or is_t),
                 color=(WHITE if is_h else (RED if is_t else INK)),
                 align=(WD_ALIGN_PARAGRAPH.LEFT if j == 0 or j == 3 else WD_ALIGN_PARAGRAPH.CENTER))
        if is_h: shade_cell(row.cells[j], '2F3A48')
        elif is_t: shade_cell(row.cells[j], 'FFF5E6')
        elif i % 2 == 0: shade_cell(row.cells[j], 'F5F2EC')

add_body('Пример: при заводской цене 100 ₽ → дистрибьютор продаёт в РТТ по 135 ₽ → РТТ продаёт потребителю по ~202 ₽.', size=8, color=MUTE)

# 5. ETP
add_heading('5. ОРГАНИЗАЦИЯ ПРОДАЖ — ЭКСКЛЮЗИВНЫЕ ТП')
etp_data = [
    ['Параметр', 'Значение'],
    ['1 ЭТП покрывает', f'{etp_covers} РТТ'],
    ['4 ЭТП покрывают', f'{etp_total_rtt} РТТ'],
    ['Товарооборот через ЭТП', f'{etp_revenue/1e6:.0f} млн ₽/год'],
    ['Товарооборот через ТП дистрибьютора', f'{distributor_tp_revenue/1e6:.0f} млн ₽/год'],
    ['ИТОГО год 1', f'{year1} млн ₽'],
]
etp_t = doc.add_table(rows=len(etp_data), cols=2); etp_t.style = 'Table Grid'; etp_t.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, (param, value) in enumerate(etp_data):
    row = etp_t.rows[i]
    set_cell(row.cells[0], param, size=9, bold=(i == 0 or i == len(etp_data)-1), color=(WHITE if i == 0 else INK))
    set_cell(row.cells[1], value, size=9, bold=(i == 0 or i == len(etp_data)-1),
             color=(WHITE if i == 0 else (RED if i == len(etp_data)-1 else INK)), align=WD_ALIGN_PARAGRAPH.CENTER)
    if i == 0: shade_cell(row.cells[0], '2F3A48'); shade_cell(row.cells[1], '2F3A48')
    elif i == len(etp_data)-1: shade_cell(row.cells[0], 'FFF5E6'); shade_cell(row.cells[1], 'FFF5E6')
    elif i % 2 == 0: shade_cell(row.cells[0], 'F5F2EC'); shade_cell(row.cells[1], 'F5F2EC')

# 6. Revenue forecast (detail)
add_heading('6. ПРОГНОЗ ТОВАРООБОРОТА ПО ГОДАМ')
fc2 = doc.add_table(rows=4, cols=4); fc2.style = 'Table Grid'; fc2.alignment = WD_TABLE_ALIGNMENT.CENTER
fc2_data = [
    ['Год', 'Товарооборот', 'Рост', 'Комментарий'],
    ['Год 1', f'{year1} млн ₽', '—', '4 ЭТП + ТП дистрибьютора, 300 SKU'],
    ['Год 2', f'{year2} млн ₽', '×2', 'Расширение АКБ до 800+ РТТ'],
    ['Год 3', f'{year3} млн ₽', '×5,2', 'Полное покрытие 8% потенциала'],
]
for i, row_data in enumerate(fc2_data):
    row = fc2.rows[i]
    for j, val in enumerate(row_data):
        set_cell(row.cells[j], val, size=9, bold=(i == 0), color=(WHITE if i == 0 else INK),
                 align=(WD_ALIGN_PARAGRAPH.LEFT if j == 0 or j == 3 else WD_ALIGN_PARAGRAPH.CENTER))
        if i == 0: shade_cell(row.cells[j], '2F3A48')
        elif i % 2 == 1: shade_cell(row.cells[j], 'F5F2EC')

# 7. Assortment
add_heading('7. АССОРТИМЕНТ НАЧАЛА ВЗАИМОДЕЙСТВИЯ')
add_body(f'Стартовый ассортимент: {assortment} SKU из портфеля Эмпилс', bold=True)
add_body('Бренды: ОРЕОЛ · ОРЕОЛ ДИСКОНТ · NEWTONE · AVANTGARDE · TITAN · PROFIWOOD')
add_body('Категории: краски интерьерные · фасадные · эмали · грунт-эмали · грунтовки · шпатлёвки · клеи')

# ── 8. Поддержка маркетинга (NEW) ───────────────────────────────────
add_heading('8. ПОДДЕРЖКА МАРКЕТИНГА', color=DARK_GREEN, size=11)

support_table = doc.add_table(rows=6, cols=2); support_table.style = 'Table Grid'; support_table.alignment = WD_TABLE_ALIGNMENT.CENTER
support_data = [
    ['Программа', 'Описание'],
    ['Программа мотивации общего прайса', 'Бонусы и ретро-отсрочки для ТП, продающих весь ассортимент Эмпилс'],
    ['Мотивация сотрудников РТТ', 'Программа поощрения продавцов торговых точек за рекомендации продукции Эмпилс'],
    ['Продвижение фокусных банок', 'Целевое продвижение топ-SKU (TITAN, NEWTONE) с повышенной маржой'],
    ['Гибкая отсрочка', 'Индивидуальные условия отсрочки платежа в зависимости от объёма закупки и истории сотрудничества'],
    ['Команда ЭТП', '4 эксклюзивных торговых представителя, обученных и мотивируемых Эмпилс'],
]
for i, (prog, desc) in enumerate(support_data):
    row = support_table.rows[i]
    set_cell(row.cells[0], prog, size=8, bold=(i == 0), color=(WHITE if i == 0 else INK))
    set_cell(row.cells[1], desc, size=8, bold=(i == 0), color=(WHITE if i == 0 else INK))
    if i == 0: shade_cell(row.cells[0], '2F3A48'); shade_cell(row.cells[1], '2F3A48')
    elif i % 2 == 0: shade_cell(row.cells[0], 'F5F2EC'); shade_cell(row.cells[1], 'F5F2EC')

doc.save(OUT)
print(f"DOCX: {OUT} ({os.path.getsize(OUT)} bytes)")
