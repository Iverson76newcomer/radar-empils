"""
Updated commercial proposal for Олимп with:
1. Market potential in BOTH rubles AND tons
2. Current distributor table with actual turnover — "ТО, млн" = 1195 is the SUM of all distributors combined
3. "Недозагрузка региона" = theoretical potential - actual distributor turnover
4. Compact, clean tables

Key calculations:
- Russia LKМ: 2.13M tons, 146M people → per capita = 2.13M/146M = 14.6 kg/person
- Region: 8,698,000 × 14.6 kg = ~127,000 tons LKМ market
- In rubles: 127,000 tons × 77,000 ₽/ton = 9.77 млрд ₽
- Actual distributor turnover (from user's table): 1,195 млн ₽ total = 1.195 млрд ₽
- BUT: 1195 is the same for ALL rows — this is the TOTAL for the region, not per distributor
- Need to interpret: the user said "фактическую информацию о дистрибьюторах с их общим оборотом"
  The "1195" is likely the COMBINED turnover of ALL distributors in the region = 1,195 млн ₽ = 1.195 млрд ₽
- Theoretical potential: 9.77 млрд ₽
- Actual: 1.195 млрд ₽ (only ~12.2% of theoretical!)
- Недозагрузка: 9.77 - 1.195 = 8.575 млрд ₽ (87.8% of market not served!)

Wait — 1.195 млрд seems too low for 5 regions with 8.7M people.
Actually, re-reading: "ТО, млн" = 1195 means 1195 млн = 1.195 млрд ₽.
But this might be just the LKМ turnover of THESE specific distributors (competitors), not the whole region.
Or it could be 1195 as total competitive LKМ sales in the region.

Let me interpret it as: the combined turnover of these 10 distributors = 1,195 млн ₽ (~1.2 млрд ₽).
This represents the ACTUAL competitive LKМ market in the region.
Theoretical: 9.77 млрд ₽ → actual: 1.2 млрд ₽ → gap = 8.57 млрд ₽ (87.7% not covered)

Actually, re-checking: the user said "фактическую информацию о дистрибьюторах 
которые работают на этой территории с их общим оборотом" — so 1195 млн = 
combined LKМ turnover of all listed distributors.

But 1195 млн = 1.195 млрд for 10 distributors across 5 regions seems reasonable if 
these are LKМ-specific sales (not total company turnover).

Calculations:
- Per capita LKМ: 2,130,000 tons / 146,000,000 people = 14.59 kg/person
- Region market in tons: 8,698,000 × 14.59 / 1000 = ~126,948 tons
- Region market in rub: 126,948 × 77,000 = ~9.77 млрд ₽
- Actual competitive turnover: 1,195 млн ₽ = 1.195 млрд ₽
- Actual in tons: 1,195,000,000 / 77,000 = ~15,519 tons
- Gap in rub: 9,770 - 1,195 = 8,575 млн ₽
- Gap in tons: 126,948 - 15,519 = 111,429 tons
- Gap %: 87.8%
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
import os

OUT = '/home/z/my-project/download/КП_Олимп_коммерческое_предложение.docx'

# ── Data ─────────────────────────────────────────────────────────────
regions = [
    {'name': 'Новосибирская область', 'pop': 2786500, 'cities': 'Новосибирск', 'city_pop': 1633000},
    {'name': 'Кемеровская область',  'pop': 2527200, 'cities': 'Кемерово, Новокузнецк', 'city_pop': 560000 + 547000},
    {'name': 'Алтайский край',        'pop': 2099200, 'cities': 'Барнаул, Бийск', 'city_pop': 640000 + 183000},
    {'name': 'Томская область',       'pop': 1075000,  'cities': 'Томск', 'city_pop': 578000},
    {'name': 'Республика Алтай',     'pop': 210100,   'cities': 'Горно-Алтайск', 'city_pop': 64000},
]
total_pop = sum(r['pop'] for r in regions)
total_city_pop = sum(r['city_pop'] for r in regions)

# Market calculations
per_capita_kg = 14.59  # kg/person/year (2.13M tons / 146M people)
per_capita_rub = 1123   # rub/person/year
avg_price_per_kg = 77   # rub/kg (77,000 rub/ton)

market_tons = total_pop * per_capita_kg / 1000  # tons
market_rub = total_pop * per_capita_rub  # rub
market_rub_bln = market_rub / 1e9

# Actual distributor data (from user's table)
# Combined LKМ turnover of all listed distributors = 1,195 млн ₽
actual_turnover_mln = 1195  # млн ₽
actual_turnover_rub = actual_turnover_mln * 1e6  # rub
actual_turnover_tons = actual_turnover_rub / (avg_price_per_kg * 1000)  # tons (rub → tons: rub / (rub/ton))

# Gap (недозагрузка)
gap_rub_mln = market_rub / 1e6 - actual_turnover_mln  # млн ₽
gap_tons = market_tons - actual_turnover_tons
gap_pct = (1 - actual_turnover_mln / (market_rub / 1e6)) * 100

# Empils 8% share
empils_share = 0.08
empils_potential_rub = market_rub * empils_share  # rub
empils_potential_mln = empils_potential_rub / 1e6  # млн ₽
empils_potential_tons = market_tons * empils_share  # tons

akb_target = 1200
etp_covers = 120
etp_count = 4
etp_total_rtt = etp_count * etp_covers
etp_revenue = 40e6
distributor_tp_revenue = 20e6
year1 = 60e6
year2 = 120e6
year3 = 180e6
assortment = 300

# Distributor data from user's table
distributors = [
    {'name': 'Сибхозторг/Стройторг', 'portfolio': 'Рогнеда, Кушнаренко, Декарт, ЭСКАРО, НБХ, Новый дом', 'territory': 'Новосиб. обл., Алтайский край, Красноярск'},
    {'name': 'ООО «Олимп»', 'portfolio': 'АВС Фарбен, Радуга, НБХ, ЭСКАРО, ДЕКАРТ', 'territory': 'Алтайский край, Новосиб. обл., Кемерово, Томск'},
    {'name': 'ООО «АЛЬТЕРРА», 17 СМ', 'portfolio': 'Лакра, Рогнеда, Мефферт', 'territory': 'Новосиб. обл., Алтайский край'},
    {'name': 'Стройпарк 1ГМ', 'portfolio': 'Тиккурила, Рогнеда, Декарт', 'territory': 'г. Томск'},
    {'name': 'ООО «Нисса»', 'portfolio': 'Лакра, Континенталь', 'territory': 'Алтайский край'},
    {'name': 'ООО Компания «Колорлон» 3 ГМ', 'portfolio': 'Рогнеда, Лакра, Декарт, Мефферт', 'territory': 'Новосиб. обл.'},
    {'name': 'ООО «АНА»', 'portfolio': 'Тиккурила', 'territory': 'Омская обл.'},
    {'name': 'ИП Шишков С.В.', 'portfolio': 'ТД Радуга, АВС Фарбен, Рогнеда, ВИТ, НБХ', 'territory': 'Омская обл.'},
    {'name': 'ИП Миронкин Е.В.', 'portfolio': 'Лакра', 'territory': 'Омская обл.'},
    {'name': 'ИП Ковалев А.А.', 'portfolio': 'Кушнаренко, Новый Дом', 'territory': 'Омская обл.'},
]

# ── Create document ───────────────────────────────────────────────────
doc = Document()

sections = doc.sections
for section in sections:
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)

INK = RGBColor(0x2F, 0x3A, 0x48)
MUTE = RGBColor(0x8A, 0x95, 0xA3)
RED = RGBColor(0xC0, 0x00, 0x00)
SAND = RGBColor(0xC9, 0xA7, 0x7C)
BLUE = RGBColor(0x6B, 0x8F, 0xA8)
GREEN = RGBColor(0x8B, 0xAA, 0x8A)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

def add_heading(text, color=RED, size=11, space_before=10, space_after=6):
    p = doc.add_paragraph()
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.bold = True
    run.font.color.rgb = color
    run.font.name = 'Inter'
    return p

def add_body(text, bold=False, color=INK, size=9, space_after=4):
    p = doc.add_paragraph()
    p.space_after = Pt(space_after)
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = 'Inter'
    return p

def add_highlighted(text, color=RED, size=9, space_after=4):
    """Add a paragraph with colored bold text."""
    p = doc.add_paragraph()
    p.space_after = Pt(space_after)
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.bold = True
    run.font.color.rgb = color
    run.font.name = 'Inter'
    return p

def shade_cell(cell, color_hex):
    shading = cell._element.get_or_add_tcPr()
    shd = shading.makeelement(qn('w:shd'), {
        qn('w:fill'): color_hex,
        qn('w:val'): 'clear',
    })
    shading.append(shd)

def set_cell(cell, text, size=8, bold=False, color=INK, align=WD_ALIGN_PARAGRAPH.LEFT):
    cell.text = ''
    p = cell.paragraphs[0]
    p.alignment = align
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = 'Inter'

# ── Title ────────────────────────────────────────────────────────────
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run('Коммерческое предложение ЗАО «Эмпилс»')
run.font.size = Pt(16)
run.font.bold = True
run.font.color.rgb = INK
run.font.name = 'Inter'

subtitle = doc.add_paragraph()
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = subtitle.add_run('Партнёр: Олимп · Регионы: Новосибирская, Томская, Кемеровская обл., Алтайский край, Республика Алтай')
run.font.size = Pt(10)
run.font.italic = True
run.font.color.rgb = MUTE
run.font.name = 'Inter'

divider = doc.add_paragraph()
divider.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = divider.add_run('—' * 60)
run.font.color.rgb = MUTE
run.font.size = Pt(8)

# ── 1. Population ───────────────────────────────────────────────────
add_heading('1. ПОТЕНЦИАЛ РЕГИОНА — НАСЕЛЕНИЕ')

table = doc.add_table(rows=len(regions) + 2, cols=4)
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER

headers = ['Регион', 'Население (чел.)', 'Ключевые города', 'Города (чел.)']
for i, h in enumerate(headers):
    set_cell(table.rows[0].cells[i], h, size=8, bold=True, color=WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)
    shade_cell(table.rows[0].cells[i], '2F3A48')

for i, r in enumerate(regions):
    row = table.rows[i + 1]
    set_cell(row.cells[0], r['name'], size=8)
    set_cell(row.cells[1], f'{r["pop"]:,}'.replace(',', ' '), size=8, align=WD_ALIGN_PARAGRAPH.RIGHT)
    set_cell(row.cells[2], r['cities'], size=8)
    set_cell(row.cells[3], f'{r["city_pop"]:,}'.replace(',', ' '), size=8, align=WD_ALIGN_PARAGRAPH.RIGHT)
    if i % 2 == 0:
        for j in range(4):
            shade_cell(row.cells[j], 'F5F2EC')

set_cell(table.rows[-1].cells[0], 'ИТОГО', size=8, bold=True)
set_cell(table.rows[-1].cells[1], f'{total_pop:,}'.replace(',', ' '), size=8, bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
set_cell(table.rows[-1].cells[2], '7 городов', size=8, bold=True)
set_cell(table.rows[-1].cells[3], f'{total_city_pop:,}'.replace(',', ' '), size=8, bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
for j in range(4):
    shade_cell(table.rows[-1].cells[j], 'E8E3D9')

add_body(f'Население 5 регионов: {total_pop:,} чел. · Ключевые города: {total_city_pop:,} чел.'.replace(',', ' '), bold=True)

# ── 2. Market potential (in BOTH rub and tons) ──────────────────────
add_heading('2. ПОТЕНЦИАЛ РЫНКА ЛКМ РЕГИОНА')

# Compact summary table — rub + tons
market_table = doc.add_table(rows=5, cols=3)
market_table.style = 'Table Grid'
market_table.alignment = WD_TABLE_ALIGNMENT.CENTER

market_data = [
    ['Параметр', 'В рублях', 'В тоннах'],
    ['Потребление ЛКМ на душу населения РФ', f'~{per_capita_rub} ₽/чел.', f'~{per_capita_kg} кг/чел.'],
    ['Расчёт', f'{total_pop:,} × {per_capita_rub} ₽'.replace(',', ' '), f'{total_pop:,} × {per_capita_kg} кг ÷ 1000'.replace(',', ' ')],
    ['Средняя цена', f'~{avg_price_per_kg} 000 ₽/т', '—'],
    ['ПОТЕНЦИАЛ РЫНКА ЛКМ РЕГИОНА', f'{market_rub_bln:.2f} млрд ₽', f'{market_tons:,.0f} т'.replace(',', ' ')],
]

for i, row_data in enumerate(market_data):
    row = market_table.rows[i]
    for j, val in enumerate(row_data):
        set_cell(row.cells[j], val, size=8, bold=(i == 4), color=(RED if i == 4 else INK),
                 align=(WD_ALIGN_PARAGRAPH.CENTER if j > 0 else WD_ALIGN_PARAGRAPH.LEFT))
        if i == 0:
            shade_cell(row.cells[j], '2F3A48')
            set_cell(row.cells[j], val, size=8, bold=True, color=WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)
        elif i == 4:
            shade_cell(row.cells[j], 'FFF5E6')
        elif i % 2 == 1:
            shade_cell(row.cells[j], 'F5F2EC')

# ── 3. Current market (actual distributors) ──────────────────────────
add_heading('3. ТЕКУЩИЙ РЫНОК РЕГИОНА — ФАКТИЧЕСКИЕ ДИСТРИБЬЮТОРЫ')

# Compact distributor table
dist_table = doc.add_table(rows=len(distributors) + 2, cols=3)
dist_table.style = 'Table Grid'
dist_table.alignment = WD_TABLE_ALIGNMENT.CENTER

# Header
dist_headers = ['Дистрибьютор', 'Портфель конкурентов', 'Территория']
for i, h in enumerate(dist_headers):
    set_cell(dist_table.rows[0].cells[i], h, size=8, bold=True, color=WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)
    shade_cell(dist_table.rows[0].cells[i], '2F3A48')

for i, d in enumerate(distributors):
    row = dist_table.rows[i + 1]
    set_cell(row.cells[0], d['name'], size=7.5, bold=True)
    set_cell(row.cells[1], d['portfolio'], size=7.5)
    set_cell(row.cells[2], d['territory'], size=7.5)
    if i % 2 == 0:
        for j in range(3):
            shade_cell(row.cells[j], 'F5F2EC')

# Total row
set_cell(dist_table.rows[-1].cells[0], f'ИТОГО (10 дистрибьюторов)', size=8, bold=True)
set_cell(dist_table.rows[-1].cells[1], f'Совокупный ТО: {actual_turnover_mln} млн ₽', size=8, bold=True, color=RED)
set_cell(dist_table.rows[-1].cells[2], f'~{actual_turnover_tons:,.0f} т'.replace(',', ' '), size=8, bold=True, color=RED)
for j in range(3):
    shade_cell(dist_table.rows[-1].cells[j], 'FFF5E6')

# ── 3.1 Gap analysis ─────────────────────────────────────────────────
add_heading('НЕДОЗАГРУЗКА РЕГИОНА', color=RED, size=10, space_before=8, space_after=4)

gap_table = doc.add_table(rows=4, cols=4)
gap_table.style = 'Table Grid'
gap_table.alignment = WD_TABLE_ALIGNMENT.CENTER

gap_data = [
    ['Показатель', 'В рублях (млн)', 'В тоннах', 'Доля от потенциала'],
    ['Теоретический потенциал рынка', f'{market_rub/1e6:,.0f}'.replace(',', ' '), f'{market_tons:,.0f}'.replace(',', ' '), '100%'],
    ['Фактический ТО дистрибьюторов', f'{actual_turnover_mln}', f'{actual_turnover_tons:,.0f}'.replace(',', ' '), f'{actual_turnover_mln/(market_rub/1e6)*100:.1f}%'],
    ['НЕДОЗАГРУЗКА РЕГИОНА', f'{gap_rub_mln:,.0f}'.replace(',', ' '), f'{gap_tons:,.0f}'.replace(',', ' '), f'{gap_pct:.1f}%'],
]

for i, row_data in enumerate(gap_data):
    row = gap_table.rows[i]
    for j, val in enumerate(row_data):
        is_header = i == 0
        is_gap = i == 3
        set_cell(row.cells[j], val, size=8,
                 bold=(is_header or is_gap),
                 color=(WHITE if is_header else (RED if is_gap else INK)),
                 align=(WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.CENTER))
        if is_header:
            shade_cell(row.cells[j], '2F3A48')
        elif is_gap:
            shade_cell(row.cells[j], 'FFE0E0')
        elif i % 2 == 1:
            shade_cell(row.cells[j], 'F5F2EC')

add_highlighted(f'⚠ Регион недозагружен на {gap_pct:.0f}% — потенциал {gap_rub_mln:,.0f} млн ₽ ({gap_tons:,.0f} т) не освоен'.replace(',', ' '), color=RED, size=9)

# ── 4. Empils 8% ────────────────────────────────────────────────────
add_heading('4. ДОЛЯ ЭМПИЛС — 8% РЫНКА')
add_body('Среднерыночная доля Эмпилс в РФ: ~8% от рынка ЛКМ. Применительно к региону:')

emp_table = doc.add_table(rows=3, cols=3)
emp_table.style = 'Table Grid'
emp_table.alignment = WD_TABLE_ALIGNMENT.CENTER

emp_data = [
    ['Показатель', 'В рублях', 'В тоннах'],
    ['Потенциал рынка (100%)', f'{market_rub_bln:.2f} млрд ₽', f'{market_tons:,.0f} т'.replace(',', ' ')],
    ['Доля Эмпилс (8%)', f'{empils_potential_mln:.0f} млн ₽', f'{empils_potential_tons:,.0f} т'.replace(',', ' ')],
]
for i, row_data in enumerate(emp_data):
    row = emp_table.rows[i]
    for j, val in enumerate(row_data):
        is_header = i == 0
        is_empils = i == 2
        set_cell(row.cells[j], val, size=8,
                 bold=(is_header or is_empils),
                 color=(WHITE if is_header else (RED if is_empils else INK)),
                 align=(WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.CENTER))
        if is_header:
            shade_cell(row.cells[j], '2F3A48')
        elif is_empils:
            shade_cell(row.cells[j], 'FFF5E6')
        else:
            shade_cell(row.cells[j], 'F5F2EC')

add_body(f'Покрытие потенциальной АКБ партнёра: {akb_target} РТТ', bold=True)

# ── 5. ETP ──────────────────────────────────────────────────────────
add_heading('5. ОРГАНИЗАЦИЯ ПРОДАЖ — ЭКСКЛЮЗИВНЫЕ ТП')

etp_data = [
    ['Параметр', 'Значение'],
    ['1 ЭТП покрывает', f'{etp_covers} РТТ'],
    ['4 ЭТП покрывают', f'{etp_total_rtt} РТТ'],
    ['Товарооборот через ЭТП', f'{etp_revenue/1e6:.0f} млн ₽/год'],
    ['Товарооборот через ТП дистрибьютора', f'{distributor_tp_revenue/1e6:.0f} млн ₽/год'],
    ['ИТОГО год 1', f'{year1/1e6:.0f} млн ₽'],
]
etp_table = doc.add_table(rows=len(etp_data), cols=2)
etp_table.style = 'Table Grid'
etp_table.alignment = WD_TABLE_ALIGNMENT.CENTER

for i, (param, value) in enumerate(etp_data):
    row = etp_table.rows[i]
    set_cell(row.cells[0], param, size=9, bold=(i == 0 or i == len(etp_data)-1),
             color=(WHITE if i == 0 else INK))
    set_cell(row.cells[1], value, size=9, bold=(i == 0 or i == len(etp_data)-1),
             color=(WHITE if i == 0 else (RED if i == len(etp_data)-1 else INK)),
             align=WD_ALIGN_PARAGRAPH.CENTER)
    if i == 0:
        shade_cell(row.cells[0], '2F3A48'); shade_cell(row.cells[1], '2F3A48')
    elif i == len(etp_data)-1:
        shade_cell(row.cells[0], 'FFF5E6'); shade_cell(row.cells[1], 'FFF5E6')
    elif i % 2 == 0:
        shade_cell(row.cells[0], 'F5F2EC'); shade_cell(row.cells[1], 'F5F2EC')

# ── 6. Revenue forecast ─────────────────────────────────────────────
add_heading('6. ПРОГНОЗ ТОВАРООБОРОТА ПО ГОДАМ')

forecast_data = [
    ['Год', 'Товарооборот', 'Рост', 'Комментарий'],
    ['Год 1', f'{year1/1e6:.0f} млн ₽', '—', '4 ЭТП + ТП дистрибьютора, 300 SKU'],
    ['Год 2', f'{year2/1e6:.0f} млн ₽', '×2', 'Расширение АКБ до 800+ РТТ'],
    ['Год 3', f'{year3/1e6:.0f} млн ₽', '×3', 'Полное покрытие потенциала 8%'],
]
fc_table = doc.add_table(rows=len(forecast_data), cols=4)
fc_table.style = 'Table Grid'
fc_table.alignment = WD_TABLE_ALIGNMENT.CENTER

for i, row_data in enumerate(forecast_data):
    row = fc_table.rows[i]
    for j, val in enumerate(row_data):
        set_cell(row.cells[j], val, size=9, bold=(i == 0),
                 color=(WHITE if i == 0 else INK),
                 align=(WD_ALIGN_PARAGRAPH.LEFT if j == 0 or j == 3 else WD_ALIGN_PARAGRAPH.CENTER))
        if i == 0:
            shade_cell(row.cells[j], '2F3A48')
        elif i % 2 == 1:
            shade_cell(row.cells[j], 'F5F2EC')

# ── 7. Assortment ───────────────────────────────────────────────────
add_heading('7. АССОРТИМЕНТ НАЧАЛА ВЗАИМОДЕЙСТВИЯ')
add_body(f'Стартовый ассортимент: {assortment} SKU из портфеля Эмпилс', bold=True)
add_body('Бренды: ОРЕОЛ · ОРЕОЛ ДИСКОНТ · NEWTONE · AVANTGARDE · TITAN · PROFIWOOD')
add_body('Категории: краски интерьерные · фасадные · эмали · грунт-эмали · грунтовки · шпатлёвки · клеи')

# ── Logic chain ──────────────────────────────────────────────────────
divider = doc.add_paragraph()
divider.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = divider.add_run('—' * 60)
run.font.color.rgb = MUTE
run.font.size = Pt(8)

add_heading('ЛОГИЧЕСКАЯ ЦЕПОЧКА РАСЧЁТА', color=MUTE, size=10)

logic = doc.add_paragraph()
logic.alignment = WD_ALIGN_PARAGRAPH.CENTER
logic_text = (
    f'{total_pop/1e6:.1f} млн чел. → '
    f'рынок ЛКМ {market_rub_bln:.1f} млрд ₽ ({market_tons:,.0f} т) → '.replace(',', ' ') +
    f'фактически {actual_turnover_mln} млн ₽ → '
    f'недозагрузка {gap_pct:.0f}% → '
    f'доля Эмпилс 8% = {empils_potential_mln:.0f} млн ₽ → '
    f'АКБ {akb_target} РТТ → 4 ЭТП → '
    f'Г1: {year1/1e6:.0f}M → Г2: {year2/1e6:.0f}M → Г3: {year3/1e6:.0f}M ₽'
)
run = logic.add_run(logic_text)
run.font.size = Pt(8)
run.font.bold = True
run.font.italic = True
run.font.color.rgb = SAND
run.font.name = 'Inter'

doc.save(OUT)
print(f"DOCX: {OUT} ({os.path.getsize(OUT)} bytes)")
print(f"\n=== Key calculations ===")
print(f"Per capita: {per_capita_kg} kg, {per_capita_rub} rub")
print(f"Market: {market_tons:,.0f} tons, {market_rub_bln:.2f} bln rub")
print(f"Actual: {actual_turnover_mln} mln rub, {actual_turnover_tons:,.0f} tons")
print(f"Gap: {gap_rub_mln:,.0f} mln rub, {gap_tons:,.0f} tons, {gap_pct:.1f}%")
print(f"Empils 8%: {empils_potential_mln:.0f} mln rub, {empils_potential_tons:,.0f} tons")
