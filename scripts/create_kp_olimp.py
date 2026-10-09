"""
Create commercial proposal for distributor "Олимп" in 2 formats:
1. PDF (A4) — document format
2. PPTX slide — presentation format

Logic chain (top-down):
  Population → Market potential → Empils share (8%) → AKБ coverage →
  ETP coverage → Revenue forecast → Year 1/2/3
"""
from pptx import Presentation
from pptx.util import Emu, Pt, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm as rl_cm
from reportlab.lib.colors import HexColor
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.platypus.flowables import HRFlowable
import os

# ── Data ─────────────────────────────────────────────────────────────
regions = [
    {'name': 'Новосибирская область', 'pop': 2786500, 'city': 'Новосибирск', 'city_pop': 1633000},
    {'name': 'Кемеровская область',  'pop': 2527200, 'city': 'Кемерово',     'city_pop': 560000},
    {'name': 'Алтайский край',       'pop': 2099200, 'city': 'Барнаул',      'city_pop': 640000},
    {'name': 'Томская область',      'pop': 1075000,  'city': 'Томск',        'city_pop': 578000},
    {'name': 'Республика Алтай',     'pop': 210100,   'city': 'Горно-Алтайск', 'city_pop': 64000},
]
total_pop = sum(r['pop'] for r in regions)
total_city_pop = sum(r['city_pop'] for r in regions)

# Market calculation
# Russia LKМ consumption: ~10 kg per capita → ~1.18M tons for 146M people
# Average price: ~770 rub/kg → market = ~1.18M × 770 = ~909 billion rub (too high, use per-capita rub)
# Better: Russia LKМ market ~2.13M tons × ~77000 rub/ton = ~164 billion rub for 146M people
# Per capita: 164B / 146M = ~1123 rub/person/year
per_capita_lkm = 1123  # rub/year
market_potential = total_pop * per_capita_lkm  # rub
empils_share = 0.08
empils_potential = market_potential * empils_share

# AKБ
akb_target = 1200  # РТТ
# Empils share of 8% = covering 8% of market = needs ~1200 РТТ (Олимп's AKБ)

# ETP (exclusive trade representatives)
etp_covers = 120  # РТТ per ETP
etp_count = 4
etp_total_rtt = etp_count * etp_covers  # 480-500 РТТ
etp_revenue = 40e6  # 40 million rub from ETPs
distributor_tp_revenue = 20e6  # 20 million from distributor's own TP
year1 = 60e6  # 60 million
year2 = 120e6
year3 = 180e6

assortment = 300  # SKU

# ════════════════════════════════════════════════════════════════════════
# FORMAT 1: PDF (A4)
# ════════════════════════════════════════════════════════════════════════
OUT_PDF = '/home/z/my-project/download/КП_Олимп_коммерческое_предложение.pdf'

doc = SimpleDocTemplate(OUT_PDF, pagesize=A4,
                         leftMargin=2*rl_cm, rightMargin=2*rl_cm,
                         topMargin=1.5*rl_cm, bottomMargin=1.5*rl_cm)

styles = getSampleStyleSheet()
INK = HexColor('#2F3A48')
MUTE = HexColor('#8A95A3')
RED = HexColor('#C00000')
SAND = HexColor('#C9A77C')
BLUE = HexColor('#6B8FA8')
GREEN = HexColor('#8BAA8A')
BG = HexColor('#F5F2EC')
WHITE = HexColor('#FFFFFF')

# Custom styles
title_style = ParagraphStyle('Title', parent=styles['Title'],
    fontSize=16, textColor=INK, spaceAfter=6, alignment=1, fontName='Helvetica-Bold')
subtitle_style = ParagraphStyle('Subtitle', parent=styles['Normal'],
    fontSize=10, textColor=MUTE, spaceAfter=12, alignment=1, fontName='Helvetica-Oblique')
section_style = ParagraphStyle('Section', parent=styles['Heading2'],
    fontSize=11, textColor=RED, spaceBefore=10, spaceAfter=6, fontName='Helvetica-Bold')
body_style = ParagraphStyle('Body', parent=styles['Normal'],
    fontSize=9, textColor=INK, spaceAfter=4, fontName='Helvetica', leading=12)
bold_style = ParagraphStyle('Bold', parent=styles['Normal'],
    fontSize=9, textColor=INK, spaceAfter=4, fontName='Helvetica-Bold', leading=12)
kpi_style = ParagraphStyle('KPI', parent=styles['Normal'],
    fontSize=10, textColor=RED, spaceAfter=4, fontName='Courier-Bold', leading=14)

elements = []

# Title
elements.append(Paragraph('Коммерческое предложение ЗАО «Эмпилс»', title_style))
elements.append(Paragraph('Партнёру: Олимп · Регион: Новосибирская, Томская, Кемеровская обл., Алтайский край, Республика Алтай', subtitle_style))
elements.append(HRFlowable(width='100%', thickness=1, color=MUTE))
elements.append(Spacer(1, 6))

# 1. Population
elements.append(Paragraph('1. ПОТЕНЦИАЛ РЕГИОНА — НАСЕЛЕНИЕ', section_style))
pop_data = [['Регион', 'Население (чел.)', 'Крупный город', 'Город (чел.)']]
for r in regions:
    pop_data.append([r['name'], f'{r["pop"]:,}'.replace(',', ' '), r['city'], f'{r["city_pop"]:,}'.replace(',', ' ')])
pop_data.append(['ИТОГО', f'{total_pop:,}'.replace(',', ' '), '—', f'{total_city_pop:,}'.replace(',', ' ')])

pop_table = Table(pop_data, colWidths=[5*rl_cm, 3.5*rl_cm, 3.5*rl_cm, 3*rl_cm])
pop_table.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,0), INK),
    ('TEXTCOLOR', (0,0), (-1,0), WHITE),
    ('FONTSIZE', (0,0), (-1,-1), 8),
    ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
    ('ALIGN', (1,0), (1,-1), 'RIGHT'),
    ('ALIGN', (3,0), (3,-1), 'RIGHT'),
    ('ROWBACKGROUNDS', (0,1), (-1,-2), [BG, WHITE]),
    ('BACKGROUND', (0,-1), (-1,-1), HexColor('#E8E3D9')),
    ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
    ('GRID', (0,0), (-1,-1), 0.5, HexColor('#D9D2C2')),
    ('TOPPADDING', (0,0), (-1,-1), 4),
    ('BOTTOMPADDING', (0,0), (-1,-1), 4),
]))
elements.append(pop_table)
elements.append(Spacer(1, 4))
elements.append(Paragraph(f'Совокупное население 5 регионов: <b>{total_pop:,} чел.</b> · Городское население в крупных городах: <b>{total_city_pop:,} чел.</b>'.replace(',', ' '), body_style))

# 2. Market potential
elements.append(Paragraph('2. ПОТЕНЦИАЛ РЫНКА ЛКМ РЕГИОНА', section_style))
elements.append(Paragraph(
    f'Среднее потребление ЛКМ на душу населения в РФ: ~{per_capita_lkm} ₽/чел. в год '
    f'(по данным рынка: 2,13 млн т × ~77 000 ₽/т ÷ 146 млн чел.)', body_style))
elements.append(Paragraph(
    f'<b>Потенциал рынка ЛКМ региона:</b> {total_pop:,} × {per_capita_lkm} ₽ = '
    f'<font color="#C00000"><b>{market_potential/1e9:.2f} млрд ₽/год</b></font>'.replace(',', ' '), bold_style))
elements.append(Spacer(1, 4))

# 3. Empils share
elements.append(Paragraph('3. ДОЛЯ ЭМПИЛС — 8% РЫНКА', section_style))
elements.append(Paragraph(
    f'Среднерыночная доля Эмпилс в РФ: ~8% от рынка ЛКМ. '
    f'Применительно к региону:', body_style))
elements.append(Paragraph(
    f'<b>Потенциал товарооборота Эмпилс в регионе:</b> '
    f'<font color="#C00000"><b>{empils_potential/1e6:.0f} млн ₽/год</b></font> '
    f'({market_potential/1e9:.2f} млрд × 8%)', bold_style))
elements.append(Paragraph(
    f'Это и есть покрытие потенциальной АКБ партнёра на уровне <b>{akb_target} РТТ</b>.', body_style))
elements.append(Spacer(1, 4))

# 4. ETP coverage
elements.append(Paragraph('4. ОРГАНИЗАЦИЯ ПРОДАЖ — ЭКСКЛЮЗИВНЫЕ ТП', section_style))
etp_data = [
    ['Параметр', 'Значение'],
    ['1 ЭТП покрывает', f'{etp_covers} РТТ'],
    ['4 ЭТП покрывают', f'{etp_total_rtt} РТТ (~480-500)'],
    ['Товарооборот через ЭТП', f'{etp_revenue/1e6:.0f} млн ₽/год'],
    ['Товарооборот через ТП дистрибьютора', f'{distributor_tp_revenue/1e6:.0f} млн ₽/год'],
    ['ИТОГО год 1', f'{year1/1e6:.0f} млн ₽'],
]
etp_table = Table(etp_data, colWidths=[8*rl_cm, 4*rl_cm])
etp_table.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,0), INK),
    ('TEXTCOLOR', (0,0), (-1,0), WHITE),
    ('FONTSIZE', (0,0), (-1,-1), 9),
    ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
    ('ALIGN', (1,0), (1,-1), 'CENTER'),
    ('ROWBACKGROUNDS', (0,1), (-1,-2), [BG, WHITE]),
    ('BACKGROUND', (0,-1), (-1,-1), HexColor('#FFF5E6')),
    ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
    ('TEXTCOLOR', (0,-1), (-1,-1), RED),
    ('GRID', (0,0), (-1,-1), 0.5, HexColor('#D9D2C2')),
    ('TOPPADDING', (0,0), (-1,-1), 5),
    ('BOTTOMPADDING', (0,0), (-1,-1), 5),
]))
elements.append(etp_table)
elements.append(Spacer(1, 4))

# 5. Revenue forecast
elements.append(Paragraph('5. ПРОГНОЗ ТОВАРООБОРОТА ПО ГОДАМ', section_style))
forecast_data = [
    ['Год', 'Товарооборот', 'Рост', 'Комментарий'],
    ['Год 1', f'{year1/1e6:.0f} млн ₽', '—', '4 ЭТП + ТП дистрибьютора, 300 SKU'],
    ['Год 2', f'{year2/1e6:.0f} млн ₽', '×2', 'Расширение АКБ до 800+ РТТ'],
    ['Год 3', f'{year3/1e6:.0f} млн ₽', '×3', 'Полное покрытие потенциала 8%'],
]
forecast_table = Table(forecast_data, colWidths=[2*rl_cm, 3.5*rl_cm, 2*rl_cm, 5.5*rl_cm])
forecast_table.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,0), INK),
    ('TEXTCOLOR', (0,0), (-1,0), WHITE),
    ('FONTSIZE', (0,0), (-1,-1), 9),
    ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
    ('ALIGN', (1,0), (2,-1), 'CENTER'),
    ('ROWBACKGROUNDS', (0,1), (-1,-1), [BG, WHITE]),
    ('GRID', (0,0), (-1,-1), 0.5, HexColor('#D9D2C2')),
    ('TOPPADDING', (0,0), (-1,-1), 5),
    ('BOTTOMPADDING', (0,0), (-1,-1), 5),
]))
elements.append(forecast_table)
elements.append(Spacer(1, 6))

# 6. Assortment
elements.append(Paragraph('6. АССОРТИМЕНТ НАЧАЛА ВЗАИМОДЕЙСТВИЯ', section_style))
elements.append(Paragraph(
    f'Стартовый ассортимент: <b>{assortment} SKU</b> из портфеля Эмпилс<br/>'
    f'Бренды: ОРЕОЛ · ОРЕОЛ ДИСКОНТ · NEWTONE · AVANTGARDE · TITAN · PROFIWOOD<br/>'
    f'Категории: краски интерьерные · фасадные · эмали · грунт-эмали · грунтовки · шпатлёвки · клеи', body_style))

elements.append(Spacer(1, 8))
elements.append(HRFlowable(width='100%', thickness=1, color=MUTE))
elements.append(Spacer(1, 4))

# Logic chain summary
elements.append(Paragraph('ЛОГИЧЕСКАЯ ЦЕПОЧКА РАСЧЁТА', section_style))
logic_text = (
    f'5 регионов ({total_pop/1e6:.1f} млн чел.) → '
    f'рынок ЛКМ {market_potential/1e9:.1f} млрд ₽ → '
    f'доля Эмпилс 8% = {empils_potential/1e6:.0f} млн ₽ → '
    f'АКБ {akb_target} РТТ → '
    f'4 ЭТП × {etp_covers} РТТ = {etp_total_rtt} РТТ + ТП дистр. → '
    f'Год 1: {year1/1e6:.0f} млн → Год 2: {year2/1e6:.0f} млн → Год 3: {year3/1e6:.0f} млн ₽'
)
elements.append(Paragraph(logic_text, ParagraphStyle('Logic', parent=body_style,
    fontSize=8, textColor=SAND, fontName='Helvetica-BoldOblique', alignment=1)))

doc.build(elements)
print(f"PDF: {OUT_PDF} ({os.path.getsize(OUT_PDF)} bytes)")


# ════════════════════════════════════════════════════════════════════════
# FORMAT 2: PPTX slide
# ════════════════════════════════════════════════════════════════════════
OUT_PPTX = '/home/z/my-project/download/КП_Олимп_слайд.pptx'

# Create standalone presentation
prs_pptx = Presentation()
prs_pptx.slide_width = Emu(int(25.4 * 360000))  # 25.4 cm
prs_pptx.slide_height = Emu(int(14.29 * 360000))  # 14.29 cm

slide = prs_pptx.slides.add_slide(prs_pptx.slide_layouts[6])  # blank

# Palette
BG_RGB   = RGBColor(0xF5, 0xF2, 0xEC)
WHITE_RGB = RGBColor(0xFF, 0xFF, 0xFF)
INK_RGB  = RGBColor(0x2F, 0x3A, 0x48)
MUTE_RGB = RGBColor(0x8A, 0x95, 0xA3)
RED_RGB  = RGBColor(0xC0, 0x00, 0x00)
SAND_RGB = RGBColor(0xC9, 0xA7, 0x7C)
BLUE_RGB = RGBColor(0x6B, 0x8F, 0xA8)
GREEN_RGB = RGBColor(0x8B, 0xAA, 0x8A)
BORDER_RGB = RGBColor(0xD9, 0xD2, 0xC2)
DIV_RGB  = RGBColor(0xEF, 0xE9, 0xDD)

def cm(v):
    return Emu(int(v * 360000))

def add_text(slide, x, y, w, h, text, size=11, bold=False, color=INK_RGB, align=PP_ALIGN.LEFT, va=MSO_ANCHOR.TOP, font="Inter", spacing=None):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    tf.vertical_anchor = va
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = text
    r.font.name = font; r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color
    if spacing:
        rPr = r._r.get_or_add_rPr(); rPr.set('spc', str(spacing))
    return tb

def add_rect(slide, x, y, w, h, fill, line=None, line_w=6350, corner_radius=None):
    shp_type = MSO_SHAPE.ROUNDED_RECTANGLE if corner_radius else MSO_SHAPE.RECTANGLE
    s = slide.shapes.add_shape(shp_type, x, y, w, h)
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line: s.line.color.rgb = line; s.line.width = Emu(line_w)
    else: s.line.fill.background()
    if corner_radius is not None:
        try: s.adjustments[0] = corner_radius
        except: pass
    s.shadow.inherit = False
    return s

# Background
bg = slide.background; bg.fill.solid(); bg.fill.fore_color.rgb = BG_RGB

# Header
add_text(slide, cm(0.5), cm(0.4), cm(20), cm(0.7),
         "Коммерческое предложение Эмпилс · Партнёр: Олимп",
         size=18, bold=True, color=INK_RGB)
add_text(slide, cm(0.5), cm(1.1), cm(20), cm(0.4),
         "Новосибирская · Томская · Кемеровская обл. · Алтайский край · Республика Алтай",
         size=10, color=MUTE_RGB)
# Pill
add_rect(slide, cm(20.0), cm(0.35), cm(5.0), cm(0.7), WHITE_RGB, line=RED_RGB, line_w=9525, corner_radius=0.5)
add_text(slide, cm(20.0), cm(0.45), cm(5.0), cm(0.55),
         "КП · 2026", size=10, bold=True, color=RED_RGB, align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE, spacing=150)
# Divider
add_rect(slide, cm(0.5), cm(1.7), cm(24.5), cm(0.06), BORDER_RGB)

# ── LEFT: Logic chain (steps 1-6 as vertical flow) ──────────────────
LEFT_X = cm(0.5); LEFT_Y = cm(2.0); LEFT_W = cm(12.0)

steps = [
    {"n": "1", "title": "ПОТЕНЦИАЛ РЕГИОНА", "value": f"{total_pop/1e6:.1f} млн чел.",
     "detail": f"5 регионов: Новосибирская ({regions[0]['pop']/1e6:.2f}M), Кемеровская ({regions[1]['pop']/1e6:.2f}M), Алтайский край ({regions[2]['pop']/1e6:.2f}M), Томская ({regions[3]['pop']/1e6:.2f}M), Респ. Алтай ({regions[4]['pop']/1e3:.0f}K)",
     "accent": MUTE_RGB},
    {"n": "2", "title": "РЫНОК ЛКМ РЕГИОНА", "value": f"{market_potential/1e9:.1f} млрд ₽",
     "detail": f"~{per_capita_lkm} ₽/чел. × {total_pop/1e6:.1f} млн = потенциал рынка",
     "accent": BLUE_RGB},
    {"n": "3", "title": "ДОЛЯ ЭМПИЛС (8%)", "value": f"{empils_potential/1e6:.0f} млн ₽",
     "detail": f"Средняя доля Эмпилс в РФ · Покрытие АКБ {akb_target} РТТ",
     "accent": RED_RGB},
    {"n": "4", "title": "ОРГАНИЗАЦИЯ ПРОДАЖ", "value": f"4 ЭТП = {etp_total_rtt} РТТ",
     "detail": f"1 ЭТП = {etp_covers} РТТ · 4 ЭТП = {etp_revenue/1e6:.0f} млн ₽ + ТП дистр. {distributor_tp_revenue/1e6:.0f} млн ₽",
     "accent": SAND_RGB},
    {"n": "5", "title": "ПРОГНОЗ ТОВАРООБОРОТА", "value": f"Г1: {year1/1e6:.0f} → Г3: {year3/1e6:.0f} млн ₽",
     "detail": f"Год 1: {year1/1e6:.0f}M · Год 2: {year2/1e6:.0f}M · Год 3: {year3/1e6:.0f}M",
     "accent": GREEN_RGB},
    {"n": "6", "title": "АССОРТИМЕНТ", "value": f"{assortment} SKU",
     "detail": "ОРЕОЛ · NEWTONE · AVANTGARDE · TITAN · PROFIWOOD",
     "accent": INK_RGB},
]

step_h = cm(1.55)
step_gap = cm(0.15)
for i, s in enumerate(steps):
    sy = LEFT_Y + i * (step_h + step_gap)
    # Card
    add_rect(slide, LEFT_X, sy, LEFT_W, step_h, WHITE_RGB, line=BORDER_RGB, line_w=6350, corner_radius=0.03)
    # Number
    add_rect(slide, LEFT_X, sy, cm(0.8), step_h, s['accent'], corner_radius=0.03)
    add_text(slide, LEFT_X, sy + cm(0.35), cm(0.8), cm(0.8),
             s['n'], size=18, bold=True, color=WHITE_RGB, font="Roboto Mono",
             align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)
    # Title
    add_text(slide, LEFT_X + cm(1.0), sy + cm(0.15), cm(10.5), cm(0.35),
             s['title'], size=9, bold=True, color=INK_RGB, font="Inter", spacing=100)
    # Value (big)
    add_text(slide, LEFT_X + cm(1.0), sy + cm(0.5), cm(10.5), cm(0.45),
             s['value'], size=14, bold=True, color=s['accent'], font="Roboto Mono")
    # Detail
    add_text(slide, LEFT_X + cm(1.0), sy + cm(1.0), cm(10.5), cm(0.4),
             s['detail'], size=8, color=MUTE_RGB, font="Inter")

# ── RIGHT: Revenue forecast chart + summary ───────────────────────────
RIGHT_X = cm(13.0); RIGHT_Y = cm(2.0); RIGHT_W = cm(12.0)

# Revenue forecast card
add_rect(slide, RIGHT_X, RIGHT_Y, RIGHT_W, cm(5.5), WHITE_RGB, line=BORDER_RGB, line_w=6350, corner_radius=0.03)
add_text(slide, RIGHT_X + cm(0.3), RIGHT_Y + cm(0.2), cm(11.0), cm(0.4),
         "ПРОГНОЗ ТОВАРООБОРОТА ПО ГОДАМ", size=10, bold=True, color=RED_RGB, font="Inter", spacing=200)
add_rect(slide, RIGHT_X + cm(0.3), RIGHT_Y + cm(0.65), cm(11.4), cm(0.04), DIV_RGB)

# 3 year bars (visual)
year_data = [
    {"year": "Год 1", "value": year1, "color": MUTE_RGB, "comment": "4 ЭТП + ТП дистр.\n300 SKU старт"},
    {"year": "Год 2", "value": year2, "color": SAND_RGB, "comment": "Расширение АКБ\nдо 800+ РТТ"},
    {"year": "Год 3", "value": year3, "color": RED_RGB,  "comment": "Полное покрытие\nпотенциала 8%"},
]

bar_y = RIGHT_Y + cm(1.0)
bar_max_h = cm(3.5)
bar_w = cm(3.0)
bar_gap = cm(0.7)
bar_x0 = RIGHT_X + cm(1.0)
max_val = year3

for i, yd in enumerate(year_data):
    bx = bar_x0 + i * (bar_w + bar_gap)
    bh = int(bar_max_h * (yd['value'] / max_val))
    # Bar
    add_rect(slide, bx, bar_y + bar_max_h - bh, bar_w, bh, yd['color'], corner_radius=0.05)
    # Value on top
    add_text(slide, bx, bar_y + bar_max_h - bh - cm(0.45), bar_w, cm(0.4),
             f"{yd['value']/1e6:.0f} млн ₽", size=12, bold=True, color=yd['color'],
             font="Roboto Mono", align=PP_ALIGN.CENTER)
    # Year label below
    add_text(slide, bx, bar_y + bar_max_h + cm(0.05), bar_w, cm(0.3),
             yd['year'], size=10, bold=True, color=INK_RGB, align=PP_ALIGN.CENTER, font="Inter")
    # Comment below year
    add_text(slide, bx - cm(0.2), bar_y + bar_max_h + cm(0.35), bar_w + cm(0.4), cm(0.7),
             yd['comment'], size=7.5, color=MUTE_RGB, align=PP_ALIGN.CENTER, font="Inter")

# KPI summary card below
KPI_Y = RIGHT_Y + cm(6.0)
add_rect(slide, RIGHT_X, KPI_Y, RIGHT_W, cm(5.5), WHITE_RGB, line=BORDER_RGB, line_w=6350, corner_radius=0.03)
add_text(slide, RIGHT_X + cm(0.3), KPI_Y + cm(0.2), cm(11.0), cm(0.4),
         "КЛЮЧЕВЫЕ ПАРАМЕТРЫ", size=10, bold=True, color=RED_RGB, font="Inter", spacing=200)
add_rect(slide, RIGHT_X + cm(0.3), KPI_Y + cm(0.65), cm(11.4), cm(0.04), DIV_RGB)

kpis = [
    ("Население 5 регионов", f"{total_pop/1e6:.1f} млн чел.", MUTE_RGB),
    ("Рынок ЛКМ региона", f"{market_potential/1e9:.1f} млрд ₽", BLUE_RGB),
    ("Доля Эмпилс (8%)", f"{empils_potential/1e6:.0f} млн ₽", RED_RGB),
    ("АКБ партнёра", f"{akb_target} РТТ", SAND_RGB),
    ("ЭТП покрытие", f"4 × {etp_covers} = {etp_total_rtt} РТТ", GREEN_RGB),
    ("Ассортимент", f"{assortment} SKU", INK_RGB),
]

for i, (label, value, color) in enumerate(kpis):
    col = i % 2
    row = i // 2
    kx = RIGHT_X + cm(0.3) + col * cm(5.8)
    ky = KPI_Y + cm(1.0) + row * cm(1.4)
    # Accent bar
    add_rect(slide, kx, ky, cm(0.06), cm(1.1), color)
    # Label
    add_text(slide, kx + cm(0.2), ky + cm(0.1), cm(5.0), cm(0.3),
             label, size=8, color=MUTE_RGB, font="Inter")
    # Value
    add_text(slide, kx + cm(0.2), ky + cm(0.4), cm(5.0), cm(0.5),
             value, size=13, bold=True, color=color, font="Roboto Mono")

# Bottom strip
STRIP_Y = cm(13.5)
add_rect(slide, cm(0.5), STRIP_Y, cm(24.5), cm(0.7), INK_RGB, corner_radius=0.03)
add_text(slide, cm(0.8), STRIP_Y + cm(0.12), cm(24.0), cm(0.5),
         f"{total_pop/1e6:.1f} млн чел. → {market_potential/1e9:.1f} млрд ₽ → 8% = {empils_potential/1e6:.0f} млн ₽ → {akb_target} РТТ → 4 ЭТП → Г1: {year1/1e6:.0f}M → Г2: {year2/1e6:.0f}M → Г3: {year3/1e6:.0f}M ₽",
         size=8, bold=True, color=SAND_RGB, font="Roboto Mono", align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)

prs_pptx.save(OUT_PPTX)
print(f"PPTX: {OUT_PPTX} ({os.path.getsize(OUT_PPTX)} bytes)")
