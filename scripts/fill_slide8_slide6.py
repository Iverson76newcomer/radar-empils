"""
Two tasks:
1. Slide 8 (Запуск новинок) — fill LEFT half with analytics for 5 new product launches in 2026
   Data source: Новинки за 2026 год первичка и вторичка.xlsx
   5 новинки: PROFIWOOD антисептик, Ореол OSB, NEWTONE ABSOLUTE BLACK, NEWTONE SKY, Ореол резиновая

2. Slide 6 (currently empty) — create content about email marketing concept:
   "Касание брендом" — periodic email campaigns to potential distributors
   working with competitors build subconscious brand awareness.
"""
import openpyxl
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import shutil
import re

SRC = "/home/z/my-project/upload/Презентация_для_ГД_final.pptx"
OUT = "/home/z/my-project/download/Презентация_для_ГД_final.pptx"
EXCEL_SRC = "/home/z/my-project/upload/Новинки за 2026 год первичка и вторичка.xlsx"

shutil.copy(SRC, OUT)

# ── Palette ──────────────────────────────────────────────────────────
BG          = RGBColor(0xF5, 0xF2, 0xEC)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
CARD_BORDER = RGBColor(0xD9, 0xD2, 0xC2)
DIVIDER     = RGBColor(0xEF, 0xE9, 0xDD)
INK         = RGBColor(0x2F, 0x3A, 0x48)
MUTE        = RGBColor(0x8A, 0x95, 0xA3)
RED         = RGBColor(0xC0, 0x00, 0x00)
SAND        = RGBColor(0xC9, 0xA7, 0x7C)
BLUE        = RGBColor(0x6B, 0x8F, 0xA8)
GREEN       = RGBColor(0x8B, 0xAA, 0x8A)
DARK_BG     = RGBColor(0x2F, 0x3A, 0x48)
ACCENT_PURPLE = RGBColor(0x7B, 0x6F, 0xA8)  # for email concept

def cm(v):
    return Emu(int(v * 360000))

def add_text(slide, x, y, w, h, text, *,
             size=11, bold=False, color=INK, align=PP_ALIGN.LEFT,
             va=MSO_ANCHOR.TOP, font="Inter", spacing=None):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = 0; tf.margin_right = 0
    tf.margin_top = 0; tf.margin_bottom = 0
    tf.vertical_anchor = va
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    if spacing:
        rPr = r._r.get_or_add_rPr()
        rPr.set('spc', str(spacing))
    return tb

def add_rect(slide, x, y, w, h, fill, *,
             line=None, line_w=6350, corner_radius=None):
    shp_type = MSO_SHAPE.ROUNDED_RECTANGLE if corner_radius else MSO_SHAPE.RECTANGLE
    s = slide.shapes.add_shape(shp_type, x, y, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line:
        s.line.color.rgb = line
        s.line.width = Emu(line_w)
    else:
        s.line.fill.background()
    if corner_radius is not None:
        try: s.adjustments[0] = corner_radius
        except: pass
    s.shadow.inherit = False
    return s


# ════════════════════════════════════════════════════════════════════════
# STEP 1: Read Excel data and aggregate by новинка
# ════════════════════════════════════════════════════════════════════════

wb_src = openpyxl.load_workbook(EXCEL_SRC, data_only=True)
ws = wb_src['Первичные и вторичные продажи']

# Group rows by novinka number (column A = №, 1-5)
novinki = {}  # {1: {'name': ..., 'rows': [...]}, ...}

current_n = None
current_data = None
for row in ws.iter_rows(min_row=3, values_only=True):
    if not row or not any(row):
        continue
    n_val = row[0]
    name = row[4] if len(row) > 4 else None
    series = row[2] if len(row) > 2 else None
    p1_kg = row[8] if len(row) > 8 else 0
    p1_rub = row[9] if len(row) > 9 else 0
    p2_kg = row[10] if len(row) > 10 else 0
    p2_rub = row[11] if len(row) > 11 else 0
    akb = row[12] if len(row) > 12 else 0

    # Skip subtotal rows (where series is None and name is None)
    if series is None and name is None:
        continue

    if n_val is not None:
        # New novinka starts here
        current_n = n_val
        # Identify novinka short name from the first SKU
        # Use series + type to create short name
        sku_text = str(name) if name else ''
        # Extract a short name based on patterns
        if 'PROFIWOOD' in sku_text and 'Антисептик' in sku_text:
            short_name = 'PROFIWOOD Антисептик кроющий ВД'
            nov_type = 'Антисептик кроющий ВД'
        elif 'OSB' in sku_text:
            short_name = 'Ореол Краска для плит OSB'
            nov_type = 'Краска для OSB'
        elif 'ABSOLUTE BLACK' in sku_text:
            short_name = 'NEWTONE ABSOLUTE BLACK'
            nov_type = 'Интерьерная моющаяся'
        elif 'SKY' in sku_text:
            short_name = 'NEWTONE SKY потолочная'
            nov_type = 'Потолочная'
        elif 'резиновая' in sku_text.lower():
            short_name = 'Ореол Краска резиновая ВД'
            nov_type = 'Резиновая'
        else:
            short_name = sku_text[:60]
            nov_type = '—'

        # Extract color from SKU name
        # Get volume from SKU name
        vol_match = re.search(r'(\d+(?:\.\d+)?)\s*(л|кг)', sku_text, re.IGNORECASE)
        volume = vol_match.group(0) if vol_match else ''

        current_data = {
            'short_name': short_name,
            'type': nov_type,
            'series': str(series) if series else '',
            'skus': [],
            'total_p1_kg': 0,
            'total_p1_rub': 0,
            'total_p2_kg': 0,
            'total_p2_rub': 0,
            'total_akb': 0,
            'sku_count': 0,
        }
        novinki[current_n] = current_data

    if current_data is not None and name:
        # Parse color from name
        name_lower = sku_text.lower() if 'sku_text' in dir() else str(name).lower()
        name_str = str(name)
        # Try to identify color
        color = ''
        color_keywords = {
            'бежевый': 'бежевый', 'серый туман': 'серый туман',
            'скандинавский серый': 'сканд. серый',
            'горький шоколад': 'горький шоколад', 'шоколадный мусс': 'шоколадный мусс',
            'белый': 'белый', 'база а': 'база А', 'база с': 'база С',
            'черная': 'чёрный', 'чёрная': 'чёрный',
            'ультрабелая': 'ультрабелый',
        }
        for kw, c in color_keywords.items():
            if kw in name_str.lower():
                color = c
                break

        # Extract volume
        vol_match = re.search(r'(\d+(?:\.\d+)?)\s*(л|кг)', name_str, re.IGNORECASE)
        volume = vol_match.group(0) if vol_match else ''

        current_data['skus'].append({
            'name': name_str,
            'color': color,
            'volume': volume,
            'p1_kg': float(p1_kg) if p1_kg else 0,
            'p1_rub': float(p1_rub) if p1_rub else 0,
            'p2_kg': float(p2_kg) if p2_kg else 0,
            'p2_rub': float(p2_rub) if p2_rub else 0,
            'akb': int(akb) if akb else 0,
        })
        current_data['total_p1_kg'] += float(p1_kg) if p1_kg else 0
        current_data['total_p1_rub'] += float(p1_rub) if p1_rub else 0
        current_data['total_p2_kg'] += float(p2_kg) if p2_kg else 0
        current_data['total_p2_rub'] += float(p2_rub) if p2_rub else 0
        current_data['total_akb'] += int(akb) if akb else 0
        current_data['sku_count'] += 1

# Print summary
print("=== 5 НОВИНОК 2026 ===")
for n, d in novinki.items():
    colors = set(s['color'] for s in d['skus'] if s['color'])
    volumes = set(s['volume'] for s in d['skus'] if s['volume'])
    print(f"\nНовинка #{n}: {d['short_name']}")
    print(f"  SKU: {d['sku_count']}")
    print(f"  Цвета: {sorted(colors)}")
    print(f"  Фасовки: {sorted(volumes)}")
    print(f"  Первичка: {d['total_p1_kg']:.1f} кг / {d['total_p1_rub']:,.0f} ₽")
    print(f"  Вторичка: {d['total_p2_kg']:.1f} кг / {d['total_p2_rub']:,.0f} ₽")
    print(f"  АКБ: {d['total_akb']}")

# Grand totals
total_p1_rub = sum(d['total_p1_rub'] for d in novinki.values())
total_p2_rub = sum(d['total_p2_rub'] for d in novinki.values())
total_skus = sum(d['sku_count'] for d in novinki.values())
total_akb = sum(d['total_akb'] for d in novinki.values())
print(f"\n=== ИТОГО ===")
print(f"  SKU: {total_skus}, Первичка: {total_p1_rub:,.0f} ₽, Вторичка: {total_p2_rub:,.0f} ₽, АКБ: {total_akb}")


# ════════════════════════════════════════════════════════════════════════
# STEP 2: Open presentation and modify slide 8 (Запуск новинок)
# ════════════════════════════════════════════════════════════════════════

prs = Presentation(OUT)

# Find slide 8 (Запуск новинок)
target_idx = None
for i, slide in enumerate(prs.slides):
    for shape in slide.shapes:
        if shape.has_text_frame and 'Запуск новинок' in shape.text_frame.text:
            target_idx = i
            break
    if target_idx is not None:
        break

print(f"\nSlide 'Запуск новинок' at index {target_idx}")
slide8 = list(prs.slides)[target_idx]

# Remove shapes in the LEFT half (x < 12 cm, y > 2.5 cm) — the empty placeholder
sp_tree = slide8.shapes._spTree
shapes_to_remove = []
for elem in sp_tree:
    if not (elem.tag.endswith('}sp') or elem.tag.endswith('}pic') or elem.tag.endswith('}grpSp')):
        continue
    # Get offset
    off = elem.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}off')
    if off is None:
        continue
    x_emu = int(off.get('x', 0))
    y_emu = int(off.get('y', 0))
    x_cm = x_emu / 360000
    y_cm = y_emu / 360000
    # Left half: x < 12 cm, y > 2.5 cm (skip header shapes which are full-width)
    if x_cm < 12.0 and y_cm > 2.5:
        shapes_to_remove.append(elem)
    # Also remove the bottom divider if it's full-width at y=13.5
    # (we'll re-add it later if needed)

print(f"Removing {len(shapes_to_remove)} placeholder shapes from left half of slide 8")
for elem in shapes_to_remove:
    sp_tree.remove(elem)


# ════════════════════════════════════════════════════════════════════════
# STEP 3: Fill LEFT half of slide 8 with 5 новинок analytics
# Left card area: x=0.5, y=2.67, w=11.7, h=10.72
# ════════════════════════════════════════════════════════════════════════

LEFT_X = cm(0.5)
LEFT_Y = cm(2.67)
LEFT_W = cm(11.7)
LEFT_H = cm(10.72)

# Card background
add_rect(slide8, LEFT_X, LEFT_Y, LEFT_W, LEFT_H,
         WHITE, line=CARD_BORDER, line_w=6350, corner_radius=0.02)

# Header
add_text(slide8, LEFT_X + cm(0.4), LEFT_Y + cm(0.35),
         cm(10.9), cm(0.45),
         "РЕЗУЛЬТАТЫ ЗАПУСКА НОВИНОК 2026",
         size=11, bold=True, color=RED, font="Inter", spacing=200)
add_text(slide8, LEFT_X + cm(0.4), LEFT_Y + cm(0.85),
         cm(10.9), cm(0.35),
         f"5 новинок · {total_skus} SKU · АКБ {total_akb}",
         size=10, color=MUTE, font="Inter")
# Divider
add_rect(slide8, LEFT_X + cm(0.4), LEFT_Y + cm(1.3),
         cm(10.9), cm(0.06), DIVIDER)

# Total summary at the top of analytics
add_text(slide8, LEFT_X + cm(0.4), LEFT_Y + cm(1.5),
         cm(10.9), cm(0.4),
         f"Первичка: {total_p1_rub/1e6:.2f} млн ₽   ·   Вторичка: {total_p2_rub/1e6:.2f} млн ₽",
         size=10, bold=True, color=INK, font="Roboto Mono")
# Divider
add_rect(slide8, LEFT_X + cm(0.4), LEFT_Y + cm(1.95),
         cm(10.9), cm(0.04), DIVIDER)

# 5 новинки compact rows
# Each row: 1.6 cm tall, 5 rows = 8.0 cm (fits in 10.72 - 2.0 = 8.72 cm available)
ROW_H = cm(1.65)
ROW_Y0 = LEFT_Y + cm(2.15)

NOVINKI_ACCENTS = [RED, SAND, BLUE, GREEN, ACCENT_PURPLE]

for i, (n, d) in enumerate(sorted(novinki.items())):
    row_y = ROW_Y0 + i * ROW_H
    accent = NOVINKI_ACCENTS[i % 5]

    # Row background (alternating)
    if i % 2 == 0:
        add_rect(slide8, LEFT_X + cm(0.4), row_y,
                 cm(10.9), ROW_H - cm(0.05), BG, corner_radius=0.03)
    # Accent vertical bar
    add_rect(slide8, LEFT_X + cm(0.4), row_y,
             cm(0.08), ROW_H - cm(0.05), accent)

    # Novinka number badge
    add_text(slide8, LEFT_X + cm(0.6), row_y + cm(0.1),
             cm(0.7), cm(0.4),
             f"{n:02d}", size=11, bold=True, color=accent,
             font="Roboto Mono")

    # Novinka name (compact)
    name_short = d['short_name']
    if len(name_short) > 35:
        name_short = name_short[:33] + '…'
    add_text(slide8, LEFT_X + cm(1.35), row_y + cm(0.1),
             cm(9.5), cm(0.4),
             name_short, size=9.5, bold=True, color=INK,
             font="Inter")

    # Stats row 1: SKU count + colors + фасовки
    colors_str = ', '.join(sorted(set(s['color'] for s in d['skus'] if s['color'])))[:40]
    volumes_str = ', '.join(sorted(set(s['volume'] for s in d['skus'] if s['volume'])))

    # SKU + фасовки line
    add_text(slide8, LEFT_X + cm(1.35), row_y + cm(0.5),
             cm(9.5), cm(0.3),
             f"SKU: {d['sku_count']}  ·  Фасовки: {volumes_str}",
             size=8.5, color=MUTE, font="Inter")

    # Colors line
    if colors_str:
        add_text(slide8, LEFT_X + cm(1.35), row_y + cm(0.8),
                 cm(9.5), cm(0.3),
                 f"Цвета: {colors_str}",
                 size=8.5, color=MUTE, font="Inter")

    # KPI metrics on the right side of the row
    add_text(slide8, LEFT_X + cm(7.5), row_y + cm(0.1),
             cm(3.8), cm(0.4),
             f"{d['total_p1_rub']/1e3:.0f} тыс ₽",
             size=10, bold=True, color=accent,
             font="Roboto Mono", align=PP_ALIGN.RIGHT)
    add_text(slide8, LEFT_X + cm(7.5), row_y + cm(0.5),
             cm(3.8), cm(0.3),
             f"Первичка · {d['total_p1_kg']:.0f} кг",
             size=8, color=MUTE, font="Inter", align=PP_ALIGN.RIGHT)
    add_text(slide8, LEFT_X + cm(7.5), row_y + cm(0.8),
             cm(3.8), cm(0.3),
             f"АКБ: {d['total_akb']}  ·  SKU: {d['sku_count']}",
             size=8, color=MUTE, font="Inter", align=PP_ALIGN.RIGHT)

    # Row divider
    if i < 4:
        add_rect(slide8, LEFT_X + cm(0.4), row_y + ROW_H - cm(0.1),
                 cm(10.9), cm(0.02), DIVIDER)

# Bottom note inside left card
NOTE_Y = ROW_Y0 + 5 * ROW_H + cm(0.15)
add_rect(slide8, LEFT_X + cm(0.4), NOTE_Y,
         cm(10.9), cm(0.04), DIVIDER)
add_text(slide8, LEFT_X + cm(0.4), NOTE_Y + cm(0.15),
         cm(10.9), cm(0.5),
         "Период: 01.01.2026 — 06.10.2026 · Источник: 1С",
         size=8, color=MUTE, font="Inter", align=PP_ALIGN.CENTER)

print(f"Slide 8 left half filled with 5 новинки analytics")


# ════════════════════════════════════════════════════════════════════════
# STEP 4: Build slide 6 — Email marketing concept "Касание брендом"
# ════════════════════════════════════════════════════════════════════════

# Find slide 6 (currently empty, just 2 shapes)
slide6_idx = 5  # 0-indexed
slide6 = list(prs.slides)[slide6_idx]

# Remove all existing shapes from slide 6 (it's mostly empty)
sp_tree6 = slide6.shapes._spTree
shapes_to_remove6 = []
for elem in sp_tree6:
    if elem.tag.endswith('}sp') or elem.tag.endswith('}pic') or elem.tag.endswith('}grpSp') or elem.tag.endswith('}cxnSp'):
        shapes_to_remove6.append(elem)
print(f"\nSlide 6: removing {len(shapes_to_remove6)} existing shapes")
for elem in shapes_to_remove6:
    sp_tree6.remove(elem)

# Set background
bg = slide6.background
bg.fill.solid()
bg.fill.fore_color.rgb = BG

# ── HEADER ────────────────────────────────────────────────────────────
add_text(slide6, cm(0.5), cm(0.5),
         cm(20.0), cm(0.95),
         "Email-рассылка · Касание брендом",
         size=20, bold=True, color=INK, font="Inter")
add_text(slide6, cm(0.5), cm(1.55),
         cm(20.0), cm(0.5),
         "Стратегия формирования узнаваемости бренда Эмпилс у потенциальных партнёров",
         size=11, color=MUTE, font="Inter")
# Status pill
add_rect(slide6, cm(20.0), cm(0.45),
         cm(5.0), cm(0.78),
         WHITE, line=SAND, line_w=9525, corner_radius=0.5)
add_text(slide6, cm(20.0), cm(0.58),
         cm(5.0), cm(0.65),
         "СТРАТЕГИЯ · 2026", size=10, bold=True,
         color=SAND, font="Inter", align=PP_ALIGN.CENTER, spacing=150)
# Divider
add_rect(slide6, cm(0.5), cm(2.22),
         cm(24.5), cm(0.08), CARD_BORDER)


# ════════════════════════════════════════════════════════════════════════
# LEFT HALF (5.0 cm to 12.5 cm) — Concept of "Касание брендом"
# ════════════════════════════════════════════════════════════════════════
LEFT_X = cm(0.5)
LEFT_Y = cm(2.67)
LEFT_W = cm(11.7)
LEFT_H = cm(10.72)

# Card background
add_rect(slide6, LEFT_X, LEFT_Y, LEFT_W, LEFT_H,
         WHITE, line=CARD_BORDER, line_w=6350, corner_radius=0.02)

# Header
add_text(slide6, LEFT_X + cm(0.4), LEFT_Y + cm(0.35),
         cm(10.9), cm(0.45),
         "КОНЦЕПЦИЯ «КАСАНИЕ БРЕНДОМ»",
         size=11, bold=True, color=RED, font="Inter", spacing=200)
add_text(slide6, LEFT_X + cm(0.4), LEFT_Y + cm(0.85),
         cm(10.9), cm(0.35),
         "Формирование подсознательной узнаваемости",
         size=10, color=MUTE, font="Inter")
add_rect(slide6, LEFT_X + cm(0.4), LEFT_Y + cm(1.3),
         cm(10.9), cm(0.06), DIVIDER)

# Main concept text block 1
add_text(slide6, LEFT_X + cm(0.4), LEFT_Y + cm(1.55),
         cm(10.9), cm(0.4),
         "Суть подхода",
         size=10, bold=True, color=INK, font="Inter", spacing=100)
add_text(slide6, LEFT_X + cm(0.4), LEFT_Y + cm(2.0),
         cm(10.9), cm(1.6),
         "Email-рассылка по потенциальным партнёрам, работающим с конкурентами — это не спам и не прямой продажи. Это регулярное касание потенциального дистрибьютора с брендом Эмпилс. Каждое письмо — точка контакта, которая формирует информационное присутствие бренда в сознании партнёра.",
         size=9.5, color=INK, font="Inter", va=MSO_ANCHOR.TOP)

# Divider
add_rect(slide6, LEFT_X + cm(0.4), LEFT_Y + cm(3.7),
         cm(10.9), cm(0.04), DIVIDER)

# Mechanism — how it works
add_text(slide6, LEFT_X + cm(0.4), LEFT_Y + cm(3.85),
         cm(10.9), cm(0.4),
         "Как это работает",
         size=10, bold=True, color=INK, font="Inter", spacing=100)

# Mechanism bullets
mechanism_items = [
    ("1.", "Касание", "Партнёр получает email от Эмпилс — первое знакомство с брендом"),
    ("2.", "Регулярность", "Периодичность рассылки создаёт узнаваемый ритм присутствия"),
    ("3.", "Накопление", "Каждое касание усиливает информационный след бренда"),
    ("4.", "Узнаваемость", "На подсознательном уровне формируется знание об Эмпилс"),
    ("5.", "Решение", "Когда возникает потребность — партнёр вспоминает об Эмпилс"),
]

mech_y = LEFT_Y + cm(4.35)
for num, title, desc in mechanism_items:
    # Number
    add_text(slide6, LEFT_X + cm(0.4), mech_y,
             cm(0.5), cm(0.35),
             num, size=11, bold=True, color=ACCENT_PURPLE,
             font="Roboto Mono")
    # Title
    add_text(slide6, LEFT_X + cm(1.0), mech_y,
             cm(2.5), cm(0.35),
             title, size=10, bold=True, color=INK,
             font="Inter")
    # Description
    add_text(slide6, LEFT_X + cm(3.5), mech_y,
             cm(7.8), cm(0.35),
             desc, size=9, color=MUTE,
             font="Inter")
    mech_y += cm(0.5)

# Divider
add_rect(slide6, LEFT_X + cm(0.4), LEFT_Y + cm(7.0),
         cm(10.9), cm(0.04), DIVIDER)

# Key principle quote
add_text(slide6, LEFT_X + cm(0.4), LEFT_Y + cm(7.2),
         cm(10.9), cm(0.4),
         "Ключевой принцип",
         size=10, bold=True, color=RED, font="Inter", spacing=100)
# Quote
add_rect(slide6, LEFT_X + cm(0.4), LEFT_Y + cm(7.65),
         cm(0.08), cm(2.5), ACCENT_PURPLE)
add_text(slide6, LEFT_X + cm(0.7), LEFT_Y + cm(7.7),
         cm(10.5), cm(2.4),
         "«Рано или поздно периодичность касания формирует подсознательное знание о бренде. Когда партнёр, работающий с конкурентом, столкнётся с потребностью в альтернативе — он вспомнит об Эмпилс. Не потому что мы ему навязали, а потому что мы были рядом.»",
         size=10, color=INK, font="Inter",
         va=MSO_ANCHOR.TOP)

# Bottom note
add_rect(slide6, LEFT_X + cm(0.4), LEFT_Y + cm(10.25),
         cm(10.9), cm(0.04), DIVIDER)
add_text(slide6, LEFT_X + cm(0.4), LEFT_Y + cm(10.4),
         cm(10.9), cm(0.3),
         "Цель: системное формирование информационного присутствия бренда",
         size=8.5, color=MUTE, font="Inter", align=PP_ALIGN.CENTER)


# ════════════════════════════════════════════════════════════════════════
# RIGHT HALF — Visual scheme of "Touch → Awareness"
# ════════════════════════════════════════════════════════════════════════
RIGHT_X = cm(13.0)
RIGHT_Y = cm(2.67)
RIGHT_W = cm(11.7)
RIGHT_H = cm(10.72)

add_rect(slide6, RIGHT_X, RIGHT_Y, RIGHT_W, RIGHT_H,
         WHITE, line=CARD_BORDER, line_w=6350, corner_radius=0.02)

# Header
add_text(slide6, RIGHT_X + cm(0.4), RIGHT_Y + cm(0.35),
         cm(10.9), cm(0.45),
         "КАСКАД КАСАНИЙ",
         size=11, bold=True, color=RED, font="Inter", spacing=200)
add_text(slide6, RIGHT_X + cm(0.4), RIGHT_Y + cm(0.85),
         cm(10.9), cm(0.35),
         "От первого касания до партнёрства",
         size=10, color=MUTE, font="Inter")
add_rect(slide6, RIGHT_X + cm(0.4), RIGHT_Y + cm(1.3),
         cm(10.9), cm(0.06), DIVIDER)

# 5 stages of cascade
STAGES = [
    {"n": "1", "title": "ПЕРВОЕ КАСАНИЕ",
     "desc": "Email #1: знакомство с брендом Эмпилс",
     "metric": "День 0",
     "accent": ACCENT_PURPLE,
     "awareness": "0%"},
    {"n": "2", "title": "ИНФОРМАЦИОННОЕ ПРИСУТСТВИЕ",
     "desc": "Email #2-3: каталог, кейсы, преимущества",
     "metric": "Месяц 1",
     "accent": BLUE,
     "awareness": "15%"},
    {"n": "3", "title": "УЗНАВАЕМОСТЬ",
     "desc": "Email #4-6: отраслевые новости, продукт года",
     "metric": "Месяц 2-3",
     "accent": SAND,
     "awareness": "35%"},
    {"n": "4", "title": "ДОВЕРИЕ",
     "desc": "Email #7-9: отзывы, рейтинги, сертификаты",
     "metric": "Месяц 4-5",
     "accent": AMBER if False else RGBColor(0xD4, 0xA5, 0x74),
     "awareness": "60%"},
    {"n": "5", "title": "РЕШЕНИЕ О ПАРТНЁРСТВЕ",
     "desc": "Email #10+: спецпредложение для перехода",
     "metric": "Месяц 6+",
     "accent": RED,
     "awareness": "Готов"},
]

stage_y = RIGHT_Y + cm(1.6)
stage_h = cm(1.65)

for i, s in enumerate(STAGES):
    sy = stage_y + i * stage_h

    # Stage card background
    add_rect(slide6, RIGHT_X + cm(0.4), sy,
             cm(10.9), stage_h - cm(0.1), BG, corner_radius=0.04)
    # Accent left bar
    add_rect(slide6, RIGHT_X + cm(0.4), sy,
             cm(0.08), stage_h - cm(0.1), s['accent'])

    # Big number circle
    add_rect(slide6, RIGHT_X + cm(0.65), sy + cm(0.3),
             cm(0.9), cm(0.9),
             WHITE, line=s['accent'], line_w=9525, corner_radius=0.5)
    add_text(slide6, RIGHT_X + cm(0.65), sy + cm(0.45),
             cm(0.9), cm(0.6),
             s['n'], size=14, bold=True, color=s['accent'],
             font="Roboto Mono", align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)

    # Stage title
    add_text(slide6, RIGHT_X + cm(1.75), sy + cm(0.25),
             cm(7.5), cm(0.4),
             s['title'], size=10, bold=True, color=INK,
             font="Inter", spacing=50)

    # Stage description
    add_text(slide6, RIGHT_X + cm(1.75), sy + cm(0.7),
             cm(7.5), cm(0.35),
             s['desc'], size=8.5, color=MUTE,
             font="Inter")

    # Stage metric (timing)
    add_text(slide6, RIGHT_X + cm(1.75), sy + cm(1.05),
             cm(7.5), cm(0.3),
             s['metric'], size=8, bold=True, color=s['accent'],
             font="Roboto Mono")

    # Awareness metric (right side)
    add_text(slide6, RIGHT_X + cm(9.5), sy + cm(0.4),
             cm(1.7), cm(0.4),
             s['awareness'], size=12, bold=True, color=s['accent'],
             font="Roboto Mono", align=PP_ALIGN.RIGHT)
    add_text(slide6, RIGHT_X + cm(9.5), sy + cm(0.85),
             cm(1.7), cm(0.3),
             "узнаваемость", size=7.5, color=MUTE,
             font="Inter", align=PP_ALIGN.RIGHT)

# Bottom note
add_rect(slide6, RIGHT_X + cm(0.4), RIGHT_Y + cm(10.25),
         cm(10.9), cm(0.04), DIVIDER)
add_text(slide6, RIGHT_X + cm(0.4), RIGHT_Y + cm(10.4),
         cm(10.9), cm(0.3),
         "10+ касаний за 6 месяцев → готовность к партнёрству",
         size=8.5, color=MUTE, font="Inter", align=PP_ALIGN.CENTER)

# Save
prs.save(OUT)
print(f"\nSaved: {OUT}")
print(f"Slide 6: Email concept created")
print(f"Slide 8: Left half filled with novinki analytics")
