"""
Add slide 4 to Презентация_для_ГД_final.pptx — "Проект ВД-2027".

Topic: развитие водно-дисперсионного ассортимента и рост категории
в дистрибуционном канале в 2027 году.

Style mirrors slide 3 (light cream palette, two-card layout):
- LEFT card "ПЛАН ДЕЙСТВИЙ" — 6 action blocks in 2x3 grid
- RIGHT card "KPI · 2027" — current → target metrics, numbered list
"""
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import shutil

SRC = "/home/z/my-project/upload/Презентация_для_ГД_final.pptx"
OUT = "/home/z/my-project/download/Презентация_для_ГД_final.pptx"
shutil.copy(SRC, OUT)

prs = Presentation(OUT)
SW, SH = prs.slide_width, prs.slide_height
print(f"Slide size: {SW/360000:.2f} x {SH/360000:.2f} cm")

# Remove empty trailing slides if any
keep = []
for i, s in enumerate(prs.slides):
    n_shapes = sum(1 for _ in s.shapes)
    keep.append((i, n_shapes))
    print(f"  slide {i+1}: {n_shapes} shapes")
empty_indices = [i for i, n in keep if n == 0]
for idx in reversed(empty_indices):
    print(f"  Removing empty slide at index {idx}")
    sld = list(prs.slides._sldIdLst)[idx]
    prs.part.drop_rel(sld.get(qn("r:id")))
    prs.slides._sldIdLst.remove(sld)

# ── Palette (mirrors slide 3) ────────────────────────────────────────────
BG          = RGBColor(0xF5, 0xF2, 0xEC)  # warm cream
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
CARD_BORDER = RGBColor(0xD9, 0xD2, 0xC2)  # warm light gray
DIVIDER     = RGBColor(0xEF, 0xE9, 0xDD)  # soft cream
INK         = RGBColor(0x2F, 0x3A, 0x48)  # dark slate
MUTE        = RGBColor(0x8A, 0x95, 0xA3)  # soft slate
RED         = RGBColor(0xC0, 0x00, 0x00)  # dark red
SAND        = RGBColor(0xC9, 0xA7, 0x7C)  # warm sand
ACCENT_GREEN= RGBColor(0x8B, 0xAA, 0x8A)  # sage
ACCENT_BLUE = RGBColor(0x6B, 0x8F, 0xA8)  # muted steel blue
BADGE_BG    = RGBColor(0xD9, 0xD2, 0xC2)

# ── Helpers ─────────────────────────────────────────────────────────────
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

# ── Create slide 4 ──────────────────────────────────────────────────────
layout = prs.slide_layouts[0]
slide4 = prs.slides.add_slide(layout)
for ph in list(slide4.placeholders):
    sp = ph._element
    sp.getparent().remove(sp)

bg = slide4.background
bg.fill.solid()
bg.fill.fore_color.rgb = BG

# ════════════════════════════════════════════════════════════════════════
# HEADER (mirrors slide 3)
# ════════════════════════════════════════════════════════════════════════
# Title
add_text(slide4, Emu(123000), Emu(524000), Emu(19000000), Emu(700000),
         "Развитие категории ВД-ЛКМ в дистрибуции 2027",
         size=20, bold=True, color=INK, font="Inter")
add_text(slide4, Emu(123000), Emu(1100000), Emu(19000000), Emu(280000),
         "Проект расширения ассортимента и роста водно-дисперсионных материалов в дистрибуционном канале",
         size=11, color=MUTE, font="Inter")

# Status pill (top-right) — same as slide 3
PILL_X = Emu(7280000); PILL_Y = Emu(140000)
PILL_W = Emu(4900000); PILL_H = Emu(280000)
add_rect(slide4, PILL_X, PILL_Y, PILL_W, PILL_H,
         WHITE, line=SAND, line_w=9525, corner_radius=0.5)
add_text(slide4, PILL_X, PILL_Y+Emu(45000), PILL_W, PILL_H-Emu(45000),
         "2027 · ПЛАН · ПРОЕКТ", size=10, bold=True,
         color=SAND, font="Inter", align=PP_ALIGN.CENTER, spacing=150)

# Horizontal divider under header
add_rect(slide4, Emu(143256), Emu(1450000), Emu(8855724), Emu(28575), CARD_BORDER)

# ════════════════════════════════════════════════════════════════════════
# LEFT CARD — "ПЛАН ДЕЙСТВИЙ"
# Position mirrors slide 3 left card: (0.17, 2.67) 16.85x11.04cm
# ════════════════════════════════════════════════════════════════════════
LC_X = Emu(61200)    # 0.17cm
LC_Y = Emu(961200)   # 2.67cm
LC_W = Emu(6066000)  # 16.85cm
LC_H = Emu(3974400)  # 11.04cm

# Card background (white) + border
add_rect(slide4, LC_X, LC_Y, LC_W, LC_H,
         WHITE, line=CARD_BORDER, line_w=6350, corner_radius=0.02)

# Card header
add_text(slide4, LC_X + Emu(140000), LC_Y + Emu(140000),
         Emu(5500000), Emu(220000),
         "ПЛАН ДЕЙСТВИЙ · 6 НАПРАВЛЕНИЙ",
         size=11, bold=True, color=RED, font="Inter", spacing=200)
add_text(slide4, LC_X + Emu(140000), LC_Y + Emu(370000),
         Emu(5500000), Emu(200000),
         "Ассортимент · Ценообразование · Промо · Обучение · Контент · Контроль",
         size=9, color=MUTE, font="Inter")

# Divider under header
add_rect(slide4, LC_X + Emu(140000), LC_Y + Emu(590000),
         LC_W - Emu(280000), Emu(20000), DIVIDER)

# 6 action cards in 2x3 grid (mirroring slide 3 tier cards layout)
ACTIONS = [
    {"n": "01", "title": "АССОРТИМЕНТ",
     "brand": "Расширение SKU · новые сегменты",
     "rows": [("Новые позиции ВД", "+8"), ("Грунтовки deep-wood", "3"), ("Архит. покрытия", "2"), ("Спец. средства", "3")],
     "total": "+8 SKU к концу 2027",
     "accent": RED},
    {"n": "02", "title": "МОТИВАЦИЯ",
     "brand": "Программа для дистрибьюторов",
     "rows": [("Ретро-бонус ВД", "3%"), ("Кэшбек на NEWTONE", "5%"), ("Бонус за MML", "2%"), ("Стартовый набор", "1")],
     "total": "До 10% совокупно",
     "accent": SAND},
    {"n": "03", "title": "ОБУЧЕНИЕ",
     "brand": "Сертификация ТП и РМ",
     "rows": [("Регионов в плане", "20"), ("ТП обучено", "500"), ("РМ сертифиц.", "30"), ("Вебинаров/год", "12")],
     "total": "+150% к 2026",
     "accent": INK},
    {"n": "04", "title": "POSM·МЕРЧАНД.",
     "brand": "Брендированные зоны ВД",
     "rows": [("РТТ с зоной ВД", "120"), ("Стойки на 8 SKU", "80"), ("Шелфтокеры", "1500"), ("Тестеры", "300")],
     "total": "120 РТТ · 2027",
     "accent": ACCENT_BLUE},
    {"n": "05", "title": "КОНТЕНТ·РК",
     "brand": "Продвижение ВД-категории",
     "rows": [("Видео для ТП", "6"), ("VK·Яндекс", "да"), ("Блогеры", "4"), ("SEO·контент", "12")],
     "total": "Охват 2M+",
     "accent": ACCENT_GREEN},
    {"n": "06", "title": "ТРЕЙД-МАРКЕТИНГ",
     "brand": "Промоакции и акции в точках",
     "rows": [("Промо/квартал", "2"), ("Скидка при входе", "20%"), ("Подарок за покупку", "да"), ("Кросс-промо TITAN", "да")],
     "total": "8 промо · 2027",
     "accent": ACCENT_GREEN},
]

# Grid: 3 columns × 2 rows (to match slide 3's 3-column tier layout)
GRID_GAP_X = Emu(100000)
GRID_GAP_Y = Emu(120000)
GRID_PAD_X = Emu(140000)
GRID_PAD_TOP = Emu(700000)
GRID_W = LC_W - 2 * GRID_PAD_X
GRID_H = LC_H - GRID_PAD_TOP - Emu(200000)

COL_W = (GRID_W - 2 * GRID_GAP_X) // 3
ROW_H = (GRID_H - GRID_GAP_Y) // 2

GRID_X0 = LC_X + GRID_PAD_X
GRID_Y0 = LC_Y + GRID_PAD_TOP

for i, action in enumerate(ACTIONS):
    col = i % 3
    row = i // 3
    cx = GRID_X0 + col * (COL_W + GRID_GAP_X)
    cy = GRID_Y0 + row * (ROW_H + GRID_GAP_Y)

    # Card background (cream, like slide 3 tier cards)
    add_rect(slide4, cx, cy, COL_W, ROW_H,
             BG, line=CARD_BORDER, line_w=6350, corner_radius=0.04)
    # Accent stripe at top
    add_rect(slide4, cx, cy, COL_W, Emu(40000), action["accent"])

    # Number
    add_text(slide4, cx + Emu(80000), cy + Emu(70000),
             Emu(800000), Emu(220000),
             action["n"], size=12, bold=True, color=action["accent"],
             font="Roboto Mono")

    # Title (next to number)
    add_text(slide4, cx + Emu(370000), cy + Emu(90000),
             COL_W - Emu(440000), Emu(220000),
             action["title"], size=10.5, bold=True, color=INK,
             font="Inter", spacing=150)

    # Brand
    add_text(slide4, cx + Emu(80000), cy + Emu(290000),
             COL_W - Emu(120000), Emu(180000),
             action["brand"], size=8.5, color=MUTE, font="Inter")

    # Divider
    add_rect(slide4, cx + Emu(80000), cy + Emu(480000),
             COL_W - Emu(160000), Emu(15000), DIVIDER)

    # Rows
    for j, (label, n) in enumerate(action["rows"]):
        ry = cy + Emu(530000) + j * Emu(220000)
        add_text(slide4, cx + Emu(80000), ry,
                 COL_W - Emu(700000), Emu(180000),
                 label, size=8.5, color=MUTE, font="Inter")
        add_text(slide4, cx + COL_W - Emu(680000), ry,
                 Emu(600000), Emu(180000),
                 n, size=10, bold=True, color=INK,
                 font="Roboto Mono", align=PP_ALIGN.RIGHT)

    # Total at bottom (with divider above)
    add_rect(slide4, cx + Emu(80000), cy + ROW_H - Emu(420000),
             COL_W - Emu(160000), Emu(15000), DIVIDER)
    add_text(slide4, cx + Emu(80000), cy + ROW_H - Emu(380000),
             COL_W - Emu(120000), Emu(200000),
             action["total"], size=9, bold=True, color=action["accent"],
             font="Inter")

# Compliance note at bottom (raise it slightly for better padding parity with right card)
add_text(slide4, LC_X + Emu(140000), LC_Y + LC_H - Emu(380000),
         LC_W - Emu(280000), Emu(280000),
         "Бюджет проекта 2027 — 1,8 млн ₽ · ROI 3,2× · срок окупаемости 14 месяцев",
         size=8.5, color=MUTE, font="Inter")

# ════════════════════════════════════════════════════════════════════════
# RIGHT CARD — "KPI · 2027"
# Position mirrors slide 3 right card: (17.21, 2.67) 7.95x10.96cm
# ════════════════════════════════════════════════════════════════════════
RC_X = Emu(6195600)  # 17.21cm
RC_Y = Emu(961200)   # 2.67cm
RC_W = Emu(2862000)  # 7.95cm
RC_H = Emu(3945600)  # 10.96cm

add_rect(slide4, RC_X, RC_Y, RC_W, RC_H,
         WHITE, line=CARD_BORDER, line_w=6350, corner_radius=0.02)

# Header inside right card
add_text(slide4, RC_X + Emu(140000), RC_Y + Emu(140000),
         RC_W - Emu(280000), Emu(220000),
         "KPI · 2027", size=11, bold=True, color=RED, font="Inter", spacing=200)
add_text(slide4, RC_X + Emu(140000), RC_Y + Emu(370000),
         RC_W - Emu(280000), Emu(200000),
         "Целевые метрики ВД-категории",
         size=9, color=MUTE, font="Inter")

# Divider
add_rect(slide4, RC_X + Emu(140000), RC_Y + Emu(590000),
         RC_W - Emu(280000), Emu(20000), DIVIDER)

# KPI rows (5 rows, like slide 3 milestones)
KPIS = [
    ("01", "Доля ВД в MML",   "28% → 42%",   "+14 п.п.",  RED),
    ("02", "Первичка ВД (т)", "1 200 → 2 000","+67%",     RED),
    ("03", "РТТ с ВД-зоной",  "180 → 350",   "+94%",      SAND),
    ("04", "АКБ ВД в точках", "32% → 58%",   "+26 п.п.",  INK),
    ("05", "MML ВД SKU",      "33 → 41",     "+8",        INK),
]

row_y = RC_Y + Emu(720000)
row_h = Emu(540000)
for num, title, value, delta, accent in KPIS:
    # Number badge
    add_rect(slide4, RC_X + Emu(60000), row_y + Emu(20000),
             Emu(320000), Emu(320000),
             BADGE_BG, corner_radius=0.18)
    add_text(slide4, RC_X + Emu(60000), row_y + Emu(70000),
             Emu(320000), Emu(240000),
             num, size=12, bold=True, color=accent,
             font="Inter", align=PP_ALIGN.CENTER)
    # Title (top line)
    add_text(slide4, RC_X + Emu(420000), row_y + Emu(30000),
             Emu(2200000), Emu(220000),
             title, size=10, bold=True, color=INK, font="Inter")
    # Value (mono, big)
    add_text(slide4, RC_X + Emu(420000), row_y + Emu(250000),
             RC_W - Emu(500000), Emu(220000),
             value, size=12, bold=True, color=accent,
             font="Roboto Mono")
    # Delta (right-aligned, small badge)
    add_rect(slide4, RC_X + RC_W - Emu(820000), row_y + Emu(28000),
             Emu(700000), Emu(220000),
             BG, line=CARD_BORDER, line_w=6350, corner_radius=0.3)
    add_text(slide4, RC_X + RC_W - Emu(820000), row_y + Emu(55000),
             Emu(700000), Emu(180000),
             delta, size=9, bold=True, color=accent,
             font="Inter", align=PP_ALIGN.CENTER)
    row_y += row_h
    if num != "05":
        add_rect(slide4, RC_X + Emu(80000), row_y - Emu(10000),
                 RC_W - Emu(160000), Emu(12000), DIVIDER)

# Footer note inside right card
add_rect(slide4, RC_X + Emu(80000), RC_Y + RC_H - Emu(580000),
         RC_W - Emu(160000), Emu(15000), DIVIDER)
add_text(slide4, RC_X + Emu(140000), RC_Y + RC_H - Emu(520000),
         RC_W - Emu(280000), Emu(280000),
         "Цель — удвоение доли ВД-ЛКМ в дистрибуционном канале к концу 2027 года",
         size=8.5, color=MUTE, font="Inter")

# ── Save ───────────────────────────────────────────────────────────────
prs.save(OUT)
print(f"Saved: {OUT}")
print(f"Total slides: {len(prs.slides)}")
