"""
Add slide 3 to Презентация_для_ГД.pptx — "Выход в жилое строительство".

Style mirrors slide 2 (Белые пятна):
- Light/cream palette: #F5F2EC bg, #FFFFFF cards, #D9D2C2 borders, #EFE9DD dividers
- Dark slate text #2F3A48, muted #8A95A3
- Dark red accent #C00000 + warm sand #C9A77C
- Two-card layout: LEFT (16.86x10.03cm) + RIGHT (7.94x9.99cm)
- Numbered badge list on the right (5 rows, like the meetings list)

Content:
- LEFT card "АССОРТИМЕНТ" — 3 product tiers + 25 SKU total
- RIGHT card "КАТАЛОГ" — 5 project milestones (v1, ТЗ, etc.)
"""
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import shutil, os

SRC = "/home/z/my-project/upload/Презентация_для_ГД.pptx"
OUT = "/home/z/my-project/download/Презентация_для_ГД.pptx"
shutil.copy(SRC, OUT)

prs = Presentation(OUT)
SW, SH = prs.slide_width, prs.slide_height
print(f"Slide size: {SW/360000:.2f} x {SH/360000:.2f} cm")

# Delete any trailing empty slides (source has an empty slide 3)
from pptx.oxml.ns import qn
keep = []
for i, s in enumerate(prs.slides):
    n_shapes = sum(1 for _ in s.shapes)
    print(f"  slide {i+1}: {n_shapes} shapes")
    keep.append((i, n_shapes))
# Remove slides with 0 shapes (highest index first to preserve ordering)
empty_indices = [i for i, n in keep if n == 0]
for idx in reversed(empty_indices):
    print(f"  Removing empty slide at index {idx}")
    sld = list(prs.slides._sldIdLst)[idx]
    prs.part.drop_rel(sld.get(qn("r:id")))
    prs.slides._sldIdLst.remove(sld)

# ── Palette (matches slide 2) ────────────────────────────────────────────
BG          = RGBColor(0xF5, 0xF2, 0xEC)  # warm cream
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
CARD_BORDER = RGBColor(0xD9, 0xD2, 0xC2)  # warm light gray
DIVIDER     = RGBColor(0xEF, 0xE9, 0xDD)  # soft cream
INK         = RGBColor(0x2F, 0x3A, 0x48)  # dark slate
MUTE        = RGBColor(0x8A, 0x95, 0xA3)  # soft slate
RED         = RGBColor(0xC0, 0x00, 0x00)  # dark red accent
SAND        = RGBColor(0xC9, 0xA7, 0x7C)  # warm sand
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

# ── Create slide 3 (use first layout, clear placeholders) ───────────────
layout = prs.slide_layouts[0]
slide3 = prs.slides.add_slide(layout)
for ph in list(slide3.placeholders):
    sp = ph._element
    sp.getparent().remove(sp)

# Slide background
bg = slide3.background
bg.fill.solid()
bg.fill.fore_color.rgb = BG

# ════════════════════════════════════════════════════════════════════════
# HEADER (mirrors slide 2 exactly)
# ════════════════════════════════════════════════════════════════════════
# TITAN logo (dark red)
add_text(slide3, Emu(228600), Emu(182880), Emu(1200000), Emu(251460),
         "TITAN", size=18, bold=True, color=RED, font="Inter", spacing=200)

# Vertical divider
add_rect(slide3, Emu(997660), Emu(182880), Emu(28575), Emu(251460), CARD_BORDER)

# Eyebrow
add_text(slide3, Emu(1116614), Emu(224939), Emu(10500000), Emu(182880),
         "ВЫХОД В ЖИЛОЕ СТРОИТЕЛЬСТВО", size=11, bold=True,
         color=INK, font="Inter", spacing=250)

# Main title
add_text(slide3, Emu(228600), Emu(457200), Emu(17000000), Emu(457200),
         "Каталог продуктов для жилищного строительства",
         size=22, bold=True, color=INK, font="Inter")

# Status pill (top-right)
PILL_X = Emu(7315200); PILL_Y = Emu(182880)
PILL_W = Emu(4800600); PILL_H = Emu(285750)
add_rect(slide3, PILL_X, PILL_Y, PILL_W, PILL_H,
         WHITE, line=SAND, line_w=9525, corner_radius=0.5)
add_text(slide3, PILL_X, PILL_Y+Emu(45720), PILL_W, PILL_H-Emu(45720),
         "СЕНТЯБРЬ 2026 · В РАБОТЕ", size=10, bold=True,
         color=SAND, font="Inter", align=PP_ALIGN.CENTER, spacing=150)

# Horizontal divider under header
add_rect(slide3, Emu(143256), Emu(548640), Emu(8855724), Emu(28575), CARD_BORDER)

# ════════════════════════════════════════════════════════════════════════
# LEFT CARD — "АССОРТИМЕНТ"
# Position mirrors slide 2 left card: (0.17, 2.16) 16.86x10.03cm
# ════════════════════════════════════════════════════════════════════════
LC_X = Emu(61200)    # 0.17cm
LC_Y = Emu(777240)   # 2.16cm
LC_W = Emu(6066960)  # 16.86cm
LC_H = Emu(4110000)  # 11.42cm (was 10.03 — extended to fit content)

# Card background (white) + border
add_rect(slide3, LC_X, LC_Y, LC_W, LC_H,
         WHITE, line=CARD_BORDER, line_w=6350, corner_radius=0.02)

# Card header
add_text(slide3, LC_X + Emu(80000), LC_Y + Emu(120000), Emu(5500000), Emu(220000),
         "АССОРТИМЕНТ", size=11, bold=True, color=RED, font="Inter", spacing=200)
add_text(slide3, LC_X + Emu(80000), LC_Y + Emu(360000), Emu(5500000), Emu(200000),
         "25 SKU · 3 уровня цен · 9 зон жилого здания",
         size=9, color=MUTE, font="Inter")

# Divider
add_rect(slide3, LC_X + Emu(80000), LC_Y + Emu(560000),
         LC_W - Emu(160000), Emu(20000), DIVIDER)

# Three tier blocks — horizontal layout
TIERS = [
    {
        "n": "01", "name": "БЮДЖЕТ",
        "brand": "ОРЕОЛ ДИСКОНТ",
        "counts": [("Краски и эмали", "5"), ("Шпатлёвки", "2"), ("Грунтовки", "4")],
        "total": "11 SKU",
        "accent": RED,
    },
    {
        "n": "02", "name": "СРЕДНИЙ",
        "brand": "ОРЕОЛ",
        "counts": [("Краски и эмали", "4"), ("Шпатлёвки", "3"), ("Грунтовки", "1")],
        "total": "8 SKU",
        "accent": SAND,
    },
    {
        "n": "03", "name": "ПРЕМИУМ",
        "brand": "NEWTONE · AVANTGARDE",
        "counts": [("Краски и эмали", "4"), ("Шпатлёвки", "1"), ("Грунтовки", "1")],
        "total": "6 SKU",
        "accent": INK,
    },
]

# Tier cards: 3 columns inside left card
TIER_GAP = Emu(120000)
TIER_W = (LC_W - Emu(160000) - 2 * TIER_GAP) // 3   # ~5.65cm each
TIER_H = Emu(2000000)   # was 2200000 — slightly shorter so SP cards fit
TIER_Y = LC_Y + Emu(700000)
TIER_X_START = LC_X + Emu(80000)

for i, tier in enumerate(TIERS):
    tx = TIER_X_START + i * (TIER_W + TIER_GAP)
    # Card background
    add_rect(slide3, tx, TIER_Y, TIER_W, TIER_H,
             BG, line=CARD_BORDER, line_w=6350, corner_radius=0.04)
    # Accent stripe at top
    add_rect(slide3, tx, TIER_Y, TIER_W, Emu(50000), tier["accent"])
    # Number + name row
    add_text(slide3, tx + Emu(80000), TIER_Y + Emu(90000),
             Emu(800000), Emu(220000),
             tier["n"], size=14, bold=True, color=tier["accent"],
             font="Roboto Mono")
    add_text(slide3, tx + Emu(360000), TIER_Y + Emu(110000),
             TIER_W - Emu(420000), Emu(220000),
             tier["name"], size=12, bold=True, color=INK,
             font="Inter", spacing=200)
    # Brand
    add_text(slide3, tx + Emu(80000), TIER_Y + Emu(330000),
             TIER_W - Emu(120000), Emu(180000),
             tier["brand"], size=9, color=MUTE, font="Inter", spacing=80)
    # Divider
    add_rect(slide3, tx + Emu(80000), TIER_Y + Emu(530000),
             TIER_W - Emu(160000), Emu(15000), DIVIDER)
    # Counts rows
    for j, (label, n) in enumerate(tier["counts"]):
        ry = TIER_Y + Emu(590000) + j * Emu(280000)
        add_text(slide3, tx + Emu(80000), ry,
                 Emu(2000000), Emu(180000),
                 label, size=9.5, color=MUTE, font="Inter")
        add_text(slide3, tx + TIER_W - Emu(700000), ry,
                 Emu(620000), Emu(180000),
                 n, size=12, bold=True, color=INK,
                 font="Roboto Mono", align=PP_ALIGN.RIGHT)
    # Total at bottom
    add_rect(slide3, tx + Emu(80000), TIER_Y + TIER_H - Emu(500000),
             TIER_W - Emu(160000), Emu(15000), DIVIDER)
    add_text(slide3, tx + Emu(80000), TIER_Y + TIER_H - Emu(450000),
             Emu(2000000), Emu(180000),
             "Итого:", size=9, color=MUTE, font="Inter")
    add_text(slide3, tx + TIER_W - Emu(700000), TIER_Y + TIER_H - Emu(460000),
             Emu(620000), Emu(200000),
             tier["total"], size=13, bold=True, color=tier["accent"],
             font="Roboto Mono", align=PP_ALIGN.RIGHT)

# Special products row (under tier cards)
SPEC_Y = TIER_Y + TIER_H + Emu(150000)   # was 180000
add_text(slide3, LC_X + Emu(80000), SPEC_Y,
         Emu(5500000), Emu(180000),
         "СПЕЦ·ПРОДУКТЫ", size=10, bold=True, color=MUTE,
         font="Inter", spacing=200)

# Two special cards
SPECS = [
    ("PROFIWOOD ББ-11", "Огнезащита древесины · Г1 В1 Д1 Т1 РП1", "I-II гр."),
    ("EMPILS E 608", "Эпоксидный компаунд для полов · Г1 В1 Д1 Т1 РП1", "Высокие нагрузки"),
]
SP_W = (LC_W - Emu(160000) - Emu(120000)) // 2
SP_H = Emu(620000)   # was 700000
SP_Y = SPEC_Y + Emu(220000)   # was 240000
for i, (name, desc, tag) in enumerate(SPECS):
    sx = LC_X + Emu(80000) + i * (SP_W + Emu(120000))
    add_rect(slide3, sx, SP_Y, SP_W, SP_H,
             BG, line=CARD_BORDER, line_w=6350, corner_radius=0.04)
    add_text(slide3, sx + Emu(80000), SP_Y + Emu(80000),
             SP_W - Emu(160000), Emu(220000),
             name, size=11, bold=True, color=INK, font="Inter")
    add_text(slide3, sx + Emu(80000), SP_Y + Emu(310000),
             SP_W - Emu(160000), Emu(180000),
             desc, size=8.5, color=MUTE, font="Inter")
    add_text(slide3, sx + SP_W - Emu(1700000), SP_Y + Emu(80000),
             Emu(1620000), Emu(220000),
             tag, size=8.5, bold=True, color=RED,
             font="Inter", align=PP_ALIGN.RIGHT, spacing=80)

# Compliance note at bottom of left card
add_text(slide3, LC_X + Emu(80000), LC_Y + LC_H - Emu(380000),
         LC_W - Emu(160000), Emu(280000),
         "Все продукты сертифицированы по ФЗ-123 · класс КМ1 превышает базовые требования для жилых зданий",
         size=8.5, color=MUTE, font="Inter")

# ════════════════════════════════════════════════════════════════════════
# RIGHT CARD — "КАТАЛОГ" (project milestones, list format like slide 2)
# Position mirrors slide 2 right card: (17.20, 2.16) 7.94x9.99cm
# ════════════════════════════════════════════════════════════════════════
RC_X = Emu(6194040)  # 17.20cm
RC_Y = Emu(777240)   # 2.16cm
RC_W = Emu(2860320)  # 7.94cm
RC_H = Emu(4110000)  # 11.42cm (was 9.99 — match left card height)

add_rect(slide3, RC_X, RC_Y, RC_W, RC_H,
         WHITE, line=CARD_BORDER, line_w=6350, corner_radius=0.02)

# Header inside right card
add_text(slide3, RC_X + Emu(80000), RC_Y + Emu(120000),
         RC_W - Emu(160000), Emu(220000),
         "КАТАЛОГ", size=11, bold=True, color=RED, font="Inter", spacing=200)
add_text(slide3, RC_X + Emu(80000), RC_Y + Emu(360000),
         RC_W - Emu(160000), Emu(200000),
         "29.09.2026 · v1 готова · 25 SKU",
         size=9, color=MUTE, font="Inter")

# Divider
add_rect(slide3, RC_X + Emu(80000), RC_Y + Emu(560000),
         RC_W - Emu(160000), Emu(20000), DIVIDER)

# Milestones list (5 rows, like slide 2's meetings)
MILESTONES = [
    ("01", "ТЗ v4",            "финальная версия · 01.10.2026",  RED),
    ("02", "Каталог v1",       "29.09.2026 · первый макет",      RED),
    ("03", "Объём",             "24-32 стр. · A4 · PDF + печать", INK),
    ("04", "Целевая аудитория", "прорабы · ИТР · технадзор",     INK),
    ("05", "Структура",         "7 разделов · 9 зон · 11 систем", INK),
]

row_y = RC_Y + Emu(680000)
row_h = Emu(500000)
for num, title, sub, accent in MILESTONES:
    # Number badge
    add_rect(slide3, RC_X + Emu(60000), row_y + Emu(20000),
             Emu(320000), Emu(320000),
             BADGE_BG, corner_radius=0.18)
    add_text(slide3, RC_X + Emu(60000), row_y + Emu(70000),
             Emu(320000), Emu(240000),
             num, size=12, bold=True, color=accent,
             font="Inter", align=PP_ALIGN.CENTER)
    # Title (bold)
    add_text(slide3, RC_X + Emu(420000), row_y + Emu(30000),
             Emu(2300000), Emu(220000),
             title, size=11, bold=True, color=INK, font="Inter")
    # Subtitle (gray)
    add_text(slide3, RC_X + Emu(420000), row_y + Emu(260000),
             RC_W - Emu(500000), Emu(200000),
             sub, size=8.5, color=accent, font="Inter")
    row_y += row_h
    if num != "05":
        add_rect(slide3, RC_X + Emu(80000), row_y - Emu(10000),
                 RC_W - Emu(160000), Emu(12000), DIVIDER)

# Footer note inside right card (slightly higher to avoid edge clipping)
add_rect(slide3, RC_X + Emu(80000), RC_Y + RC_H - Emu(550000),
         RC_W - Emu(160000), Emu(15000), DIVIDER)
add_text(slide3, RC_X + Emu(80000), RC_Y + RC_H - Emu(500000),
         RC_W - Emu(200000), Emu(220000),
         "Срок согласования макета — 2 недели после ТЗ",
         size=8, color=MUTE, font="Inter")

# ════════════════════════════════════════════════════════════════════════
# Save
# ════════════════════════════════════════════════════════════════════════
prs.save(OUT)
print(f"Saved: {OUT}")
print(f"Total slides: {len(prs.slides)}")
