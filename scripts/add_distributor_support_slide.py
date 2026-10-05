"""
Add slide "Поддержка новых дистрибьюторов" to Презентация_для_ГД_final.pptx.

Theme: When entering a new region with a new distributor, after signing the
distribution agreement, Empils offers support measures:
1. Participation in regional profile exhibition (Distributor + Empils co-branding)
2. On-site training of new distributor with focus on FTM (ФТМ) sales

Style: Minimalist, mirrors slide 5 (Запуск новинок).
- LEFT card: 4-stage support timeline (Preparation → Exhibition → Training → Support)
- RIGHT card: KPI metrics for the support program
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
print(f"Slide size: {prs.slide_width/360000:.2f} x {prs.slide_height/360000:.2f} cm")
print(f"Source slides: {len(prs.slides)}")

# Remove empty trailing slides
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

# ── Palette (matches slide 5) ────────────────────────────────────────
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

# ── Helpers ─────────────────────────────────────────────────────────
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

# ── Create slide ────────────────────────────────────────────────────
layout = prs.slide_layouts[0]
slide_new = prs.slides.add_slide(layout)
for ph in list(slide_new.placeholders):
    sp = ph._element
    sp.getparent().remove(sp)

bg = slide_new.background
bg.fill.solid()
bg.fill.fore_color.rgb = BG

# ════════════════════════════════════════════════════════════════════════
# HEADER (mirrors slide 5 layout exactly)
# Title at y=0.50cm, subtitle at y=1.56cm, divider at y=2.22cm, cards at y=2.67cm
# ════════════════════════════════════════════════════════════════════════
# Title
add_text(slide_new, Emu(140000), Emu(180000), Emu(8800000), Emu(340000),
         "Поддержка новых дистрибьюторов",
         size=20, bold=True, color=INK, font="Inter")
# Subtitle
add_text(slide_new, Emu(140000), Emu(560000), Emu(8800000), Emu(360000),
         "Программа выхода в новые регионы: от подписания договора до первых продаж",
         size=11, color=MUTE, font="Inter")
# Status pill (top-right)
PILL_X = Emu(7200000); PILL_Y = Emu(160000)
PILL_W = Emu(1800000); PILL_H = Emu(280000)
add_rect(slide_new, PILL_X, PILL_Y, PILL_W, PILL_H,
         WHITE, line=SAND, line_w=9525, corner_radius=0.5)
add_text(slide_new, PILL_X, PILL_Y+Emu(45000), PILL_W, PILL_H-Emu(45000),
         "ПРОГРАММА · 2026", size=10, bold=True,
         color=SAND, font="Inter", align=PP_ALIGN.CENTER, spacing=150)
# Divider
add_rect(slide_new, Emu(140000), Emu(800000), Emu(8855724), Emu(28575), CARD_BORDER)

# ════════════════════════════════════════════════════════════════════════
# LEFT CARD — "4 ЭТАПА ПОДДЕРЖКИ" (timeline)
# Position mirrors slide 5 left card: x=0.17, y=2.67, w=16.85, h=9.72
# ════════════════════════════════════════════════════════════════════════
LC_X = Emu(61200)     # 0.17cm
LC_Y = Emu(960000)    # 2.67cm
LC_W = Emu(6066960)   # 16.85cm
LC_H = Emu(3500000)   # 9.72cm

add_rect(slide_new, LC_X, LC_Y, LC_W, LC_H,
         WHITE, line=CARD_BORDER, line_w=6350, corner_radius=0.02)

add_text(slide_new, LC_X + Emu(200000), LC_Y + Emu(140000),
         Emu(8000000), Emu(220000),
         "4 ЭТАПА ПОДДЕРЖКИ НОВОГО ДИСТРИБЬЮТОРА",
         size=11, bold=True, color=RED, font="Inter", spacing=200)
add_text(slide_new, LC_X + Emu(200000), LC_Y + Emu(370000),
         Emu(8000000), Emu(180000),
         "От подписания договора до первых продаж в новом регионе",
         size=10, color=MUTE, font="Inter")
# Divider under header
add_rect(slide_new, LC_X + Emu(200000), LC_Y + Emu(580000),
         LC_W - Emu(400000), Emu(20000), DIVIDER)

# 4 timeline stages — horizontal layout
STAGES = [
    {
        "n": "01",
        "title": "ПОДГОТОВКА",
        "subtitle": "Анализ региона · Выбор выставки · План обучения",
        "duration": "Дни 1—30",
        "accent": RED,
    },
    {
        "n": "02",
        "title": "ВЫСТАВКА",
        "subtitle": "Совместное участие Дистрибьютор + Эмпилс",
        "duration": "Дни 30—45",
        "accent": SAND,
    },
    {
        "n": "03",
        "title": "ОБУЧЕНИЕ",
        "subtitle": "Выездное обучение с упором на продажи ФТМ",
        "duration": "Дни 45—60",
        "accent": SAND,
    },
    {
        "n": "04",
        "title": "СОПРОВОЖДЕНИЕ",
        "subtitle": "Мониторинг · KPI · Корректировка",
        "duration": "Дни 60—90",
        "accent": INK,
    },
]

STEP_GAP = Emu(120000)
STEP_PAD_X = Emu(200000)
STEP_W = (LC_W - 2 * STEP_PAD_X - 3 * STEP_GAP) // 4
STEP_H = Emu(2300000)
STEP_Y = LC_Y + Emu(750000)
STEP_X0 = LC_X + STEP_PAD_X

# Connecting line through badges
LINE_Y = STEP_Y + Emu(280000)
add_rect(slide_new, STEP_X0 + Emu(80000), LINE_Y,
         LC_W - 2 * STEP_PAD_X - Emu(160000), Emu(12000), DIVIDER)

for i, stage in enumerate(STAGES):
    cx = STEP_X0 + i * (STEP_W + STEP_GAP)

    # Number badge (circle)
    BADGE_SIZE = Emu(560000)
    add_rect(slide_new, cx + (STEP_W - BADGE_SIZE) // 2,
             STEP_Y, BADGE_SIZE, BADGE_SIZE,
             WHITE, line=stage["accent"], line_w=12700, corner_radius=0.5)
    add_text(slide_new, cx + (STEP_W - BADGE_SIZE) // 2,
             STEP_Y + Emu(110000),
             BADGE_SIZE, Emu(340000),
             stage["n"], size=14, bold=True, color=stage["accent"],
             font="Roboto Mono", align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)

    # Stage title (smaller font, no letter-spacing, to fit long words on one line)
    add_text(slide_new, cx, STEP_Y + Emu(680000),
             STEP_W, Emu(280000),
             stage["title"], size=10, bold=True, color=INK,
             font="Inter", align=PP_ALIGN.CENTER)

    # Duration
    add_text(slide_new, cx, STEP_Y + Emu(980000),
             STEP_W, Emu(220000),
             stage["duration"], size=9, bold=True, color=stage["accent"],
             font="Roboto Mono", align=PP_ALIGN.CENTER)

    # Subtitle
    add_text(slide_new, cx + Emu(40000), STEP_Y + Emu(1300000),
             STEP_W - Emu(80000), Emu(800000),
             stage["subtitle"], size=9.5, color=MUTE,
             font="Inter", align=PP_ALIGN.CENTER)

# Bottom note
add_rect(slide_new, LC_X + Emu(200000), LC_Y + LC_H - Emu(550000),
         LC_W - Emu(400000), Emu(15000), DIVIDER)
add_text(slide_new, LC_X + Emu(200000), LC_Y + LC_H - Emu(500000),
         LC_W - Emu(400000), Emu(380000),
         "Программа обязательна при выходе в регионы, где продукция Эмпилс ранее не была представлена.",
         size=9, color=MUTE, font="Inter", align=PP_ALIGN.CENTER)

# ════════════════════════════════════════════════════════════════════════
# RIGHT CARD — "KPI ПРОГРАММЫ" (4 metrics)
# Position mirrors slide 5 right card: x=17.21cm, y=2.67cm, w=7.95cm, h=9.72cm
# ════════════════════════════════════════════════════════════════════════
RC_X = Emu(6195600)
RC_Y = Emu(960000)
RC_W = Emu(2862000)
RC_H = Emu(3500000)

add_rect(slide_new, RC_X, RC_Y, RC_W, RC_H,
         WHITE, line=CARD_BORDER, line_w=6350, corner_radius=0.02)

# Header
add_text(slide_new, RC_X + Emu(200000), RC_Y + Emu(140000),
         RC_W - Emu(400000), Emu(220000),
         "KPI ПРОГРАММЫ",
         size=11, bold=True, color=RED, font="Inter", spacing=200)
add_text(slide_new, RC_X + Emu(200000), RC_Y + Emu(370000),
         RC_W - Emu(400000), Emu(180000),
         "Метрики за первые 90 дней работы",
         size=10, color=MUTE, font="Inter")
add_rect(slide_new, RC_X + Emu(200000), RC_Y + Emu(580000),
         RC_W - Emu(400000), Emu(20000), DIVIDER)

# 4 KPI rows
KPIS = [
    ("1 выставка", "Профильное мероприятие в регионе", RED),
    ("100%", "ТП дистрибьютора обучено продажам ФТМ", SAND),
    ("30 дней", "От договора до первой отгрузки", SAND),
    ("90 дней", "Период активной поддержки Эмпилс", INK),
]

kpi_y = RC_Y + Emu(720000)
kpi_h = Emu(580000)
for i, (value, label, accent) in enumerate(KPIS):
    # Accent vertical bar on left
    add_rect(slide_new, RC_X + Emu(200000), kpi_y + Emu(30000),
             Emu(40000), kpi_h - Emu(60000), accent)

    # Value (large, mono)
    add_text(slide_new, RC_X + Emu(360000), kpi_y + Emu(40000),
             Emu(1600000), Emu(300000),
             value, size=16, bold=True, color=accent,
             font="Roboto Mono")

    # Label
    add_text(slide_new, RC_X + Emu(200000), kpi_y + Emu(330000),
             RC_W - Emu(400000), Emu(200000),
             label, size=10, color=MUTE, font="Inter")

    kpi_y += kpi_h
    if i < len(KPIS) - 1:
        add_rect(slide_new, RC_X + Emu(200000), kpi_y - Emu(80000),
                 RC_W - Emu(400000), Emu(12000), DIVIDER)

# Footer note inside right card
add_rect(slide_new, RC_X + Emu(200000), RC_Y + RC_H - Emu(450000),
         RC_W - Emu(400000), Emu(15000), DIVIDER)
add_text(slide_new, RC_X + Emu(200000), RC_Y + RC_H - Emu(400000),
         RC_W - Emu(400000), Emu(280000),
         "Без выполнения KPI программа поддержки не продлевается",
         size=9, color=MUTE, font="Inter", align=PP_ALIGN.CENTER)

# ── Save ───────────────────────────────────────────────────────────
prs.save(OUT)
print(f"\nSaved: {OUT}")
print(f"Total slides: {len(prs.slides)}")
