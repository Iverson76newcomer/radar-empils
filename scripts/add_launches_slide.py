"""
Add slide 7 "Запуск новинок и стандартизация" to Презентация_для_ГД_final.pptx.

Style: MINIMALIST — clean, plenty of whitespace, simple two-card layout
- LEFT card: 4 этапа стандарта вывода новинки (vertical timeline)
- RIGHT card: метрики поддержки периода выхода

Visual style mirrors slide 5 (Housing) — same pastel palette, same card geometry.
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

# ── Palette (mirrors slide 5) ────────────────────────────────────────
BG          = RGBColor(0xF5, 0xF2, 0xEC)  # warm cream
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
CARD_BORDER = RGBColor(0xD9, 0xD2, 0xC2)  # warm light gray
DIVIDER     = RGBColor(0xEF, 0xE9, 0xDD)  # soft cream
INK         = RGBColor(0x2F, 0x3A, 0x48)  # dark slate
MUTE        = RGBColor(0x8A, 0x95, 0xA3)  # soft slate
RED         = RGBColor(0xC0, 0x00, 0x00)  # dark red accent
SAND        = RGBColor(0xC9, 0xA7, 0x7C)  # warm sand

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

# ── Create slide 7 ──────────────────────────────────────────────────
layout = prs.slide_layouts[0]
slide7 = prs.slides.add_slide(layout)
for ph in list(slide7.placeholders):
    sp = ph._element
    sp.getparent().remove(sp)

bg = slide7.background
bg.fill.solid()
bg.fill.fore_color.rgb = BG

# ════════════════════════════════════════════════════════════════════════
# HEADER (mirrors slide 5 minimal style)
# ════════════════════════════════════════════════════════════════════════
# Title (large, clean)
add_text(slide7, Emu(140000), Emu(524000), Emu(19000000), Emu(700000),
         "Запуск новинок · Стандартизация процесса",
         size=22, bold=True, color=INK, font="Inter")

# Subtitle (single line, muted)
add_text(slide7, Emu(140000), Emu(1100000), Emu(19000000), Emu(280000),
         "Единый стандарт вывода новинок и поддержки в период выхода на рынок",
         size=12, color=MUTE, font="Inter")

# Status pill (top-right)
PILL_X = Emu(7280000); PILL_Y = Emu(450000)
PILL_W = Emu(4900000); PILL_H = Emu(280000)
add_rect(slide7, PILL_X, PILL_Y, PILL_W, PILL_H,
         WHITE, line=SAND, line_w=9525, corner_radius=0.5)
add_text(slide7, PILL_X, PILL_Y+Emu(45000), PILL_W, PILL_H-Emu(45000),
         "СТАНДАРТ · ВНЕДРЕНИЕ", size=10, bold=True,
         color=SAND, font="Inter", align=PP_ALIGN.CENTER, spacing=150)

# Horizontal divider under header
add_rect(slide7, Emu(140000), Emu(1450000), Emu(22250000), Emu(28575), CARD_BORDER)

# ════════════════════════════════════════════════════════════════════════
# LEFT CARD — "СТАНДАРТ ВЫВОДА" (4 этапа timeline)
# Position: (0.17cm, 2.85cm) - wide enough for horizontal timeline
# ════════════════════════════════════════════════════════════════════════
LC_X = Emu(61200)     # 0.17cm
LC_Y = Emu(1026000)   # 2.85cm
LC_W = Emu(16800000)  # 16.8cm — wide left card
LC_H = Emu(3500000)   # 9.7cm

# Card background
add_rect(slide7, LC_X, LC_Y, LC_W, LC_H,
         WHITE, line=CARD_BORDER, line_w=6350, corner_radius=0.02)

# Card header (left aligned, simple)
add_text(slide7, LC_X + Emu(200000), LC_Y + Emu(140000),
         Emu(8000000), Emu(220000),
         "СТАНДАРТ ВЫВОДА НОВИНКИ",
         size=11, bold=True, color=RED, font="Inter", spacing=200)
add_text(slide7, LC_X + Emu(200000), LC_Y + Emu(370000),
         Emu(8000000), Emu(180000),
         "4 обязательных этапа для каждой новинки",
         size=10, color=MUTE, font="Inter")

# Divider under header
add_rect(slide7, LC_X + Emu(200000), LC_Y + Emu(580000),
         LC_W - Emu(400000), Emu(20000), DIVIDER)

# 4 timeline steps — horizontal layout
STAGES = [
    {
        "n": "01",
        "title": "ПОДГОТОВКА",
        "subtitle": "Анализ рынка · Ценообразование · MML",
        "duration": "−60 дней",
        "accent": RED,
    },
    {
        "n": "02",
        "title": "ЗАПУСК",
        "subtitle": "POSM · Обучение · Контент",
        "duration": "0 — 30 дней",
        "accent": SAND,
    },
    {
        "n": "03",
        "title": "СОПРОВОЖДЕНИЕ",
        "subtitle": "Промо · МК · Куратор · Аналитика",
        "duration": "30 — 90 дней",
        "accent": SAND,
    },
    {
        "n": "04",
        "title": "ОЦЕНКА",
        "subtitle": "KPI · Решение по масштабу",
        "duration": "90 — 180 дней",
        "accent": INK,
    },
]

# Horizontal timeline: 4 stages side-by-side
STEP_GAP = Emu(120000)
STEP_PAD_X = Emu(200000)
STEP_W = (LC_W - 2 * STEP_PAD_X - 3 * STEP_GAP) // 4
STEP_H = Emu(2300000)
STEP_Y = LC_Y + Emu(750000)
STEP_X0 = LC_X + STEP_PAD_X

# Draw a horizontal connecting line first (behind steps)
LINE_Y = STEP_Y + Emu(280000)  # at number badge center
add_rect(slide7, STEP_X0 + Emu(80000), LINE_Y,
         LC_W - 2 * STEP_PAD_X - Emu(160000), Emu(12000), DIVIDER)

for i, stage in enumerate(STAGES):
    cx = STEP_X0 + i * (STEP_W + STEP_GAP)

    # Number badge (circle-like rounded square)
    BADGE_SIZE = Emu(560000)
    add_rect(slide7, cx + (STEP_W - BADGE_SIZE) // 2,
             STEP_Y, BADGE_SIZE, BADGE_SIZE,
             WHITE, line=stage["accent"], line_w=12700, corner_radius=0.5)
    add_text(slide7, cx + (STEP_W - BADGE_SIZE) // 2,
             STEP_Y + Emu(110000),
             BADGE_SIZE, Emu(340000),
             stage["n"], size=14, bold=True, color=stage["accent"],
             font="Roboto Mono", align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)

    # Stage title (below badge)
    add_text(slide7, cx, STEP_Y + Emu(680000),
             STEP_W, Emu(280000),
             stage["title"], size=11, bold=True, color=INK,
             font="Inter", align=PP_ALIGN.CENTER, spacing=100)

    # Duration (smaller, accent color)
    add_text(slide7, cx, STEP_Y + Emu(980000),
             STEP_W, Emu(220000),
             stage["duration"], size=9, bold=True, color=stage["accent"],
             font="Roboto Mono", align=PP_ALIGN.CENTER)

    # Subtitle (muted, smaller)
    add_text(slide7, cx + Emu(40000), STEP_Y + Emu(1300000),
             STEP_W - Emu(80000), Emu(800000),
             stage["subtitle"], size=9.5, color=MUTE,
             font="Inter", align=PP_ALIGN.CENTER)

# Bottom note inside left card
add_rect(slide7, LC_X + Emu(200000), LC_Y + LC_H - Emu(550000),
         LC_W - Emu(400000), Emu(15000), DIVIDER)
add_text(slide7, LC_X + Emu(200000), LC_Y + LC_H - Emu(500000),
         LC_W - Emu(400000), Emu(380000),
         "Стандарт обязателен для всех новинок. Без прохождения 4 этапов новинка не вводится в MML.",
         size=9, color=MUTE, font="Inter", align=PP_ALIGN.CENTER)

# ════════════════════════════════════════════════════════════════════════
# RIGHT CARD — "ПОДДЕРЖКА В ПЕРИОД ВЫХОДА" (KPI metrics)
# Position: (18cm, 2.85cm) - smaller, focused
# ════════════════════════════════════════════════════════════════════════
RC_X = Emu(17000000)  # ~4.72cm from right edge (slide width 25.4cm)
RC_Y = Emu(1026000)   # 2.85cm
RC_W = Emu(8500000)   # ~2.36cm — wait that's too small. Let me re-check.
# Actually slide width is 9144000 EMU = 25.4cm
# We want right card to span from ~17.5cm to ~25cm, so width = 7.5cm
# 7.5cm = 7.5 * 360000 = 2700000 EMU
RC_X = Emu(17000000 // 360000 * 360000)  # round to 47220 EMU... let me just use direct
# Simpler: left card ends at LC_X + LC_W = 61200 + 16800000 = 16861200
# Slide width = 9144000 (no wait, that's wrong)
# Actually slide is 9144000 EMU wide × 5143500 tall? Let me re-check

# Print actual slide size
print(f"Slide W x H = {prs.slide_width} x {prs.slide_height} EMU")
print(f"  = {prs.slide_width/360000:.2f} x {prs.slide_height/360000:.2f} cm")

# Hmm earlier we saw 25.40 x 14.29 cm. So slide_width = 25.4 * 360000 = 9144000 EMU
# But that's a custom size — let me just use it
slide_w_emu = prs.slide_width  # 9144000
slide_h_emu = prs.slide_height  # 5143500... wait no

# Wait — earlier in this very session (housing slide), we used 16.85cm + 7.95cm = 24.8cm for cards
# which is less than 25.4cm. So LC_W = 6066960 EMU was 16.85cm → 1cm = 360000 EMU ✓
# Now I used LC_W = 16800000 EMU = 46.67cm — that's TOO BIG!

# Reset dimensions properly. Slide is 25.4 x 14.29 cm.
# Want LEFT card 16.85 wide (matches existing slides), RIGHT card 7.95 wide
# Gap = 25.4 - 16.85 - 7.95 = 0.6 cm

LC_W = Emu(6066960)   # 16.85cm
LC_H = Emu(3500000)   # ~9.72cm (we want shorter — minimalist)

# Re-place right card
RC_X = Emu(6195600)   # 17.21cm (matches slide 5)
RC_Y = Emu(1026000)
RC_W = Emu(2862000)   # 7.95cm
RC_H = Emu(3500000)   # match left card height

# Wait — I created the left card already with wrong width. Let me redo.
# Actually I can't easily undo. Let me check if shapes were already added.
# Hmm, the script adds shapes to slide7 directly. The wrong-sized left card is already there.
# Easier approach: rewrite from scratch — delete slide7 and recreate.

# Delete the slide we just created
sld_lst = prs.slides._sldIdLst
slide_ids = list(sld_lst)
last_sld = slide_ids[-1]
sld_lst.remove(last_sld)
# drop rel
prs.part.drop_rel(last_sld.get(qn("r:id")))

print("Removed bad slide, recreating...")

# Create new slide
slide7 = prs.slides.add_slide(layout)
for ph in list(slide7.placeholders):
    sp = ph._element
    sp.getparent().remove(sp)
bg = slide7.background
bg.fill.solid()
bg.fill.fore_color.rgb = BG

# Header (correct EMU values — 1cm = 360000 EMU)
# Layout matches slide 5: title at y=1.45, divider at y=2.04, cards at y=2.67
# But we have a subtitle too, so we shift down by ~0.5cm
# Title: y=0.50cm h=0.94cm
add_text(slide7, Emu(140000), Emu(180000), Emu(8800000), Emu(340000),
         "Запуск новинок · Стандартизация процесса",
         size=20, bold=True, color=INK, font="Inter")
# Subtitle: y=1.55cm h=0.50cm
add_text(slide7, Emu(140000), Emu(560000), Emu(8800000), Emu(360000),
         "Единый стандарт вывода новинок и поддержки в период выхода на рынок",
         size=11, color=MUTE, font="Inter")
# Status pill at top-right (y=0.40cm)
PILL_X = Emu(7200000); PILL_Y = Emu(160000)   # 20.0cm, 0.44cm
PILL_W = Emu(1800000); PILL_H = Emu(280000)  # 5.0cm, 0.78cm
add_rect(slide7, PILL_X, PILL_Y, PILL_W, PILL_H,
         WHITE, line=SAND, line_w=9525, corner_radius=0.5)
add_text(slide7, PILL_X, PILL_Y+Emu(45000), PILL_W, PILL_H-Emu(45000),
         "СТАНДАРТ · ВНЕДРЕНИЕ", size=10, bold=True,
         color=SAND, font="Inter", align=PP_ALIGN.CENTER, spacing=150)
# Divider under header at y=2.20cm
add_rect(slide7, Emu(140000), Emu(800000), Emu(8855724), Emu(28575), CARD_BORDER)

# ════════════════════════════════════════════════════════════════════════
# LEFT CARD — "СТАНДАРТ ВЫВОДА" (4 этапа timeline)
# Position mirrors slide 5 left card: x=0.17, y=2.67, w=16.85, h=11.04 (but we use h=10.8 to fit slide height 14.29)
# ════════════════════════════════════════════════════════════════════════
LC_X = Emu(61200)     # 0.17cm
LC_Y = Emu(960000)     # 2.67cm
LC_W = Emu(6066960)   # 16.85cm
LC_H = Emu(3500000)   # 9.72cm

add_rect(slide7, LC_X, LC_Y, LC_W, LC_H,
         WHITE, line=CARD_BORDER, line_w=6350, corner_radius=0.02)

add_text(slide7, LC_X + Emu(200000), LC_Y + Emu(140000),
         Emu(8000000), Emu(220000),
         "СТАНДАРТ ВЫВОДА НОВИНКИ",
         size=11, bold=True, color=RED, font="Inter", spacing=200)
add_text(slide7, LC_X + Emu(200000), LC_Y + Emu(370000),
         Emu(8000000), Emu(180000),
         "4 обязательных этапа для каждой новинки",
         size=10, color=MUTE, font="Inter")
add_rect(slide7, LC_X + Emu(200000), LC_Y + Emu(580000),
         LC_W - Emu(400000), Emu(20000), DIVIDER)

# 4 timeline steps — horizontal layout
STAGES = [
    {
        "n": "01",
        "title": "ПОДГОТОВКА",
        "subtitle": "Анализ рынка · Ценообразование · MML",
        "duration": "−60 дней",
        "accent": RED,
    },
    {
        "n": "02",
        "title": "ЗАПУСК",
        "subtitle": "POSM · Обучение · Контент",
        "duration": "0 — 30 дней",
        "accent": SAND,
    },
    {
        "n": "03",
        "title": "СОПРОВОЖДЕНИЕ",
        "subtitle": "Промо · МК · Куратор · Аналитика",
        "duration": "30 — 90 дней",
        "accent": SAND,
    },
    {
        "n": "04",
        "title": "ОЦЕНКА",
        "subtitle": "KPI · Решение по масштабу",
        "duration": "90 — 180 дней",
        "accent": INK,
    },
]

STEP_GAP = Emu(120000)
STEP_PAD_X = Emu(200000)
STEP_W = (LC_W - 2 * STEP_PAD_X - 3 * STEP_GAP) // 4
STEP_H = Emu(2300000)
STEP_Y = LC_Y + Emu(750000)
STEP_X0 = LC_X + STEP_PAD_X

# Connecting line behind badges
LINE_Y = STEP_Y + Emu(280000)
add_rect(slide7, STEP_X0 + Emu(80000), LINE_Y,
         LC_W - 2 * STEP_PAD_X - Emu(160000), Emu(12000), DIVIDER)

for i, stage in enumerate(STAGES):
    cx = STEP_X0 + i * (STEP_W + STEP_GAP)

    # Number badge (circle)
    BADGE_SIZE = Emu(560000)
    add_rect(slide7, cx + (STEP_W - BADGE_SIZE) // 2,
             STEP_Y, BADGE_SIZE, BADGE_SIZE,
             WHITE, line=stage["accent"], line_w=12700, corner_radius=0.5)
    add_text(slide7, cx + (STEP_W - BADGE_SIZE) // 2,
             STEP_Y + Emu(110000),
             BADGE_SIZE, Emu(340000),
             stage["n"], size=14, bold=True, color=stage["accent"],
             font="Roboto Mono", align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)

    # Stage title (smaller font, no letter-spacing, to fit «СОПРОВОЖДЕНИЕ» on one line)
    add_text(slide7, cx, STEP_Y + Emu(680000),
             STEP_W, Emu(280000),
             stage["title"], size=9, bold=True, color=INK,
             font="Inter", align=PP_ALIGN.CENTER)

    # Duration
    add_text(slide7, cx, STEP_Y + Emu(980000),
             STEP_W, Emu(220000),
             stage["duration"], size=9, bold=True, color=stage["accent"],
             font="Roboto Mono", align=PP_ALIGN.CENTER)

    # Subtitle
    add_text(slide7, cx + Emu(40000), STEP_Y + Emu(1300000),
             STEP_W - Emu(80000), Emu(800000),
             stage["subtitle"], size=9.5, color=MUTE,
             font="Inter", align=PP_ALIGN.CENTER)

# Bottom note
add_rect(slide7, LC_X + Emu(200000), LC_Y + LC_H - Emu(550000),
         LC_W - Emu(400000), Emu(15000), DIVIDER)
add_text(slide7, LC_X + Emu(200000), LC_Y + LC_H - Emu(500000),
         LC_W - Emu(400000), Emu(380000),
         "Стандарт обязателен для всех новинок. Без прохождения 4 этапов новинка не вводится в MML.",
         size=9, color=MUTE, font="Inter", align=PP_ALIGN.CENTER)

# ════════════════════════════════════════════════════════════════════════
# RIGHT CARD — "ПОДДЕРЖКА В ПЕРИОД ВЫХОДА" (KPI metrics)
# Position mirrors slide 5 right card
# ════════════════════════════════════════════════════════════════════════
RC_X = Emu(6195600)
RC_Y = Emu(960000)   # match LC_Y
RC_W = Emu(2862000)
RC_H = Emu(3500000)  # match LC_H

add_rect(slide7, RC_X, RC_Y, RC_W, RC_H,
         WHITE, line=CARD_BORDER, line_w=6350, corner_radius=0.02)

# Header
add_text(slide7, RC_X + Emu(200000), RC_Y + Emu(140000),
         RC_W - Emu(400000), Emu(220000),
         "ПОДДЕРЖКА ВЫХОДА",
         size=11, bold=True, color=RED, font="Inter", spacing=200)
add_text(slide7, RC_X + Emu(200000), RC_Y + Emu(370000),
         RC_W - Emu(400000), Emu(180000),
         "Метрики периода вывода (90 дней)",
         size=10, color=MUTE, font="Inter")
add_rect(slide7, RC_X + Emu(200000), RC_Y + Emu(580000),
         RC_W - Emu(400000), Emu(20000), DIVIDER)

# 4 KPI rows — minimal, each is a number + label
KPIS = [
    ("60 дней",   "Подготовка до запуска",         RED),
    ("100%",      "Наличие POSM к старту",          SAND),
    ("350 ТП",    "Обучено в первый месяц",         SAND),
    ("90 дней",   "Период активной поддержки",      INK),
]

kpi_y = RC_Y + Emu(720000)
kpi_h = Emu(580000)
for i, (value, label, accent) in enumerate(KPIS):
    # Accent vertical bar on left
    add_rect(slide7, RC_X + Emu(200000), kpi_y + Emu(30000),
             Emu(40000), kpi_h - Emu(60000), accent)

    # Value (large, mono)
    add_text(slide7, RC_X + Emu(360000), kpi_y + Emu(40000),
             Emu(1600000), Emu(300000),
             value, size=16, bold=True, color=accent,
             font="Roboto Mono")

    # Label (smaller, muted, right of value)
    add_text(slide7, RC_X + Emu(200000), kpi_y + Emu(330000),
             RC_W - Emu(400000), Emu(200000),
             label, size=10, color=MUTE, font="Inter")

    kpi_y += kpi_h
    if i < len(KPIS) - 1:
        add_rect(slide7, RC_X + Emu(200000), kpi_y - Emu(80000),
                 RC_W - Emu(400000), Emu(12000), DIVIDER)

# Footer note inside right card
add_rect(slide7, RC_X + Emu(200000), RC_Y + RC_H - Emu(450000),
         RC_W - Emu(400000), Emu(15000), DIVIDER)
add_text(slide7, RC_X + Emu(200000), RC_Y + RC_H - Emu(400000),
         RC_W - Emu(400000), Emu(280000),
         "Без выполнения KPI поддержка не продлевается",
         size=9, color=MUTE, font="Inter", align=PP_ALIGN.CENTER)

# ── Save ───────────────────────────────────────────────────────────
prs.save(OUT)
print(f"\nSaved: {OUT}")
print(f"Total slides: {len(prs.slides)}")
