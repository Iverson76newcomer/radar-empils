"""
Redesign slide 7 "Запуск новинок":
- Keep only phases 01, 02, 03 (drop 04 — ОЦЕНКА)
- Move all 3 phases to the RIGHT half of the slide
- Leave LEFT half empty (for user to fill in 2026 launch results)
- Stack 3 phases vertically in the right half
"""
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import shutil, re

SRC = "/home/z/my-project/download/Презентация_для_ГД_final.pptx"
OUT = "/home/z/my-project/download/Презентация_для_ГД_final.pptx"

# Palette
BG          = RGBColor(0xF5, 0xF2, 0xEC)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
CARD_BORDER = RGBColor(0xD9, 0xD2, 0xC2)
DIVIDER     = RGBColor(0xEF, 0xE9, 0xDD)
INK         = RGBColor(0x2F, 0x3A, 0x48)
MUTE        = RGBColor(0x8A, 0x95, 0xA3)
RED         = RGBColor(0xC0, 0x00, 0x00)
SAND        = RGBColor(0xC9, 0xA7, 0x7C)
BLUE        = RGBColor(0x6B, 0x8F, 0xA8)
DASHED_BG   = RGBColor(0xEF, 0xE9, 0xDD)

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

def cm(v):
    return Emu(int(v * 360000))


# Open presentation
prs = Presentation(SRC)

# Find slide 7 (Запуск новинок) by content
target_idx = None
for i, slide in enumerate(prs.slides):
    for shape in slide.shapes:
        if shape.has_text_frame and 'Запуск новинок' in shape.text_frame.text:
            target_idx = i
            break
    if target_idx is not None:
        break

print(f"Found 'Запуск новинок' slide at index {target_idx} (slide {target_idx+1})")

slide_new = list(prs.slides)[target_idx]

# Remove all existing shapes from slide 7
sp_tree = slide_new.shapes._spTree
shapes_to_remove = []
for elem in sp_tree:
    if elem.tag.endswith('}sp') or elem.tag.endswith('}pic') or elem.tag.endswith('}grpSp') or elem.tag.endswith('}cxnSp'):
        shapes_to_remove.append(elem)
for elem in shapes_to_remove:
    sp_tree.remove(elem)
print(f"Removed {len(shapes_to_remove)} shapes")

# Reset background
bg = slide_new.background
bg.fill.solid()
bg.fill.fore_color.rgb = BG


# ════════════════════════════════════════════════════════════════════════
# HEADER (full-width, spans both halves)
# ════════════════════════════════════════════════════════════════════════
# Title
add_text(slide_new, cm(0.5), cm(0.5),
         cm(20.0), cm(0.95),
         "Запуск новинок · Стандарт вывода 2026",
         size=20, bold=True, color=INK, font="Inter")
# Subtitle
add_text(slide_new, cm(0.5), cm(1.55),
         cm(20.0), cm(0.5),
         "Левая часть — результаты запуска новинок 2026 · Правая — стандарт вывода",
         size=11, color=MUTE, font="Inter")
# Status pill
PILL_X = cm(20.0); PILL_Y = cm(0.45)
PILL_W = cm(5.0); PILL_H = cm(0.78)
add_rect(slide_new, PILL_X, PILL_Y, PILL_W, PILL_H,
         WHITE, line=SAND, line_w=9525, corner_radius=0.5)
add_text(slide_new, PILL_X, PILL_Y+cm(0.13), PILL_W, PILL_H-cm(0.13),
         "СТАНДАРТ · 2026", size=10, bold=True,
         color=SAND, font="Inter", align=PP_ALIGN.CENTER, spacing=150)
# Divider
add_rect(slide_new, cm(0.5), cm(2.22),
         cm(24.5), cm(0.08), CARD_BORDER)


# ════════════════════════════════════════════════════════════════════════
# LEFT HALF — empty placeholder for user to fill in launch results
# Position: x=0.5, y=2.67, w=11.7, h=10.7
# ════════════════════════════════════════════════════════════════════════
LEFT_X = cm(0.5)
LEFT_Y = cm(2.67)
LEFT_W = cm(11.7)
LEFT_H = cm(10.72)

# Light dashed/empty card (just border, no fill — for user to add content later)
add_rect(slide_new, LEFT_X, LEFT_Y, LEFT_W, LEFT_H,
         WHITE, line=CARD_BORDER, line_w=6350, corner_radius=0.02)

# Section header inside left card
add_text(slide_new, LEFT_X + cm(0.5), LEFT_Y + cm(0.4),
         cm(10.7), cm(0.5),
         "РЕЗУЛЬТАТЫ ЗАПУСКА НОВИНОК 2026",
         size=11, bold=True, color=RED, font="Inter", spacing=200)
add_text(slide_new, LEFT_X + cm(0.5), LEFT_Y + cm(0.95),
         cm(10.7), cm(0.4),
         "Заполняется по итогам года",
         size=10, color=MUTE, font="Inter")

# Divider
add_rect(slide_new, LEFT_X + cm(0.5), LEFT_Y + cm(1.5),
         cm(10.7), cm(0.06), DIVIDER)

# Placeholder text in the middle (light, italic, centered)
add_text(slide_new, LEFT_X + cm(0.5), LEFT_Y + cm(4.5),
         cm(10.7), cm(2.0),
         "Место для заполнения\n\nФактические показатели запуска:\n· TITAN грунт-эмаль 3в1\n· NEWTONE LOTUS\n· AVANTGARDE интерьерная",
         size=11, color=MUTE, font="Inter",
         align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)


# ════════════════════════════════════════════════════════════════════════
# RIGHT HALF — 3 phases stacked vertically
# Position: x=13.0, y=2.67, w=11.7, h=10.7
# ════════════════════════════════════════════════════════════════════════
RIGHT_X = cm(13.0)
RIGHT_Y = cm(2.67)
RIGHT_W = cm(11.7)
RIGHT_H = cm(10.72)

# Card background for right area
add_rect(slide_new, RIGHT_X, RIGHT_Y, RIGHT_W, RIGHT_H,
         WHITE, line=CARD_BORDER, line_w=6350, corner_radius=0.02)

# Header inside right card
add_text(slide_new, RIGHT_X + cm(0.4), RIGHT_Y + cm(0.4),
         cm(10.0), cm(0.5),
         "СТАНДАРТ ВЫВОДА · 3 ФАЗЫ",
         size=11, bold=True, color=RED, font="Inter", spacing=200)
add_text(slide_new, RIGHT_X + cm(0.4), RIGHT_Y + cm(0.95),
         cm(10.0), cm(0.4),
         "Подготовка → Запуск → Сопровождение",
         size=10, color=MUTE, font="Inter")
# Divider
add_rect(slide_new, RIGHT_X + cm(0.4), RIGHT_Y + cm(1.5),
         cm(10.9), cm(0.06), DIVIDER)

# 3 phase cards stacked vertically
# Each card: 10.9 cm wide × 2.7 cm tall, gap 0.2 cm
# Total: 3 × 2.7 + 2 × 0.2 = 8.5 cm (fits in 10.72 - 1.5 = 9.22 cm available)
PHASE_W = cm(10.9)
PHASE_H = cm(2.7)
PHASE_GAP = cm(0.2)
PHASE_X = RIGHT_X + cm(0.4)
PHASE_Y0 = RIGHT_Y + cm(1.85)

PHASES = [
    {
        "n": "01",
        "title": "ПОДГОТОВКА",
        "duration": "−60 до 0 дней",
        "bullets": ["Анализ рынка и конкурентов", "Ценообразование и MML", "Подготовка POSM и контента"],
        "kpi_value": "60 дней",
        "kpi_desc": "до запуска",
        "accent": RED,
    },
    {
        "n": "02",
        "title": "ЗАПУСК",
        "duration": "0 — 30 дней",
        "bullets": ["Старт продаж во всех пилотах", "Обучение ТП: 350 человек", "POSM в 100% пилотных точек"],
        "kpi_value": "100%",
        "kpi_desc": "наличие POSM",
        "accent": SAND,
    },
    {
        "n": "03",
        "title": "СОПРОВОЖДЕНИЕ",
        "duration": "30 — 90 дней",
        "bullets": ["Промоакции и мастер-классы", "Куратор на каждую точку", "Аналитика продаж по неделе"],
        "kpi_value": "350 ТП",
        "kpi_desc": "обучено в 1-й мес.",
        "accent": BLUE,
    },
]

for i, p in enumerate(PHASES):
    py = PHASE_Y0 + i * (PHASE_H + PHASE_GAP)

    # Card background
    add_rect(slide_new, PHASE_X, py, PHASE_W, PHASE_H,
             BG, line=CARD_BORDER, line_w=6350, corner_radius=0.04)

    # Left accent strip with number
    STRIP_W = cm(2.4)
    add_rect(slide_new, PHASE_X, py, STRIP_W, PHASE_H, WHITE, corner_radius=0.04)
    # Accent vertical bar
    add_rect(slide_new, PHASE_X, py, cm(0.11), PHASE_H, p['accent'])

    # Big number
    add_text(slide_new, PHASE_X, py + cm(0.5),
             STRIP_W, cm(1.1),
             p['n'], size=28, bold=True, color=p['accent'],
             font="Roboto Mono", align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)
    # Duration
    add_text(slide_new, PHASE_X, py + cm(1.7),
             STRIP_W, cm(0.4),
             p['duration'], size=9, bold=True, color=MUTE,
             font="Roboto Mono", align=PP_ALIGN.CENTER)
    # Phase label
    add_text(slide_new, PHASE_X, py + cm(2.15),
             STRIP_W, cm(0.35),
             "ФАЗА " + p['n'], size=7.5, bold=True, color=p['accent'],
             font="Inter", align=PP_ALIGN.CENTER, spacing=200)

    # Right content area
    CONTENT_X = PHASE_X + STRIP_W + cm(0.3)
    CONTENT_W = PHASE_W - STRIP_W - cm(0.6)

    # Title
    add_text(slide_new, CONTENT_X, py + cm(0.3),
             CONTENT_W, cm(0.45),
             p['title'], size=12, bold=True, color=INK,
             font="Inter", spacing=100)
    # Bullets (compact, 2 lines each)
    bullet_y = py + cm(0.95)
    for j, b in enumerate(p['bullets']):
        # Bullet dot
        add_text(slide_new, CONTENT_X, bullet_y,
                 cm(0.2), cm(0.32),
                 "·", size=12, bold=True, color=p['accent'],
                 font="Inter")
        # Bullet text
        add_text(slide_new, CONTENT_X + cm(0.25), bullet_y + cm(0.03),
                 CONTENT_W - cm(0.25), cm(0.32),
                 b, size=9, color=INK, font="Inter")
        bullet_y += cm(0.36)

    # KPI in bottom-right of card
    KPI_X = CONTENT_X + CONTENT_W - cm(3.0)
    KPI_Y = py + PHASE_H - cm(0.85)
    # KPI value
    add_text(slide_new, KPI_X, KPI_Y,
             cm(3.0), cm(0.4),
             p['kpi_value'], size=14, bold=True, color=p['accent'],
             font="Roboto Mono", align=PP_ALIGN.RIGHT)
    # KPI desc
    add_text(slide_new, KPI_X, KPI_Y + cm(0.4),
             cm(3.0), cm(0.3),
             p['kpi_desc'], size=8, color=MUTE,
             font="Inter", align=PP_ALIGN.RIGHT)

# Bottom note (full-width, below both halves)
add_rect(slide_new, cm(0.5), cm(13.5),
         cm(24.5), cm(0.04), DIVIDER)
add_text(slide_new, cm(0.5), cm(13.6),
         cm(24.5), cm(0.4),
         "Стандарт обязателен для всех новинок. Без прохождения 3 фаз новинка не вводится в MML дистрибьюторов.",
         size=9, color=MUTE, font="Inter", align=PP_ALIGN.CENTER)


# Save
prs.save(OUT)
print(f"\nSaved: {OUT}")
print(f"Slide 7 redesigned — 3 phases in right half, left half empty for results")
