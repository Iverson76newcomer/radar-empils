"""
Restructure Презентация_для_ГД_final.pptx:

1. Insert a NEW TITLE/COVER slide at the beginning (slide 1)
2. Replace the existing "Запуск новинок" slide (was slide 6, will become slide 7) with a
   NEW DESIGN that is visually different from the other 4-stage timeline slides.

The new "Запуск новинок" slide design uses a vertical 2×2 quadrant grid (NOT horizontal timeline)
with each quadrant showing one stage + its KPIs inline. This breaks the visual pattern of slides
5, 7 which use horizontal timelines + separate KPI card on the right.
"""
from pptx import Presentation
from pptx.util import Emu, Pt, Inches, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import shutil
import copy

SRC = "/home/z/my-project/upload/Презентация_для_ГД_final.pptx"
OUT = "/home/z/my-project/download/Презентация_для_ГД_final.pptx"
shutil.copy(SRC, OUT)

prs = Presentation(OUT)
print(f"Slide size: {prs.slide_width/360000:.2f} x {prs.slide_height/360000:.2f} cm")

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
AMBER       = RGBColor(0xD4, 0xA5, 0x74)
DARK_BG     = RGBColor(0x2F, 0x3A, 0x48)  # dark slate for cover

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


# ════════════════════════════════════════════════════════════════════════
# STEP 1: DELETE the existing "Запуск новинок" slide (slide 6 in 1-indexed)
# After deletion: total slides = 7, originally 8 minus 1
# Then we'll add new slide (so total back to 8)
# Then we'll insert a cover slide at the beginning
# ════════════════════════════════════════════════════════════════════════

# Find slide to delete — must identify it by content. Find index of slide with text "Запуск новинок"
target_idx = None
for i, slide in enumerate(prs.slides):
    for shape in slide.shapes:
        if shape.has_text_frame and 'Запуск новинок' in shape.text_frame.text:
            target_idx = i
            break
    if target_idx is not None:
        break

print(f"Found 'Запуск новинок' slide at index {target_idx} (1-indexed: {target_idx+1})")

# Delete that slide
sld_lst = prs.slides._sldIdLst
slide_ids = list(sld_lst)
sld_to_remove = slide_ids[target_idx]
prs.part.drop_rel(sld_to_remove.get(qn("r:id")))
sld_lst.remove(sld_to_remove)
print(f"Deleted slide at index {target_idx}")

# Verify
print(f"Remaining slides: {len(prs.slides)}")
for i, slide in enumerate(prs.slides):
    first_text = ''
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text_frame.text:
            first_text = shape.text_frame.text[:60]
            break
    print(f"  Slide {i+1}: {first_text!r}")


# ════════════════════════════════════════════════════════════════════════
# STEP 2: Add NEW "Запуск новинок" slide at the END (will reorder later)
# DESIGN: vertical 2×2 quadrant grid (4 stages as 4 cards, NOT horizontal timeline)
# Each quadrant has:
#   - Big number (01-04) on the LEFT
#   - Title + duration on the RIGHT
#   - 3 bullet points inside
#   - KPI footer at the bottom of each card
# This is structurally different from the other timeline slides.
# ════════════════════════════════════════════════════════════════════════

layout = prs.slide_layouts[0]
slide_new = prs.slides.add_slide(layout)
for ph in list(slide_new.placeholders):
    sp = ph._element
    sp.getparent().remove(sp)

bg = slide_new.background
bg.fill.solid()
bg.fill.fore_color.rgb = BG

# Header
add_text(slide_new, Emu(140000), Emu(180000), Emu(8800000), Emu(340000),
         "Запуск новинок · Стандарт вывода",
         size=20, bold=True, color=INK, font="Inter")
add_text(slide_new, Emu(140000), Emu(560000), Emu(8800000), Emu(360000),
         "Стандартизированный процесс для каждой новинки: 4 фазы, каждая со своим KPI",
         size=11, color=MUTE, font="Inter")
# Status pill
PILL_X = Emu(7200000); PILL_Y = Emu(160000)
PILL_W = Emu(1800000); PILL_H = Emu(280000)
add_rect(slide_new, PILL_X, PILL_Y, PILL_W, PILL_H,
         WHITE, line=SAND, line_w=9525, corner_radius=0.5)
add_text(slide_new, PILL_X, PILL_Y+Emu(45000), PILL_W, PILL_H-Emu(45000),
         "СТАНДАРТ · 2026", size=10, bold=True,
         color=SAND, font="Inter", align=PP_ALIGN.CENTER, spacing=150)
# Divider
add_rect(slide_new, Emu(140000), Emu(800000), Emu(8855724), Emu(28575), CARD_BORDER)

# ── 2×2 QUADRANT GRID ──────────────────────────────────────────────────
# Each quadrant: 12.0cm wide × 5.0cm tall, gap 0.4cm
# Total grid: 24.4 × 10.4 cm, starting at x=0.5, y=2.85

QUAD_W = Emu(4320000)  # 12.0 cm
QUAD_H = Emu(1800000)  # 5.0 cm
QUAD_GAP_X = Emu(150000)  # 0.42 cm
QUAD_GAP_Y = Emu(150000)

GRID_X0 = Emu(180000)   # 0.5 cm
GRID_Y0 = Emu(1026000)  # 2.85 cm

QUADRANTS = [
    {
        "n": "01",
        "title": "ПОДГОТОВКА",
        "duration": "−60 до 0 дней",
        "bullets": [
            "Анализ рынка и конкурентов",
            "Ценообразование и MML позиция",
            "Подготовка POSM и контента",
        ],
        "kpi_label": "KPI фазы",
        "kpi_value": "60 дней",
        "kpi_desc": "до запуска",
        "accent": RED,
    },
    {
        "n": "02",
        "title": "ЗАПУСК",
        "duration": "0 — 30 дней",
        "bullets": [
            "Старт продаж во всех пилотах",
            "Обучение ТП: 350 человек",
            "POSM в 100% пилотных точек",
        ],
        "kpi_label": "KPI фазы",
        "kpi_value": "100%",
        "kpi_desc": "наличие POSM",
        "accent": SAND,
    },
    {
        "n": "03",
        "title": "СОПРОВОЖДЕНИЕ",
        "duration": "30 — 90 дней",
        "bullets": [
            "Промоакции и мастер-классы",
            "Куратор на каждую точку",
            "Аналитика продаж по неделе",
        ],
        "kpi_label": "KPI фазы",
        "kpi_value": "350 ТП",
        "kpi_desc": "обучено в 1-й мес.",
        "accent": BLUE,
    },
    {
        "n": "04",
        "title": "ОЦЕНКА",
        "duration": "90 — 180 дней",
        "bullets": [
            "Анализ KPI по фазам 1-3",
            "Решение: масштабировать / корректировать",
            "Ввод в MML или вывод из ассортимента",
        ],
        "kpi_label": "KPI фазы",
        "kpi_value": "90 дней",
        "kpi_desc": "период поддержки",
        "accent": GREEN,
    },
]

for i, q in enumerate(QUADRANTS):
    col = i % 2
    row = i // 2
    qx = GRID_X0 + col * (QUAD_W + QUAD_GAP_X)
    qy = GRID_Y0 + row * (QUAD_H + QUAD_GAP_Y)

    # Card background
    add_rect(slide_new, qx, qy, QUAD_W, QUAD_H,
             WHITE, line=CARD_BORDER, line_w=6350, corner_radius=0.03)

    # Left accent strip with number
    STRIP_W = Emu(1100000)  # 3.05 cm
    add_rect(slide_new, qx, qy, STRIP_W, QUAD_H, BG, corner_radius=0.03)
    # Accent vertical bar on left edge
    add_rect(slide_new, qx, qy, Emu(40000), QUAD_H, q['accent'])

    # Big number
    add_text(slide_new, qx, qy + Emu(280000),
             STRIP_W, Emu(700000),
             q['n'], size=36, bold=True, color=q['accent'],
             font="Roboto Mono", align=PP_ALIGN.CENTER, va=MSO_ANCHOR.MIDDLE)

    # Duration (small, below number)
    add_text(slide_new, qx, qy + Emu(1000000),
             STRIP_W, Emu(280000),
             q['duration'], size=9, bold=True, color=MUTE,
             font="Roboto Mono", align=PP_ALIGN.CENTER)

    # Phase label (small, bottom)
    add_text(slide_new, qx, qy + Emu(1300000),
             STRIP_W, Emu(280000),
             "ФАЗА " + q['n'], size=8, bold=True, color=q['accent'],
             font="Inter", align=PP_ALIGN.CENTER, spacing=200)

    # Right content area: title + bullets + KPI
    CONTENT_X = qx + STRIP_W + Emu(180000)
    CONTENT_W = QUAD_W - STRIP_W - Emu(360000)

    # Title
    add_text(slide_new, CONTENT_X, qy + Emu(180000),
             CONTENT_W, Emu(280000),
             q['title'], size=12, bold=True, color=INK,
             font="Inter", spacing=100)

    # Bullets
    bullet_y = qy + Emu(540000)
    for j, b in enumerate(q['bullets']):
        # Bullet marker
        add_text(slide_new, CONTENT_X, bullet_y,
                 Emu(80000), Emu(220000),
                 "·", size=14, bold=True, color=q['accent'],
                 font="Inter")
        # Bullet text
        add_text(slide_new, CONTENT_X + Emu(100000), bullet_y + Emu(20000),
                 CONTENT_W - Emu(100000), Emu(220000),
                 b, size=9.5, color=INK, font="Inter")
        bullet_y += Emu(260000)

    # KPI footer at bottom of card
    KPI_FOOTER_Y = qy + QUAD_H - Emu(560000)
    add_rect(slide_new, CONTENT_X, KPI_FOOTER_Y,
             CONTENT_W, Emu(20000), DIVIDER)
    # KPI label (small)
    add_text(slide_new, CONTENT_X, KPI_FOOTER_Y + Emu(40000),
             Emu(1500000), Emu(180000),
             q['kpi_label'], size=8, bold=True, color=MUTE,
             font="Inter", spacing=100)
    # KPI value (large mono)
    add_text(slide_new, CONTENT_X, KPI_FOOTER_Y + Emu(220000),
             Emu(1500000), Emu(280000),
             q['kpi_value'], size=14, bold=True, color=q['accent'],
             font="Roboto Mono")
    # KPI description
    add_text(slide_new, CONTENT_X + Emu(1600000), KPI_FOOTER_Y + Emu(220000),
             CONTENT_W - Emu(1600000), Emu(280000),
             q['kpi_desc'], size=9, color=MUTE,
             font="Inter", va=MSO_ANCHOR.MIDDLE)

# Bottom note (full-width)
add_rect(slide_new, Emu(180000), Emu(4400000),
         Emu(8855724), Emu(15000), DIVIDER)
add_text(slide_new, Emu(180000), Emu(4450000),
         Emu(8855724), Emu(280000),
         "Стандарт обязателен для всех новинок Эмпилс. Без прохождения 4 фаз новинка не вводится в MML дистрибьюторов.",
         size=9, color=MUTE, font="Inter", align=PP_ALIGN.CENTER)

print(f"\nAdded new 'Запуск новинок' slide (vertical 2x2 grid design)")
print(f"Total slides now: {len(prs.slides)}")


# ════════════════════════════════════════════════════════════════════════
# STEP 3: CREATE COVER SLIDE and insert at the BEGINNING
# DESIGN: dark slate background, large brand title, subtitle, decorative
# ════════════════════════════════════════════════════════════════════════

slide_cover = prs.slides.add_slide(layout)
for ph in list(slide_cover.placeholders):
    sp = ph._element
    sp.getparent().remove(sp)

# Dark background
bg = slide_cover.background
bg.fill.solid()
bg.fill.fore_color.rgb = DARK_BG

# Decorative top accent line
add_rect(slide_cover, Emu(0), Emu(0),
         prs.slide_width, Emu(40000), RED)

# "TITAN" brand mark (top-left, small)
add_text(slide_cover, Emu(560000), Emu(560000),
         Emu(1500000), Emu(360000),
         "TITAN", size=14, bold=True, color=RED,
         font="Inter", spacing=400)
# Vertical divider
add_rect(slide_cover, Emu(2150000), Emu(580000),
         Emu(20000), Emu(320000), MUTE)
# Eyebrow
add_text(slide_cover, Emu(2300000), Emu(600000),
         Emu(7000000), Emu(280000),
         "ПРОГРАММА РАЗВИТИЯ · 2026", size=10, bold=True,
         color=SAND, font="Inter", spacing=300)

# Main title (centered vertically on the slide)
# Slide height = 14.29 cm, center y ≈ 7 cm
add_text(slide_cover, Emu(560000), Emu(3400000),
         Emu(15000000), Emu(1100000),
         "Дорога к лидерству",
         size=44, bold=True, color=WHITE,
         font="Inter", align=PP_ALIGN.LEFT)

# Subtitle
add_text(slide_cover, Emu(560000), Emu(4700000),
         Emu(15000000), Emu(500000),
         "Программа развития бренда ЭМПИЛС в сегменте ЛКМ",
         size=18, color=RGBColor(0xCC, 0xCC, 0xCC),
         font="Inter", align=PP_ALIGN.LEFT)

# Subsubtitle
add_text(slide_cover, Emu(560000), Emu(5350000),
         Emu(15000000), Emu(340000),
         "TITAN · ОРЕОЛ · NEWTONE · AVANTGARDE",
         size=12, color=MUTE,
         font="Inter", align=PP_ALIGN.LEFT, spacing=200)

# Bottom info row
add_rect(slide_cover, Emu(0), Emu(13100000),
         prs.slide_width, Emu(15000), RGBColor(0x44, 0x4A, 0x5A))

# Date + author + tagline
add_text(slide_cover, Emu(560000), Emu(13250000),
         Emu(4500000), Emu(280000),
         "Доклад для ГД · Октябрь 2026", size=11, bold=True,
         color=WHITE, font="Inter")
add_text(slide_cover, Emu(560000), Emu(13550000),
         Emu(4500000), Emu(220000),
         "ЗАО «Эмпилс» · Ростов-на-Дону", size=9, color=MUTE,
         font="Inter")

# Right-side tagline
add_text(slide_cover, Emu(7000000), Emu(13250000),
         Emu(8800000), Emu(280000),
         "Программа TITAN · Стандартизация · Рост",
         size=11, bold=True, color=SAND,
         font="Inter", align=PP_ALIGN.RIGHT, spacing=200)
add_text(slide_cover, Emu(7000000), Emu(13550000),
         Emu(8800000), Emu(220000),
         "8 слайдов · 5 ключевых проектов",
         size=9, color=MUTE, font="Inter", align=PP_ALIGN.RIGHT)

# Decorative accent bar (left of title, vertical)
add_rect(slide_cover, Emu(280000), Emu(3400000),
         Emu(80000), Emu(2300000), RED)

print(f"\nAdded cover slide")


# ════════════════════════════════════════════════════════════════════════
# STEP 4: REORDER slides — cover goes to position 0, new "Запуск новинок"
# goes to its logical place.
# Currently the order is:
#   [0] original slide 1 (TITAN ЧТО СДЕЛАНО)
#   [1] original slide 2 (сетевой канал)
#   [2] original slide 3 (TITAN продолжение)
#   [3] original slide 4 (Белые пятна)
#   [4] original slide 5 (Поддержка новых дистрибьюторов)
#   [5] original slide 7 (Выход в жилое строительство)
#   [6] original slide 8 (Развитие ВД-ЛКМ)
#   [7] NEW Запуск новинок (vertical 2x2)
#   [8] NEW Cover
# We want:
#   [0] NEW Cover
#   [1] TITAN ЧТО СДЕЛАНО (was original 1)
#   [2] сетевой канал (was original 2)
#   [3] TITAN продолжение (was original 3)
#   [4] Белые пятна (was original 4)
#   [5] Поддержка новых дистрибьюторов (was original 5)
#   [6] NEW Запуск новинок (was added at index 7)
#   [7] Выход в жилое строительство (was original 7)
#   [8] Развитие ВД-ЛКМ (was original 8)
# ════════════════════════════════════════════════════════════════════════

sld_lst = prs.slides._sldIdLst
slide_id_elements = list(sld_lst)
# Remove all from list
for el in slide_id_elements:
    sld_lst.remove(el)

# Insert in desired order
# slide_id_elements indices in original order (after our additions):
# 0: TITAN ЧТО СДЕЛАНО
# 1: сетевой канал
# 2: TITAN продолжение
# 3: Белые пятна
# 4: Поддержка новых дистрибьюторов
# 5: Выход в жилое строительство
# 6: Развитие ВД-ЛКМ
# 7: NEW Запуск новинок
# 8: NEW Cover

new_order = [
    slide_id_elements[8],  # Cover (NEW) → position 0
    slide_id_elements[0],  # TITAN ЧТО СДЕЛАНО → position 1
    slide_id_elements[1],  # сетевой канал → position 2
    slide_id_elements[2],  # TITAN продолжение → position 3
    slide_id_elements[3],  # Белые пятна → position 4
    slide_id_elements[4],  # Поддержка новых дистрибьюторов → position 5
    slide_id_elements[7],  # Запуск новинок (NEW) → position 6
    slide_id_elements[5],  # Выход в жилое строительство → position 7
    slide_id_elements[6],  # Развитие ВД-ЛКМ → position 8
]

for el in new_order:
    sld_lst.append(el)

# Save
prs.save(OUT)
print(f"\nSaved: {OUT}")
print(f"Total slides: {len(prs.slides)}")
print(f"\nFinal slide order:")
for i, slide in enumerate(prs.slides):
    first_text = ''
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text_frame.text:
            first_text = shape.text_frame.text[:60]
            break
    print(f"  Slide {i+1}: {first_text!r}")
