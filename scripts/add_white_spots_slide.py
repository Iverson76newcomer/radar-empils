"""
Add a 'Белые пятна' slide to the user's existing Презентация_для_ГД.pptx.
Keeps slide 1 (user's contrast dashboard) intact, appends slide 2 with:
- Dark theme matching user's slide 1 (#0A0A0A bg, #E63946 red, #FFD60A yellow)
- Geographic map of 5 target regions (pre-rendered PNG)
- Right panel: meetings/distributors list with dates
- Bottom KPI summary
"""
from pptx import Presentation
from pptx.util import Emu, Pt, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from copy import deepcopy
import shutil

# Paths
SRC = "/home/z/my-project/upload/Презентация_для_ГД.pptx"
OUT = "/home/z/my-project/download/Презентация_для_ГД.pptx"
MAP_PNG = "/home/z/my-project/download/slides_titan/assets/white_spots_map.png"

# Copy source as base
shutil.copy(SRC, OUT)

prs = Presentation(OUT)
SW, SH = prs.slide_width, prs.slide_height  # 9144000 x 5143500 EMU
print(f"Slide size: {SW/914400:.2f}\" x {SH/914400:.2f}\"  ({SW/360000:.2f} x {SH/360000:.2f} cm)")

# ── Helpers ────────────────────────────────────────────────────────────────
DARK_BG    = RGBColor(0x0A, 0x0A, 0x0A)
CARD_BG    = RGBColor(0x16, 0x16, 0x16)
CARD_BORDER= RGBColor(0x33, 0x33, 0x33)
RED        = RGBColor(0xE6, 0x39, 0x46)
YELLOW     = RGBColor(0xFF, 0xD6, 0x0A)
GREEN      = RGBColor(0x2E, 0xCC, 0x71)
ORANGE     = RGBColor(0xFF, 0x8C, 0x00)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
GRAY       = RGBColor(0x88, 0x88, 0x88)
LIGHT_GRAY = RGBColor(0xCC, 0xCC, 0xCC)
DARK_GRAY  = RGBColor(0x22, 0x22, 0x22)

def add_text_box(slide, x, y, w, h, text, *,
                 size=11, bold=False, color=WHITE, align=PP_ALIGN.LEFT,
                 va=MSO_ANCHOR.TOP, font="Inter", spacing=None,
                 line_height=None):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = 0; tf.margin_right = 0
    tf.margin_top = 0; tf.margin_bottom = 0
    tf.vertical_anchor = va
    p = tf.paragraphs[0]
    p.alignment = align
    if line_height:
        p.line_spacing = line_height
    r = p.add_run()
    r.text = text
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    if spacing:
        # letter spacing
        rPr = r._r.get_or_add_rPr()
        rPr.set('spc', str(spacing))
    return tb

def add_filled_rect(slide, x, y, w, h, fill_hex, *,
                    line_hex=None, line_w_emu=6350, corner_radius=None):
    shp_type = MSO_SHAPE.ROUNDED_RECTANGLE if corner_radius else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(shp_type, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill_hex
    if line_hex:
        shp.line.color.rgb = line_hex
        shp.line.width = Emu(line_w_emu)
    else:
        shp.line.fill.background()
    if corner_radius is not None:
        # set corner radius via adj
        try:
            shp.adjustments[0] = corner_radius
        except Exception:
            pass
    shp.shadow.inherit = False
    return shp

def add_pic(slide, path, x, y, w=None, h=None):
    return slide.shapes.add_picture(path, x, y, width=w, height=h)

# ── Create new blank slide (use first layout, will remove placeholders) ────
# User's pptx has only 1 layout (index 0); use it, then clear placeholders
blank_layout = prs.slide_layouts[0]
slide2 = prs.slides.add_slide(blank_layout)

# Remove any placeholder shapes inherited from layout
for ph in list(slide2.placeholders):
    sp = ph._element
    sp.getparent().remove(sp)

# Set slide background to dark
bg = slide2.background
fill = bg.fill
fill.solid()
fill.fore_color.rgb = DARK_BG

# ── HEADER ─────────────────────────────────────────────────────────────────
HX, HY = Emu(228600), Emu(170000)  # ~0.25cm x 0.19cm
# TITAN logo (red)
add_text_box(slide2, HX, HY, Emu(1200000), Emu(280000),
             "TITAN", size=18, bold=True, color=RED, font="Inter", spacing=200)
# Divider line
add_filled_rect(slide2, Emu(1080000), Emu(190000), Emu(28000), Emu(250000), DARK_GRAY)
# Subtitle eyebrow
add_text_box(slide2, Emu(1180000), Emu(175000), Emu(2200000), Emu(180000),
             "БЕЛЫЕ ПЯТНА · ПРОЕКТ ВХОДА", size=11, bold=True,
             color=LIGHT_GRAY, font="Inter", spacing=250)
# Main title
add_text_box(slide2, Emu(1180000), Emu(330000), Emu(4500000), Emu(320000),
             "География конкурентных дистрибьюторов", size=20, bold=True,
             color=WHITE, font="Inter")
add_text_box(slide2, Emu(1180000), Emu(620000), Emu(4500000), Emu(200000),
             "План входа в регионы, где продукция Эмпилс не представлена",
             size=10, color=GRAY, font="Inter")
# Status pill (right)
PILL_X = Emu(7600000); PILL_Y = Emu(190000)
add_filled_rect(slide2, PILL_X, PILL_Y, Emu(1300000), Emu(280000),
                CARD_BG, line_hex=YELLOW, line_w_emu=6350, corner_radius=0.5)
add_text_box(slide2, PILL_X, PILL_Y+Emu(40000), Emu(1300000), Emu(220000),
             "Q4 2026 · В РАБОТЕ", size=10, bold=True, color=YELLOW,
             font="Inter", align=PP_ALIGN.CENTER, spacing=150)

# Divider line under header
add_filled_rect(slide2, Emu(228600), Emu(950000), Emu(8686800), Emu(28000), DARK_GRAY)

# ── MAIN AREA: LEFT (map) + RIGHT (meetings list) ──────────────────────────
MAP_X = Emu(228600)
MAP_Y = Emu(1100000)
MAP_W = Emu(5800000)   # ~6.35cm width
MAP_H = Emu(3400000)   # height to fit

# Map card background
add_filled_rect(slide2, MAP_X - Emu(80000), MAP_Y - Emu(80000),
                MAP_W + Emu(160000), MAP_H + Emu(160000),
                CARD_BG, line_hex=CARD_BORDER, line_w_emu=6350, corner_radius=0.04)

# Insert map image
pic = add_pic(slide2, MAP_PNG, MAP_X, MAP_Y, w=MAP_W, h=MAP_H)

# ── RIGHT PANEL: meetings list ─────────────────────────────────────────────
LIST_X = Emu(6300000)
LIST_Y = Emu(1100000)
LIST_W = Emu(2700000)
LIST_H = Emu(3400000)

# Card bg
add_filled_rect(slide2, LIST_X - Emu(80000), LIST_Y - Emu(80000),
                LIST_W + Emu(160000), LIST_H + Emu(160000),
                CARD_BG, line_hex=CARD_BORDER, line_w_emu=6350, corner_radius=0.04)

# Header inside card
add_text_box(slide2, LIST_X, LIST_Y + Emu(80000), LIST_W, Emu(220000),
             "ГРАФИК ВСТРЕЧ", size=11, bold=True, color=YELLOW,
             font="Inter", spacing=200)
add_text_box(slide2, LIST_X, LIST_Y + Emu(290000), LIST_W, Emu(200000),
             "5 регионов · 5+ дистрибьюторов", size=9, color=GRAY, font="Inter")
# Divider
add_filled_rect(slide2, LIST_X, LIST_Y + Emu(490000), LIST_W, Emu(20000), DARK_GRAY)

# Meetings list
MEETINGS = [
    ("01", "Барнаул",          "Алтайский край",      "Олимп",                "13.10",         RED),
    ("02", "Красноярск",       "Красноярский край",   "ООО «ОЧАГ»",           "05–11.11",      YELLOW),
    ("03", "Улан-Удэ",         "Республика Бурятия",  "Вегос-М",              "19–25.11",      ORANGE),
    ("04", "Омск",             "Омская область",      "ИП Шишков",            "ноябрь",        ORANGE),
    ("05", "Казань",           "Республика Татарстан","ТС Агава, Карат, Колорит", "ноябрь",     YELLOW),
]

row_y = LIST_Y + Emu(560000)
row_h = Emu(560000)
for num, city, region, distr, date, accent in MEETINGS:
    # Number badge
    add_filled_rect(slide2, LIST_X, row_y + Emu(20000), Emu(320000), Emu(320000),
                    CARD_BORDER, corner_radius=0.18)
    add_text_box(slide2, LIST_X, row_y + Emu(70000), Emu(320000), Emu(240000),
                 num, size=12, bold=True, color=accent, font="Inter",
                 align=PP_ALIGN.CENTER)
    # City + region
    add_text_box(slide2, LIST_X + Emu(370000), row_y + Emu(20000), Emu(1600000), Emu(220000),
                 city, size=11, bold=True, color=WHITE, font="Inter")
    add_text_box(slide2, LIST_X + Emu(370000), row_y + Emu(220000), Emu(1600000), Emu(180000),
                 region, size=8.5, color=GRAY, font="Inter")
    # Distributor
    add_text_box(slide2, LIST_X + Emu(370000), row_y + Emu(380000), Emu(1700000), Emu(180000),
                 distr, size=9, color=accent, font="Inter", spacing=20)
    # Date (right column) — wider, smaller font, single line
    add_text_box(slide2, LIST_X + Emu(1950000), row_y + Emu(20000), Emu(720000), Emu(220000),
                 date, size=9, bold=True, color=WHITE, font="Roboto Mono",
                 align=PP_ALIGN.RIGHT)
    row_y += row_h
    # Divider line between rows (skip last)
    if num != "05":
        add_filled_rect(slide2, LIST_X, row_y - Emu(10000), LIST_W, Emu(12000), DARK_GRAY)

# ── BOTTOM: KPI row + caption ─────────────────────────────────────────────
KPI_Y = Emu(4650000)
KPI_H = Emu(380000)
KPI_GAP = Emu(120000)
KPI_W = (Emu(8686800) - 4 * KPI_GAP) // 5

KPIS = [
    ("5",      "регионов",       YELLOW),
    ("5+",     "дистрибьюторов", RED),
    ("Q4 2026","окно встреч",    ORANGE),
    ("3 нед.","средний визит",   YELLOW),
    ("100%",   "новых географий",RED),
]

kpi_x = Emu(228600)
for value, label, accent in KPIS:
    add_filled_rect(slide2, kpi_x, KPI_Y, KPI_W, KPI_H,
                    CARD_BG, line_hex=CARD_BORDER, line_w_emu=6350, corner_radius=0.08)
    # accent strip on left
    add_filled_rect(slide2, kpi_x, KPI_Y, Emu(40000), KPI_H, accent)
    # Value
    add_text_box(slide2, kpi_x + Emu(80000), KPI_Y + Emu(50000), KPI_W - Emu(80000), Emu(200000),
                 value, size=15, bold=True, color=accent, font="Roboto Mono")
    # Label
    add_text_box(slide2, kpi_x + Emu(80000), KPI_Y + Emu(230000), KPI_W - Emu(80000), Emu(140000),
                 label, size=8, color=GRAY, font="Inter", spacing=80)
    kpi_x += KPI_W + KPI_GAP

# Footer caption
add_text_box(slide2, Emu(228600), Emu(4990000), Emu(8686800), Emu(160000),
             "Стратегия: вход через действующих конкурентных дистрибьюторов в регионах без присутствия Эмпилс · Цель — закрытие «белых пятен» до конца 2026 года",
             size=8, color=GRAY, font="Inter")

# ── Save ──────────────────────────────────────────────────────────────────
prs.save(OUT)
print(f"Saved: {OUT}")
print(f"Total slides: {len(prs.slides)}")
