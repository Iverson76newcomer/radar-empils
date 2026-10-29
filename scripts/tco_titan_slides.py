#!/usr/bin/env python3
"""
Создание 3 слайдов TCO для TITAN:
1) Ворота въездные
2) Гараж металлический
3) Дверь входная
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Cm, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import math

# ============================================================
# BRAND COLORS
# ============================================================
NAVY    = RGBColor(0x0A, 0x25, 0x40)
CYAN    = RGBColor(0x00, 0xB4, 0xD8)
ORANGE  = RGBColor(0xE8, 0x6C, 0x00)
GREEN   = RGBColor(0x1B, 0x7D, 0x46)
RED     = RGBColor(0xC0, 0x39, 0x2B)
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY  = RGBColor(0xF2, 0xF4, 0xF6)
MID_GRAY    = RGBColor(0x80, 0x88, 0x90)
DARK_GRAY   = RGBColor(0x3A, 0x3A, 0x3A)
LIGHT_GREEN = RGBColor(0xE8, 0xF5, 0xE9)
LIGHT_BLUE  = RGBColor(0xE3, 0xF2, 0xFD)
LIGHT_ORANGE = RGBColor(0xFF, 0xF3, 0xE0)
LIGHT_RED   = RGBColor(0xFF, 0xEB, 0xEE)

# TITAN brand colors
TITAN_RED    = RGBColor(0xC0, 0x39, 0x2B)
TITAN_DARK   = RGBColor(0x1A, 0x1A, 0x2E)
TITAN_SILVER = RGBColor(0xBD, 0xBD, 0xBD)

# ============================================================
# SLIDE DIMENSIONS (16:9)
# ============================================================
SLIDE_W = Emu(12192000)  # 33.867 cm
SLIDE_H = Emu(6858000)   # 19.05 cm

prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H

# Blank layout
blank_layout = prs.slide_layouts[6]

# ============================================================
# HELPER FUNCTIONS
# ============================================================
def add_shape(slide, left, top, width, height, fill_color=None, line_color=None, line_width=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    if fill_color:
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = line_width or Pt(1)
    else:
        shape.line.fill.background()
    # Smaller corner radius
    shape.adjustments[0] = 0.05
    return shape

def add_rect(slide, left, top, width, height, fill_color=None, line_color=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    if fill_color:
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
    else:
        shape.line.fill.background()
    return shape

def add_text(slide, left, top, width, height, text, font_size=12, bold=False, color=DARK_GRAY, align=PP_ALIGN.LEFT, font_name='Carlito'):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = font_name
    p.alignment = align
    return txBox

def add_bar(slide, left, top, max_width, value, max_value, fill_color, label="", cost_text="", period_text=""):
    """Draw a horizontal bar with label and cost"""
    bar_h = Cm(1.0)
    bar_width = int(max_width * (value / max_value)) if max_value > 0 else Cm(1)
    
    # Bar background (light)
    add_rect(slide, left, top, max_width, bar_h, fill_color=LIGHT_GRAY)
    # Bar fill
    add_rect(slide, left, top, bar_width, bar_h, fill_color=fill_color)
    
    # Cost label on bar
    if cost_text:
        add_text(slide, left + Cm(0.3), top - Cm(0.05), Cm(6), bar_h + Cm(0.1),
                 cost_text, font_size=11, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    
    # Period text after bar
    if period_text:
        add_text(slide, left + bar_width + Cm(0.3), top - Cm(0.05), Cm(5), bar_h + Cm(0.1),
                 period_text, font_size=10, bold=False, color=MID_GRAY, align=PP_ALIGN.LEFT)

def add_tco_slide(prs, title, subtitle, dimensions_text, area_text,
                  variants, obj_icon_text=""):
    """
    Add a TCO comparison slide.
    variants: list of dicts with keys:
      - name: str
      - material_cost: int (₽)
      - labor_cost: int (₽) — подготовка/работа
      - total_10yr: int (₽) — TCO за 10 лет
      - repaints: int — сколько перекрасок за 10 лет
      - bar_color: RGBColor
      - highlight: bool
    """
    slide = prs.slides.add_slide(blank_layout)
    
    # === BACKGROUND ===
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = WHITE
    
    # === TOP BAR (navy) ===
    add_rect(slide, 0, 0, SLIDE_W, Cm(1.2), fill_color=NAVY)
    
    # === TITLE ===
    add_text(slide, Cm(1.0), Cm(0.15), Cm(25), Cm(1.0),
             f"TCO: {title}", font_size=22, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    
    # === TITAN logo area ===
    add_text(slide, Cm(27), Cm(0.15), Cm(5), Cm(1.0),
             "TITAN", font_size=20, bold=True, color=ORANGE, align=PP_ALIGN.RIGHT)
    
    # === SUBTITLE / OBJECT INFO ===
    add_text(slide, Cm(1.0), Cm(1.5), Cm(15), Cm(0.7),
             subtitle, font_size=14, bold=True, color=NAVY, align=PP_ALIGN.LEFT)
    
    # Dimension + area badges
    badge_y = Cm(2.4)
    # Dimensions badge
    dim_shape = add_shape(slide, Cm(1.0), badge_y, Cm(8), Cm(0.9), 
                          fill_color=LIGHT_BLUE, line_color=CYAN)
    add_text(slide, Cm(1.3), badge_y + Cm(0.05), Cm(7.5), Cm(0.8),
             f"📏 {dimensions_text}", font_size=11, bold=True, color=NAVY, align=PP_ALIGN.LEFT)
    
    # Area badge
    area_shape = add_shape(slide, Cm(9.5), badge_y, Cm(6), Cm(0.9),
                           fill_color=LIGHT_ORANGE, line_color=ORANGE)
    add_text(slide, Cm(9.8), badge_y + Cm(0.05), Cm(5.5), Cm(0.8),
             f"📐 Площадь: {area_text}", font_size=11, bold=True, color=RGBColor(0x8B, 0x45, 0x00), align=PP_ALIGN.LEFT)
    
    # === 10-YEAR HEADER ===
    add_text(slide, Cm(1.0), Cm(3.6), Cm(15), Cm(0.6),
             "Совокупная стоимость владения за 10 лет (материал + работа)",
             font_size=10, bold=False, color=MID_GRAY, align=PP_ALIGN.LEFT)
    
    # === TCO BARS ===
    max_tco = max(v['total_10yr'] for v in variants)
    bar_left = Cm(9.5)
    bar_max_w = Cm(19)
    
    y_start = Cm(4.3)
    row_h = Cm(1.5)
    
    for i, v in enumerate(variants):
        y = y_start + i * row_h
        
        # Variant name
        name_color = GREEN if v.get('highlight') else DARK_GRAY
        add_text(slide, Cm(1.0), y - Cm(0.1), Cm(8.2), Cm(0.6),
                 v['name'], font_size=11, bold=True, color=name_color, align=PP_ALIGN.LEFT)
        
        # Details line
        details = f"Материал: {v['material_cost']:,}₽  |  Работа: {v['labor_cost']:,}₽  |  Перекрасок: {v['repaints']}"
        add_text(slide, Cm(1.0), y + Cm(0.45), Cm(8.2), Cm(0.5),
                 details, font_size=7.5, bold=False, color=MID_GRAY, align=PP_ALIGN.LEFT)
        
        # Bar
        bar_y = y + Cm(0.05)
        bar_h = Cm(0.9)
        bar_width = int(bar_max_w * (v['total_10yr'] / max_tco)) if max_tco > 0 else Cm(1)
        
        # Bar background
        add_rect(slide, bar_left, bar_y, bar_max_w, bar_h, fill_color=LIGHT_GRAY)
        # Bar fill
        bar_fill = v['bar_color']
        add_rect(slide, bar_left, bar_y, bar_width, bar_h, fill_color=bar_fill)
        
        # Cost on bar
        cost_str = f"{v['total_10yr']:,}₽"
        cost_color = WHITE if v['total_10yr'] / max_tco > 0.3 else DARK_GRAY
        # Place text inside bar if wide enough, else outside
        if v['total_10yr'] / max_tco > 0.5:
            text_x = bar_left + Cm(0.3)
            cost_color = WHITE
        else:
            text_x = bar_left + bar_width + Cm(0.3)
        add_text(slide, text_x, bar_y - Cm(0.05), Cm(4), bar_h + Cm(0.1),
                 cost_str, font_size=14, bold=True, color=cost_color, align=PP_ALIGN.LEFT)
    
    # === WINNER BOX ===
    winner = [v for v in variants if v.get('highlight')][0]
    savings = max_tco - winner['total_10yr']
    savings_pct = int(savings / max_tco * 100)
    
    # Large savings number on right side
    savings_box = add_shape(slide, Cm(23.5), Cm(3.6), Cm(7), Cm(3.0),
                            fill_color=LIGHT_GREEN, line_color=GREEN)
    add_text(slide, Cm(23.5), Cm(3.7), Cm(7), Cm(1.2),
             "ЭКОНОМИЯ", font_size=9, bold=True, color=GREEN, align=PP_ALIGN.CENTER)
    add_text(slide, Cm(23.5), Cm(4.6), Cm(7), Cm(1.4),
             f"{savings:,}₽", font_size=26, bold=True, color=GREEN, align=PP_ALIGN.CENTER)
    add_text(slide, Cm(23.5), Cm(5.8), Cm(7), Cm(0.8),
             f"{savings_pct}% за 10 лет", font_size=10, bold=False, color=GREEN, align=PP_ALIGN.CENTER)
    
    # Winner text bar
    win_y = Cm(9.2)
    win_shape = add_shape(slide, Cm(1.0), win_y, Cm(30), Cm(1.3),
                          fill_color=NAVY)
    
    add_text(slide, Cm(1.5), win_y + Cm(0.15), Cm(29), Cm(1.0),
             f"TITAN 3в1: красите 1 раз на 10 лет вместо {variants[0]['repaints']} раз — экономия {savings:,}₽ ({savings_pct}%)",
             font_size=13, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    
    # === BOTTOM BAR ===
    add_rect(slide, 0, Cm(11.5), SLIDE_W, Cm(0.3), fill_color=ORANGE)
    
    # Source note
    add_text(slide, Cm(1.0), Cm(11.9), Cm(30), Cm(0.5),
             "* Расчёт: TITAN 3в1 — 10 лет службы, ПФ-115+ГФ-021 — 4-5 лет, ПФ-115 ЭКОНОМ — 2-3 года. Работа = подготовка + покраска.",
             font_size=7, bold=False, color=MID_GRAY, align=PP_ALIGN.LEFT)


# ============================================================
# SLIDE 1: ВЪЕЗДНЫЕ ВОРОТА
# ============================================================
add_tco_slide(prs,
    title="ВЪЕЗДНЫЕ ВОРОТА",
    subtitle="Ворота распашные металлические (дача / гараж / дом)",
    dimensions_text="3.5 × 2.0 м (стандарт СНиП)",
    area_text="16 м² (обе стороны + стойки)",
    variants=[
        {
            'name': '❌ ПФ-115 ЭКОНОМ (без грунта)',
            'material_cost': 550,
            'labor_cost': 8000,
            'total_10yr': 8550,
            'repaints': 4,
            'bar_color': RED,
            'highlight': False,
        },
        {
            'name': '⚡ ПФ-115 + ГФ-021 (грунт + эмаль)',
            'material_cost': 1000,
            'labor_cost': 6000,
            'total_10yr': 7000,
            'repaints': 2,
            'bar_color': ORANGE,
            'highlight': False,
        },
        {
            'name': '✅ TITAN 3в1 (грунт+эмаль+защита, 10 лет)',
            'material_cost': 3500,
            'labor_cost': 2000,
            'total_10yr': 5500,
            'repaints': 1,
            'bar_color': GREEN,
            'highlight': True,
        },
    ],
)

# ============================================================
# SLIDE 2: ГАРАЖ МЕТАЛЛИЧЕСКИЙ
# ============================================================
add_tco_slide(prs,
    title="ГАРАЖ МЕТАЛЛИЧЕСКИЙ",
    subtitle="Гараж-ракушка / металлический пенал (типовой)",
    dimensions_text="3.0 × 6.0 × 2.0 м (стандарт)",
    area_text="54 м² (все стены + крыша)",
    variants=[
        {
            'name': '❌ ПФ-115 ЭКОНОМ (без грунта)',
            'material_cost': 1800,
            'labor_cost': 24000,
            'total_10yr': 25800,
            'repaints': 4,
            'bar_color': RED,
            'highlight': False,
        },
        {
            'name': '⚡ ПФ-115 + ГФ-021 (грунт + эмаль)',
            'material_cost': 3200,
            'labor_cost': 18000,
            'total_10yr': 21200,
            'repaints': 2,
            'bar_color': ORANGE,
            'highlight': False,
        },
        {
            'name': '✅ TITAN 3в1 (грунт+эмаль+защита, 10 лет)',
            'material_cost': 12000,
            'labor_cost': 6000,
            'total_10yr': 18000,
            'repaints': 1,
            'bar_color': GREEN,
            'highlight': True,
        },
    ],
)

# ============================================================
# SLIDE 3: ДВЕРЬ ВХОДНАЯ
# ============================================================
add_tco_slide(prs,
    title="ДВЕРЬ ВХОДНАЯ МЕТАЛЛИЧЕСКАЯ",
    subtitle="Входная дверь в дом / квартиру",
    dimensions_text="960 × 2050 мм (стандарт)",
    area_text="4.4 м² (оба полотна + рама)",
    variants=[
        {
            'name': '❌ ПФ-115 ЭКОНОМ (без грунта)',
            'material_cost': 150,
            'labor_cost': 3000,
            'total_10yr': 3150,
            'repaints': 4,
            'bar_color': RED,
            'highlight': False,
        },
        {
            'name': '⚡ ПФ-115 + ГФ-021 (грунт + эмаль)',
            'material_cost': 300,
            'labor_cost': 2000,
            'total_10yr': 2300,
            'repaints': 2,
            'bar_color': ORANGE,
            'highlight': False,
        },
        {
            'name': '✅ TITAN 3в1 (грунт+эмаль+защита, 10 лет)',
            'material_cost': 900,
            'labor_cost': 800,
            'total_10yr': 1700,
            'repaints': 1,
            'bar_color': GREEN,
            'highlight': True,
        },
    ],
)

# ============================================================
# SAVE
# ============================================================
output = '/home/z/my-project/download/TCO_TITAN_3_объекта.pptx'
prs.save(output)
print(f"Saved: {output}")
