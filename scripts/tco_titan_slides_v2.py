#!/usr/bin/env python3
"""
Создание 6 слайдов TCO для TITAN (3 объекта × 2 слайда):
- Слайд А: Расчёт (методология, откуда цифры)
- Слайд Б: Результат TCO (бары сравнения)
"""
from pptx import Presentation
from pptx.util import Pt, Cm, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

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

SLIDE_W = Emu(12192000)
SLIDE_H = Emu(6858000)

prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H
blank_layout = prs.slide_layouts[6]

# ============================================================
# HELPERS
# ============================================================
def add_rect(slide, left, top, width, height, fill_color=None, line_color=None, line_w=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    if fill_color:
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = line_w or Pt(0.5)
    else:
        shape.line.fill.background()
    return shape

def add_rounded(slide, left, top, width, height, fill_color=None, line_color=None, line_w=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    if fill_color:
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = line_w or Pt(0.5)
    else:
        shape.line.fill.background()
    shape.adjustments[0] = 0.05
    return shape

def add_text(slide, left, top, width, height, text, size=10, bold=False, color=DARK_GRAY, align=PP_ALIGN.LEFT, anchor=None):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    if anchor:
        tf.paragraphs[0].space_before = Pt(0)
        tf.paragraphs[0].space_after = Pt(0)
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = 'Carlito'
    p.alignment = align
    return txBox

def add_multiline(slide, left, top, width, height, lines, size=9, color=DARK_GRAY, align=PP_ALIGN.LEFT, line_spacing=1.1):
    """Add textbox with multiple lines, each line = (text, bold, color_override)"""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, item in enumerate(lines):
        if isinstance(item, str):
            txt, bld, clr = item, False, color
        else:
            txt = item[0]
            bld = item[1] if len(item) > 1 else False
            clr = item[2] if len(item) > 2 else color
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = txt
        p.font.size = Pt(size)
        p.font.bold = bld
        p.font.color.rgb = clr
        p.font.name = 'Carlito'
        p.alignment = align
        p.space_before = Pt(1)
        p.space_after = Pt(1)
    return txBox

def add_top_bar(slide, title_text, subtitle_text=""):
    add_rect(slide, 0, 0, SLIDE_W, Cm(1.2), fill_color=NAVY)
    add_text(slide, Cm(1.0), Cm(0.15), Cm(25), Cm(1.0),
             title_text, size=20, bold=True, color=WHITE)
    add_text(slide, Cm(27), Cm(0.15), Cm(5), Cm(1.0),
             "TITAN", size=18, bold=True, color=ORANGE, align=PP_ALIGN.RIGHT)
    if subtitle_text:
        add_text(slide, Cm(1.0), Cm(1.4), Cm(30), Cm(0.6),
                 subtitle_text, size=12, bold=True, color=NAVY)
    # Bottom accent
    add_rect(slide, 0, Cm(11.5), SLIDE_W, Cm(0.3), fill_color=ORANGE)

def add_table(slide, left, top, rows_data, col_widths, header_color=NAVY, header_text_color=WHITE):
    """Add a styled table. rows_data = list of lists. First row = header."""
    n_rows = len(rows_data)
    n_cols = len(rows_data[0])
    table_width = sum(col_widths)
    table_height = Cm(n_rows * 0.65)
    
    tbl = slide.shapes.add_table(n_rows, n_cols, left, top, table_width, table_height).table
    
    # Set column widths
    for ci, w in enumerate(col_widths):
        tbl.columns[ci].width = w
    
    # Fill cells
    for ri, row in enumerate(rows_data):
        for ci, cell_text in enumerate(row):
            cell = tbl.cell(ri, ci)
            cell.text = str(cell_text)
            
            # Style
            for paragraph in cell.text_frame.paragraphs:
                paragraph.font.name = 'Carlito'
                paragraph.font.size = Pt(8)
                paragraph.alignment = PP_ALIGN.CENTER if ci > 0 else PP_ALIGN.LEFT
                
                if ri == 0:  # Header
                    paragraph.font.bold = True
                    paragraph.font.color.rgb = header_text_color
                    paragraph.font.size = Pt(8.5)
                elif ri == len(rows_data) - 1:  # Total/last row
                    paragraph.font.bold = True
                    paragraph.font.color.rgb = NAVY
                else:
                    paragraph.font.color.rgb = DARK_GRAY
            
            # Cell fill
            tcPr = cell._tc.get_or_add_tcPr()
            if ri == 0:
                solidFill = tcPr.makeelement(qn('a:solidFill'), {})
                srgbClr = solidFill.makeelement(qn('a:srgbClr'), {'val': '%02X%02X%02X' % (header_color[0], header_color[1], header_color[2])})
                solidFill.append(srgbClr)
                tcPr.append(solidFill)
            elif ri % 2 == 0:
                solidFill = tcPr.makeelement(qn('a:solidFill'), {})
                srgbClr = solidFill.makeelement(qn('a:srgbClr'), {'val': 'F2F4F6'})
                solidFill.append(srgbClr)
                tcPr.append(solidFill)
            
            # Margins
            cell.margin_left = Cm(0.15)
            cell.margin_right = Cm(0.15)
            cell.margin_top = Cm(0.05)
            cell.margin_bottom = Cm(0.05)
    
    return tbl


# ============================================================
# OBJECT DATA
# ============================================================
objects = [
    {
        'title': 'ВЪЕЗДНЫЕ ВОРОТА',
        'subtitle': 'Ворота распашные металлические (дача / гараж / дом)',
        'dimensions': '3.5 × 2.0 м (стандарт по СНиП 2.07.01-89)',
        'area_calc': '3.5 × 2.0 × 2 стороны + стойки ≈ 16 м²',
        'area': 16,
        'variants': [
            {
                'name': 'ПФ-115 ЭКОНОМ\n(без грунта)',
                'paint_name': 'Эмаль ПФ-115 ЭКОНОМ',
                'price_per_l': 190,       # ₽/л
                'consumption': 130,       # г/м² на слой
                'layers': 2,
                'prep': 'Зачистка + обезжиривание',
                'prep_cost': 2000,        # ₽ за цикл (подготовка)
                'service_life': 2.5,      # лет
                'repaints_10yr': 4,
                'bar_color': RED,
            },
            {
                'name': 'ПФ-115 + ГФ-021\n(грунт + эмаль)',
                'paint_name': 'ГФ-021 грунт + ПФ-115 КМ эмаль',
                'price_per_l': 280,       # ₽/л средняя
                'consumption': 110,       # г/м² (грунт 60 + эмаль 50)
                'layers': 3,              # 1 грунт + 2 эмаль
                'prep': 'Зачистка + обезжиривание + грунтование',
                'prep_cost': 3000,
                'service_life': 5,
                'repaints_10yr': 2,
                'bar_color': ORANGE,
            },
            {
                'name': 'TITAN 3в1\n(грунт+эмаль+защита)',
                'paint_name': 'Грунт-эмаль TITAN 3в1',
                'price_per_l': 1100,      # ₽/л (средняя по фасовкам)
                'consumption': 120,       # г/м² на слой
                'layers': 2,
                'prep': 'Очистка от рыхлой ржавчины (без грунтования!)',
                'prep_cost': 1000,
                'service_life': 10,
                'repaints_10yr': 1,
                'bar_color': GREEN,
                'highlight': True,
            },
        ],
    },
    {
        'title': 'ГАРАЖ МЕТАЛЛИЧЕСКИЙ',
        'subtitle': 'Гараж-ракушка / металлический пенал (типовой)',
        'dimensions': '3.0 × 6.0 × 2.0 м (стандарт)',
        'area_calc': '2×(3×2) + 2×(6×2) + 3×6 = 6+24+18 = 48 м² стен + крыша ≈ 54 м²',
        'area': 54,
        'variants': [
            {
                'name': 'ПФ-115 ЭКОНОМ\n(без грунта)',
                'paint_name': 'Эмаль ПФ-115 ЭКОНОМ',
                'price_per_l': 190,
                'consumption': 130,
                'layers': 2,
                'prep': 'Зачистка + обезжиривание',
                'prep_cost': 6000,
                'service_life': 2.5,
                'repaints_10yr': 4,
                'bar_color': RED,
            },
            {
                'name': 'ПФ-115 + ГФ-021\n(грунт + эмаль)',
                'paint_name': 'ГФ-021 грунт + ПФ-115 КМ эмаль',
                'price_per_l': 280,
                'consumption': 110,
                'layers': 3,
                'prep': 'Зачистка + обезжиривание + грунтование',
                'prep_cost': 9000,
                'service_life': 5,
                'repaints_10yr': 2,
                'bar_color': ORANGE,
            },
            {
                'name': 'TITAN 3в1\n(грунт+эмаль+защита)',
                'paint_name': 'Грунт-эмаль TITAN 3в1',
                'price_per_l': 1100,
                'consumption': 120,
                'layers': 2,
                'prep': 'Очистка от рыхлой ржавчины (без грунтования!)',
                'prep_cost': 3000,
                'service_life': 10,
                'repaints_10yr': 1,
                'bar_color': GREEN,
                'highlight': True,
            },
        ],
    },
    {
        'title': 'ДВЕРЬ ВХОДНАЯ МЕТАЛЛИЧЕСКАЯ',
        'subtitle': 'Входная дверь в дом / квартиру',
        'dimensions': '960 × 2050 мм (стандарт ГОСТ 31173-2003)',
        'area_calc': '0.9×2.05×2 стороны + рама + торцы ≈ 4.4 м²',
        'area': 4.4,
        'variants': [
            {
                'name': 'ПФ-115 ЭКОНОМ\n(без грунта)',
                'paint_name': 'Эмаль ПФ-115 ЭКОНОМ',
                'price_per_l': 190,
                'consumption': 130,
                'layers': 2,
                'prep': 'Зачистка + обезжиривание',
                'prep_cost': 750,
                'service_life': 2.5,
                'repaints_10yr': 4,
                'bar_color': RED,
            },
            {
                'name': 'ПФ-115 + ГФ-021\n(грунт + эмаль)',
                'paint_name': 'ГФ-021 грунт + ПФ-115 КМ эмаль',
                'price_per_l': 280,
                'consumption': 110,
                'layers': 3,
                'prep': 'Зачистка + обезжиривание + грунтование',
                'prep_cost': 1000,
                'service_life': 5,
                'repaints_10yr': 2,
                'bar_color': ORANGE,
            },
            {
                'name': 'TITAN 3в1\n(грунт+эмаль+защита)',
                'paint_name': 'Грунт-эмаль TITAN 3в1',
                'price_per_l': 1100,
                'consumption': 120,
                'layers': 2,
                'prep': 'Очистка от рыхлой ржавчины (без грунтования!)',
                'prep_cost': 400,
                'service_life': 10,
                'repaints_10yr': 1,
                'bar_color': GREEN,
                'highlight': True,
            },
        ],
    },
]


# ============================================================
# SLIDE A: РАСЧЁТ (методология)
# ============================================================
def add_calc_slide(obj):
    slide = prs.slides.add_slide(blank_layout)
    add_top_bar(slide, f"РАСЧЁТ TCO: {obj['title']}", obj['subtitle'])
    
    # === Object parameters box ===
    add_rounded(slide, Cm(0.8), Cm(2.1), Cm(14), Cm(1.6), 
                fill_color=LIGHT_BLUE, line_color=CYAN, line_w=Pt(1))
    add_multiline(slide, Cm(1.1), Cm(2.2), Cm(13.5), Cm(1.5), [
        (f"Объект: {obj['subtitle']}", True, NAVY),
        (f"Размеры: {obj['dimensions']}", False, DARK_GRAY),
        (f"Площадь окраски: {obj['area_calc']} = {obj['area']} м²", True, RGBColor(0x8B, 0x45, 0x00)),
    ], size=8.5)
    
    # === Calculation table for each variant ===
    area = obj['area']
    
    # Build table data
    rows = [["Параметр", "ПФ-115 ЭКОНОМ\n(без грунта)", "ПФ-115 + ГФ-021\n(грунт + эмаль)", "TITAN 3в1\n(грунт+эмаль+защита)"]]
    
    DENSITY = 1200  # г/л (плотность алкидной краски)
    for v in obj['variants']:
        # Расчёт объёма краски: площадь × расход × слои / плотность = литры
        liters_needed = round(area * v['consumption'] * v['layers'] / DENSITY, 1)
        v['total_liters'] = liters_needed
        mat_cost = round(liters_needed * v['price_per_l'])
        labor_paint = round(area * 50)  # 50₽/м² покраска
        labor_total = v['prep_cost'] + labor_paint
        total_cycle = mat_cost + labor_total
        v['mat_cost'] = mat_cost
        v['labor_total'] = labor_total
        v['total_cycle'] = total_cycle
        v['total_10yr'] = total_cycle * v['repaints_10yr']
    
    # Row: Paint type
    rows.append(["Краска", obj['variants'][0]['paint_name'], obj['variants'][1]['paint_name'], obj['variants'][2]['paint_name']])
    # Row: Price per liter
    rows.append(["Цена за 1 л", f"{obj['variants'][0]['price_per_l']}₽", f"{obj['variants'][1]['price_per_l']}₽", f"{obj['variants'][2]['price_per_l']}₽"])
    # Row: Consumption
    rows.append([f"Расход (г/м² × слой)", f"{obj['variants'][0]['consumption']} × {obj['variants'][0]['layers']} сл.", f"{obj['variants'][1]['consumption']} × {obj['variants'][1]['layers']} сл.", f"{obj['variants'][2]['consumption']} × {obj['variants'][2]['layers']} сл."])
    # Row: Total paint needed
    for v in obj['variants']:
        total_l = round(area * v['consumption'] * v['layers'] / 1000, 1)
        v['total_liters'] = total_l
    rows.append(["Краски на 1 цикл", f"{obj['variants'][0]['total_liters']} л", f"{obj['variants'][1]['total_liters']} л", f"{obj['variants'][2]['total_liters']} л"])
    # Row: Material cost per cycle
    rows.append(["Материал (1 цикл)", f"{obj['variants'][0]['mat_cost']:,}₽", f"{obj['variants'][1]['mat_cost']:,}₽", f"{obj['variants'][2]['mat_cost']:,}₽"])
    # Row: Preparation
    rows.append(["Подготовка поверхности", obj['variants'][0]['prep'], obj['variants'][1]['prep'], obj['variants'][2]['prep']])
    # Row: Labor cost per cycle
    rows.append(["Работа (1 цикл)", f"{obj['variants'][0]['labor_total']:,}₽", f"{obj['variants'][1]['labor_total']:,}₽", f"{obj['variants'][2]['labor_total']:,}₽"])
    # Row: Service life
    rows.append(["Срок службы покрытия", f"{obj['variants'][0]['service_life']} года", f"{obj['variants'][1]['service_life']} лет", f"{obj['variants'][2]['service_life']} лет"])
    # Row: Repaints in 10 years
    rows.append(["Перекрасок за 10 лет", str(obj['variants'][0]['repaints_10yr']), str(obj['variants'][1]['repaints_10yr']), str(obj['variants'][2]['repaints_10yr'])])
    # Row: Total TCO
    rows.append(["ИТОГО TCO за 10 лет",
                 f"{obj['variants'][0]['total_10yr']:,}₽",
                 f"{obj['variants'][1]['total_10yr']:,}₽",
                 f"{obj['variants'][2]['total_10yr']:,}₽"])
    
    # Position table
    col_w = [Cm(4.5), Cm(4.5), Cm(5.0), Cm(5.0)]
    add_table(slide, Cm(0.8), Cm(3.9), rows, col_w, header_color=NAVY)
    
    # === Key insight box (right side) ===
    winner = obj['variants'][2]
    loser = obj['variants'][0]
    savings = loser['total_10yr'] - winner['total_10yr']
    savings_pct = int(savings / loser['total_10yr'] * 100)
    
    add_rounded(slide, Cm(19.8), Cm(2.1), Cm(12.5), Cm(2.8),
                fill_color=LIGHT_GREEN, line_color=GREEN, line_w=Pt(1.5))
    add_multiline(slide, Cm(20.2), Cm(2.3), Cm(11.8), Cm(2.5), [
        ("КЛЮЧЕВОЙ ВЫВОД", True, GREEN),
        (f"TITAN 3в1 экономит {savings:,}₽ за 10 лет", True, NAVY),
        (f"({savings_pct}% дешевле) — красите 1 раз вместо {loser['repaints_10yr']}", False, DARK_GRAY),
        ("", False, DARK_GRAY),
        ("Почему: срок службы 10 лет vs 2-5 лет,", False, MID_GRAY),
        ("не нужна грунтовка, наносится на ржавчину", False, MID_GRAY),
    ], size=8.5)
    
    # === Labor breakdown box ===
    add_rounded(slide, Cm(19.8), Cm(5.2), Cm(12.5), Cm(3.0),
                fill_color=LIGHT_ORANGE, line_color=ORANGE, line_w=Pt(1))
    add_multiline(slide, Cm(20.2), Cm(5.4), Cm(11.8), Cm(2.8), [
        ("Стоимость работы за 1 цикл:", True, ORANGE),
        (f"  Подготовка: зачистка, обезжиривание, грунтование", False, DARK_GRAY),
        (f"  Покраска: ~50₽/м² × {area} м² = {round(area*50):,}₽", False, DARK_GRAY),
        ("", False, DARK_GRAY),
        ("TITAN: подготовка проще — не нужно", False, GREEN),
        ("грунтовать, можно наносить на ржавчину!", False, GREEN),
    ], size=8)
    
    # Source footnote
    add_text(slide, Cm(0.8), Cm(11.9), Cm(30), Cm(0.4),
             "* Цена TITAN: 765₽/0.75л, 1 720₽/2л, 5 720₽/8.5л (Все Инструменты, 2026). ПФ-115 ЭКОНОМ: ~190₽/л. ГФ-021: ~145₽/л. ПФ-115 КМ: ~280₽/л.",
             size=6.5, color=MID_GRAY)


# ============================================================
# SLIDE Б: РЕЗУЛЬТАТ TCO (бары сравнения)
# ============================================================
def add_result_slide(obj):
    slide = prs.slides.add_slide(blank_layout)
    add_top_bar(slide, f"TCO: {obj['title']}", obj['subtitle'])
    
    # Object info badges
    badge_y = Cm(2.2)
    add_rounded(slide, Cm(0.8), badge_y, Cm(8), Cm(0.85), fill_color=LIGHT_BLUE, line_color=CYAN)
    add_text(slide, Cm(1.1), badge_y + Cm(0.05), Cm(7.5), Cm(0.7),
             obj['dimensions'], size=9, bold=True, color=NAVY)
    
    add_rounded(slide, Cm(9.2), badge_y, Cm(8), Cm(0.85), fill_color=LIGHT_ORANGE, line_color=ORANGE)
    add_text(slide, Cm(9.5), badge_y + Cm(0.05), Cm(7.5), Cm(0.7),
             f"Площадь: {obj['area']} м²  |  {obj['area_calc'].split('=')[-1].strip() if '=' in obj['area_calc'] else str(obj['area'])+' м²'}", 
             size=9, bold=True, color=RGBColor(0x8B, 0x45, 0x00))
    
    # Header
    add_text(slide, Cm(0.8), Cm(3.4), Cm(15), Cm(0.5),
             "Совокупная стоимость владения за 10 лет (материал + работа)",
             size=9, color=MID_GRAY)
    
    # === BARS ===
    max_tco = max(v['total_10yr'] for v in obj['variants'])
    bar_left = Cm(10.0)
    bar_max_w = Cm(15.5)
    y_start = Cm(4.0)
    row_h = Cm(1.6)
    
    for i, v in enumerate(obj['variants']):
        y = y_start + i * row_h
        name_clean = v['name'].replace('\n', ' ')
        name_color = GREEN if v.get('highlight') else DARK_GRAY
        
        add_text(slide, Cm(0.8), y - Cm(0.1), Cm(9), Cm(0.55),
                 name_clean, size=11, bold=True, color=name_color)
        add_text(slide, Cm(0.8), y + Cm(0.4), Cm(9), Cm(0.5),
                 f"Материал: {v['mat_cost']:,}₽/цикл  |  Работа: {v['labor_total']:,}₽/цикл  |  Перекрасок: {v['repaints_10yr']}",
                 size=7.5, color=MID_GRAY)
        
        bar_y = y + Cm(0.05)
        bar_h = Cm(0.85)
        bar_width = int(bar_max_w * (v['total_10yr'] / max_tco)) if max_tco > 0 else Cm(1)
        
        add_rect(slide, bar_left, bar_y, bar_max_w, bar_h, fill_color=LIGHT_GRAY)
        add_rect(slide, bar_left, bar_y, bar_width, bar_h, fill_color=v['bar_color'])
        
        cost_str = f"{v['total_10yr']:,}₽"
        if v['total_10yr'] / max_tco > 0.5:
            add_text(slide, bar_left + Cm(0.3), bar_y - Cm(0.05), Cm(4), bar_h + Cm(0.1),
                     cost_str, size=14, bold=True, color=WHITE)
        else:
            add_text(slide, bar_left + bar_width + Cm(0.3), bar_y - Cm(0.05), Cm(4), bar_h + Cm(0.1),
                     cost_str, size=14, bold=True, color=DARK_GRAY)
    
    # === SAVINGS BOX ===
    winner = [v for v in obj['variants'] if v.get('highlight')][0]
    loser = obj['variants'][0]
    savings = loser['total_10yr'] - winner['total_10yr']
    savings_pct = int(savings / loser['total_10yr'] * 100)
    
    add_rounded(slide, Cm(23), Cm(3.5), Cm(8.5), Cm(3.2),
                fill_color=LIGHT_GREEN, line_color=GREEN, line_w=Pt(2))
    add_text(slide, Cm(23), Cm(3.6), Cm(8.5), Cm(0.6),
             "ЭКОНОМИЯ ЗА 10 ЛЕТ", size=9, bold=True, color=GREEN, align=PP_ALIGN.CENTER)
    add_text(slide, Cm(23), Cm(4.3), Cm(8.5), Cm(1.4),
             f"{savings:,}₽", size=28, bold=True, color=GREEN, align=PP_ALIGN.CENTER)
    add_text(slide, Cm(23), Cm(5.6), Cm(8.5), Cm(0.6),
             f"{savings_pct}% экономии", size=11, bold=False, color=GREEN, align=PP_ALIGN.CENTER)
    
    # === WINNER BAR ===
    win_y = Cm(9.2)
    add_rounded(slide, Cm(0.8), win_y, Cm(30.5), Cm(1.3), fill_color=NAVY)
    add_text(slide, Cm(1.3), win_y + Cm(0.15), Cm(29.5), Cm(1.0),
             f"TITAN 3в1: красите 1 раз на 10 лет вместо {loser['repaints_10yr']} раз  —  экономия {savings:,}₽ ({savings_pct}%)",
             size=12, bold=True, color=WHITE)
    
    # Footnote
    add_text(slide, Cm(0.8), Cm(11.9), Cm(30), Cm(0.4),
             "* Подробнее см. слайд «Расчёт TCO»",
             size=6.5, color=MID_GRAY)


# ============================================================
# BUILD ALL SLIDES
# ============================================================
for obj in objects:
    # First: calculation slide
    add_calc_slide(obj)
    # Second: result slide
    add_result_slide(obj)

# ============================================================
# SAVE
# ============================================================
output = '/home/z/my-project/download/TCO_TITAN_3_объекта.pptx'
prs.save(output)
print(f"Saved: {output}")

# Print summary
for obj in objects:
    print(f"\n=== {obj['title']} ({obj['area']} м²) ===")
    for v in obj['variants']:
        name = v['name'].replace('\n', ' ')
        print(f"  {name}: мат={v.get('mat_cost',0):,}₽ раб={v.get('labor_total',0):,}₽ цикл={v.get('total_cycle',0):,}₽ × {v['repaints_10yr']} = {v.get('total_10yr',0):,}₽/10лет")
