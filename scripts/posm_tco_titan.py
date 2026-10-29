#!/usr/bin/env python3
"""
POSM-карта TCO для TITAN — плакат А3 для размещения у стеллажа с краской.
Дизайн: «Карта расчёта TCO» — покупатель за 10 секунд понимает выгоду.
Формат: А3 (420×297 мм), ландшафт — для размещения на полке/стенде.
"""
from reportlab.lib.pagesizes import A3, landscape
from reportlab.lib.units import mm, cm
from reportlab.lib.colors import HexColor, white, black
from reportlab.pdfgen import canvas
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.platypus import Paragraph, Frame
from reportlab.lib.styles import ParagraphStyle

# ============================================================
# COLORS
# ============================================================
NAVY    = HexColor('#0A2540')
CYAN    = HexColor('#00B4D8')
ORANGE  = HexColor('#E86C00')
GREEN   = HexColor('#1B7D46')
RED     = HexColor('#C0392B')
WHITE   = HexColor('#FFFFFF')
LGRAY   = HexColor('#F2F4F6')
MGRAY   = HexColor('#808890')
DGRAY   = HexColor('#3A3A3A')
LGREEN  = HexColor('#E8F5E9')
LORANGE = HexColor('#FFF3E0')
LRED    = HexColor('#FFEBEE')
LBLUE   = HexColor('#E3F2FD')

# Page: A3 landscape
W, H = landscape(A3)  # 420mm × 297mm → 1190.55 × 841.89 pt

output = '/home/z/my-project/download/POSM_TCO_TITAN_A3.pdf'
c = canvas.Canvas(output, pagesize=landscape(A3))

# ============================================================
# FONTS
# ============================================================
FONT = 'LiberationSans'
FONT_BOLD = 'LiberationSans-Bold'

# Register fonts
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
pdfmetrics.registerFont(TTFont('LiberationSans', '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf'))
pdfmetrics.registerFont(TTFont('LiberationSans-Bold', '/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf'))
# NotoSansSC variable font doesn't work with reportlab — skip, use LiberationSans for Cyrillic

# Fallback: use Liberation Sans for everything (good Cyrillic support)
FONT = 'LiberationSans'
FONT_BOLD = 'LiberationSans-Bold'
FONT_CN = 'LiberationSans'  # fallback for CJK, LiberationSans has basic coverage

# ============================================================
# HELPER FUNCTIONS
# ============================================================
def draw_rounded_rect(c, x, y, w, h, r=3*mm, fill_color=None, stroke_color=None, stroke_width=0.5):
    """Draw a rounded rectangle"""
    c.saveState()
    if fill_color:
        c.setFillColor(fill_color)
    if stroke_color:
        c.setStrokeColor(stroke_color)
        c.setLineWidth(stroke_width)
    else:
        c.setStrokeColor(fill_color if fill_color else white)
        c.setLineWidth(0)
    c.roundRect(x, y, w, h, r, fill=1 if fill_color else 0, stroke=1 if stroke_color else 0)
    c.restoreState()

def draw_text(c, x, y, text, size=10, bold=False, color=DGRAY, align='left', max_width=None):
    """Draw text at position"""
    c.saveState()
    font = FONT_BOLD if bold else FONT
    c.setFont(font, size)
    c.setFillColor(color)
    
    if align == 'center':
        c.drawCentredString(x, y, text)
    elif align == 'right':
        c.drawRightString(x, y, text)
    else:
        c.drawString(x, y, text)
    c.restoreState()

def draw_multiline(c, x, y, lines, size=9, color=DGRAY, line_h=None, bold_first=False):
    """Draw multiple lines, return total height used"""
    lh = line_h or size * 1.4
    total = 0
    for i, line in enumerate(lines):
        is_bold = (i == 0 and bold_first) or (isinstance(line, tuple) and len(line) > 1 and line[1])
        clr = color if not isinstance(line, tuple) or len(line) < 3 else line[2]
        txt = line if isinstance(line, str) else line[0]
        draw_text(c, x, y - total, txt, size=size, bold=is_bold, color=clr)
        total += lh
    return total

# ============================================================
# LAYOUT CONSTANTS
# ============================================================
MARGIN = 12*mm
COL_GAP = 8*mm

# Top header height
HEADER_H = 22*mm
# Object cards area
CARD_TOP = H - HEADER_H - MARGIN - 5*mm
CARD_BOTTOM = MARGIN + 32*mm  # space for bottom bar
CARD_HEIGHT = CARD_TOP - CARD_BOTTOM

# 3 columns for 3 objects
TOTAL_W = W - 2*MARGIN
COL_W = (TOTAL_W - 2*COL_GAP) / 3

# ============================================================
# DRAW: HEADER
# ============================================================
# Navy header bar
c.setFillColor(NAVY)
c.rect(0, H - HEADER_H, W, HEADER_H, fill=1, stroke=0)

# Orange accent line under header
c.setFillColor(ORANGE)
c.rect(0, H - HEADER_H - 2*mm, W, 2*mm, fill=1, stroke=0)

# Title text
draw_text(c, MARGIN, H - 14*mm, "TCO КАРТА: СКОЛЬКО СТОИТ ПОКРАСКА ЗА 10 ЛЕТ?",
          size=18, bold=True, color=WHITE)

# TITAN logo
draw_text(c, W - MARGIN, H - 14*mm, "TITAN", size=20, bold=True, color=ORANGE, align='right')

# Subtitle
draw_text(c, MARGIN, H - HEADER_H + 1*mm, 
          "Совокупная стоимость владения: материал + работа + перекраски. Красите 1 раз вместо 4!",
          size=9, color=HexColor('#B0C4D8'))

# ============================================================
# DATA: 3 OBJECTS
# ============================================================
objects = [
    {
        'icon': '[ВОРОТА]',
        'title': 'ВОРОТА',
        'subtitle': 'Распашные, 3.5×2.0 м',
        'area': '16 м²',
        'variants': [
            {'name': 'ПФ-115 ЭКОНОМ', 'color': RED, 'bg': LRED,
             'material': '665₽', 'labor': '2 800₽', 'cycle': '3 465₽',
             'life': '2.5 года', 'repaints': '4', 'tco': '13 860₽'},
            {'name': 'ПФ-115 + ГФ-021', 'color': ORANGE, 'bg': LORANGE,
             'material': '1 232₽', 'labor': '3 800₽', 'cycle': '5 032₽',
             'life': '5 лет', 'repaints': '2', 'tco': '10 064₽'},
            {'name': 'TITAN 3в1', 'color': GREEN, 'bg': LGREEN,
             'material': '3 520₽', 'labor': '1 800₽', 'cycle': '5 320₽',
             'life': '10 лет', 'repaints': '1', 'tco': '5 320₽',
             'winner': True},
        ],
        'savings': '8 540₽',
        'savings_pct': '61%',
    },
    {
        'icon': '🏠',
        'title': 'ГАРАЖ',
        'subtitle': 'Ракушка, 3×6×2 м',
        'area': '54 м²',
        'variants': [
            {'name': 'ПФ-115 ЭКОНОМ', 'color': RED, 'bg': LRED,
             'material': '2 223₽', 'labor': '8 700₽', 'cycle': '10 923₽',
             'life': '2.5 года', 'repaints': '4', 'tco': '43 692₽'},
            {'name': 'ПФ-115 + ГФ-021', 'color': ORANGE, 'bg': LORANGE,
             'material': '4 144₽', 'labor': '11 700₽', 'cycle': '15 844₽',
             'life': '5 лет', 'repaints': '2', 'tco': '31 688₽'},
            {'name': 'TITAN 3в1', 'color': GREEN, 'bg': LGREEN,
             'material': '11 880₽', 'labor': '5 700₽', 'cycle': '17 580₽',
             'life': '10 лет', 'repaints': '1', 'tco': '17 580₽',
             'winner': True},
        ],
        'savings': '26 112₽',
        'savings_pct': '59%',
    },
    {
        'icon': '🚪',
        'title': 'ДВЕРЬ',
        'subtitle': 'Входная, 960×2050 мм',
        'area': '4.4 м²',
        'variants': [
            {'name': 'ПФ-115 ЭКОНОМ', 'color': RED, 'bg': LRED,
             'material': '190₽', 'labor': '970₽', 'cycle': '1 160₽',
             'life': '2.5 года', 'repaints': '4', 'tco': '4 640₽'},
            {'name': 'ПФ-115 + ГФ-021', 'color': ORANGE, 'bg': LORANGE,
             'material': '336₽', 'labor': '1 220₽', 'cycle': '1 556₽',
             'life': '5 лет', 'repaints': '2', 'tco': '3 112₽'},
            {'name': 'TITAN 3в1', 'color': GREEN, 'bg': LGREEN,
             'material': '990₽', 'labor': '620₽', 'cycle': '1 610₽',
             'life': '10 лет', 'repaints': '1', 'tco': '1 610₽',
             'winner': True},
        ],
        'savings': '3 030₽',
        'savings_pct': '65%',
    },
]

# ============================================================
# DRAW: 3 OBJECT CARDS
# ============================================================
for i, obj in enumerate(objects):
    x = MARGIN + i * (COL_W + COL_GAP)
    y = CARD_BOTTOM
    w = COL_W
    h = CARD_HEIGHT
    
    # === Card background ===
    draw_rounded_rect(c, x, y, w, h, r=4*mm, fill_color=HexColor('#FAFBFC'), stroke_color=HexColor('#D0D8E0'), stroke_width=1)
    
    # === Object title bar ===
    title_bar_h = 16*mm
    draw_rounded_rect(c, x, y + h - title_bar_h, w, title_bar_h, r=4*mm, fill_color=NAVY)
    # Cover bottom corners of title bar
    c.setFillColor(NAVY)
    c.rect(x, y + h - title_bar_h, w, 4*mm, fill=1, stroke=0)
    
    # Object name
    draw_text(c, x + 5*mm, y + h - 11*mm, obj['title'], size=16, bold=True, color=WHITE)
    draw_text(c, x + 5*mm, y + h - 15*mm, f"{obj['subtitle']}  |  {obj['area']}", 
              size=8, color=HexColor('#B0C4D8'))
    
    # === Savings badge (top right) ===
    badge_w = 32*mm
    badge_h = 12*mm
    badge_x = x + w - badge_w - 3*mm
    badge_y = y + h - title_bar_h - badge_h - 2*mm
    draw_rounded_rect(c, badge_x, badge_y, badge_w, badge_h, r=2*mm, fill_color=GREEN)
    draw_text(c, badge_x + badge_w/2, badge_y + 7.5*mm, f"ЭКОНОМИЯ", size=6, bold=True, color=WHITE, align='center')
    draw_text(c, badge_x + badge_w/2, badge_y + 2.5*mm, f"{obj['savings']}", size=10, bold=True, color=WHITE, align='center')
    
    # === Variant comparison table ===
    table_top = y + h - title_bar_h - 3*mm
    row_h = 13.5*mm
    table_x = x + 3*mm
    table_w = w - 6*mm - 35*mm  # leave space for savings badge on first row
    
    for vi, v in enumerate(obj['variants']):
        ry = table_top - (vi + 1) * row_h
        
        # Variant row background
        row_bg = v['bg']
        draw_rounded_rect(c, table_x, ry, w - 6*mm, row_h - 1*mm, r=2*mm, fill_color=row_bg)
        
        # Winner border
        if v.get('winner'):
            draw_rounded_rect(c, table_x, ry, w - 6*mm, row_h - 1*mm, r=2*mm, 
                            stroke_color=GREEN, stroke_width=2)
        
        # Variant name (left side)
        draw_text(c, table_x + 3*mm, ry + row_h - 4.5*mm, v['name'], 
                  size=9, bold=True, color=v['color'])
        
        # Details line
        draw_text(c, table_x + 3*mm, ry + row_h - 8.5*mm,
                  f"Материал: {v['material']}  |  Работа: {v['labor']}  |  Срок: {v['life']}  |  Перекрасок: {v['repaints']}",
                  size=6, color=MGRAY)
        
        # TCO (right side, large)
        tco_x = x + w - 22*mm
        draw_text(c, tco_x, ry + row_h - 8*mm, v['tco'], 
                  size=14, bold=True, color=v['color'], align='center')
        draw_text(c, tco_x, ry + row_h - 12*mm, "за 10 лет",
                  size=5.5, color=MGRAY, align='center')
    
    # === Formula explanation (bottom of card) ===
    formula_y = table_top - (3 + 0.3) * row_h
    draw_text(c, x + 5*mm, formula_y, "Формула TCO:", size=6, bold=True, color=NAVY)
    draw_text(c, x + 5*mm, formula_y - 4*mm, 
              "TCO = (Материал + Работа) × Кол-во перекрасок",
              size=6, color=DGRAY)
    draw_text(c, x + 5*mm, formula_y - 8*mm,
              f"{'TITAN' if True else ''}: 1 раз на 10 лет vs 4 раза с ЭКОНОМ",
              size=6, bold=True, color=GREEN)

# ============================================================
# DRAW: BOTTOM BAR
# ============================================================
# Navy bottom bar
c.setFillColor(NAVY)
c.rect(0, 0, W, MARGIN + 28*mm, fill=1, stroke=0)

# Orange accent
c.setFillColor(ORANGE)
c.rect(0, MARGIN + 28*mm, W, 2*mm, fill=1, stroke=0)

# Bottom text
draw_text(c, MARGIN, MARGIN + 20*mm, 
          "ПОЧЕМУ TITAN ВЫГОДНЕЕ: Красите 1 раз на 10 лет вместо 4 раз | Не нужна грунтовка | Наносится прямо на ржавчину | Гарантия 10 лет (ООО «Рутил»)",
          size=9, bold=True, color=WHITE)

draw_text(c, MARGIN, MARGIN + 13*mm,
          "Расчёт: TITAN 3в1 — 1 100₽/л, ПФ-115 ЭКОНОМ — 190₽/л, ГФ-021 — 145₽/л, ПФ-115 КМ — 280₽/л. Работа: подготовка + покраска 50₽/м². Сроки службы подтверждены испытаниями.",
          size=6.5, color=HexColor('#8899AA'))

draw_text(c, MARGIN, MARGIN + 7*mm,
          "TITAN — У КОРРОЗИИ НЕТ ШАНСОВ  |  titanpaints.ru",
          size=8, bold=True, color=ORANGE)

# QR code placeholder
qr_x = W - MARGIN - 22*mm
qr_y = MARGIN + 3*mm
draw_rounded_rect(c, qr_x, qr_y, 22*mm, 22*mm, r=2*mm, fill_color=WHITE)
draw_text(c, qr_x + 11*mm, qr_y + 12*mm, "QR", size=12, bold=True, color=NAVY, align='center')
draw_text(c, qr_x + 11*mm, qr_y + 6*mm, "Калькулятор", size=5, color=DGRAY, align='center')
draw_text(c, qr_x + 11*mm, qr_y + 2.5*mm, "TCO онлайн", size=5, color=DGRAY, align='center')

# ============================================================
# SAVE
# ============================================================
c.save()
print(f"Saved: {output}")
print(f"Size: {W/mm:.0f} × {H/mm:.0f} мм (А3 ландшафт)")
