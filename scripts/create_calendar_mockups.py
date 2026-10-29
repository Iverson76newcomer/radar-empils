#!/usr/bin/env python3
"""
Generate 4 calendar page mockups + 1 cover for Empils 2027.
Format: vertical wall calendar, 3 months per page, 4 quarters.
Each page: image zone (top) + calendar grid (middle) + slogan panel (bottom).
"""

import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.font_manager as fm
fm.fontManager.addfont('/usr/share/fonts/truetype/chinese/LiberationSans-Regular.ttf')
fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, Rectangle, Circle, FancyArrowPatch, Polygon
import numpy as np
import calendar

plt.rcParams['font.sans-serif'] = ['Liberation Sans', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

OUT_DIR = '/home/z/my-project/download/calendar_mockups'
os.makedirs(OUT_DIR, exist_ok=True)

RED = '#C62828'
DARK_RED = '#8E0000'
GREEN = '#16A34A'
BLUE = '#2563EB'
PURPLE = '#7C3AED'
ORANGE = '#EA580C'
DARK = '#1A1F36'
LIGHT = '#F5F5F5'
WHITE = '#FFFFFF'
GRAY = '#9CA3AF'
LIGHT_GRAY = '#E5E7EB'
WEEKEND = '#FF6B35'


def draw_calendar_page(filename, quarter, months_names, color, symbol_title, 
                       slogan, big_number, big_number_label, symbol_draw_fn=None):
    """
    Draw one calendar page.
    Page is ~30cm x 60cm (ratio 1:2), 3 months.
    """
    fig, ax = plt.subplots(figsize=(8, 16), constrained_layout=True)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 20)
    ax.set_aspect('equal')
    ax.axis('off')

    # === PAGE BACKGROUND ===
    page_bg = FancyBboxPatch((0.2, 0.2), 9.6, 19.6, boxstyle="round,pad=0.1",
                             facecolor=WHITE, edgecolor=LIGHT_GRAY, linewidth=1)
    ax.add_patch(page_bg)

    # === TOP ZONE: Symbol + Year (25%) ===
    top_bg = Rectangle((0.2, 14.5), 9.6, 5.3, facecolor=color, edgecolor='none')
    ax.add_patch(top_bg)

    # Year
    ax.text(8.5, 18.5, '2027', ha='center', va='center', fontsize=28,
            color=WHITE, fontweight='bold', alpha=0.3)

    # Symbol
    if symbol_draw_fn:
        symbol_draw_fn(ax, color)
    else:
        ax.text(5, 16.8, symbol_title, ha='center', va='center',
                fontsize=16, color=WHITE, fontweight='bold')

    # Spiral binding hint
    for x in [1.5, 3.5, 5, 6.5, 8.5]:
        circle = Circle((x, 19.5), 0.15, facecolor=WHITE, edgecolor=GRAY, linewidth=0.5)
        ax.add_patch(circle)

    # === CALENDAR ZONE (60%) ===
    cal_top = 14.3
    cal_bottom = 2.5
    month_height = (cal_top - cal_bottom) / 3

    months_data = []
    for i, mname in enumerate(months_names):
        month_num = int(months_names[i].split('|')[0])
        months_data.append((month_num, months_names[i].split('|')[1]))

    for i, (mnum, mname) in enumerate(months_data):
        y_top = cal_top - i * month_height
        y_bottom = y_top - month_height

        # Month name
        ax.text(5, y_top - 0.4, mname.upper(), ha='center', va='center',
                fontsize=14, color=color, fontweight='bold')

        # Calendar grid
        cal = calendar.monthcalendar(2027, mnum)
        days = ['ПН', 'ВТ', 'СР', 'ЧТ', 'ПТ', 'СБ', 'ВС']

        grid_left = 1.5
        grid_right = 8.5
        grid_width = grid_right - grid_left
        col_width = grid_width / 7
        row_height = (month_height - 1.0) / (len(cal) + 1)

        # Day headers
        for d, day in enumerate(days):
            x = grid_left + d * col_width + col_width / 2
            ax.text(x, y_top - 0.9, day, ha='center', va='center',
                    fontsize=7, color=GRAY)

        # Dates
        for week, week_dates in enumerate(cal):
            for d, date in enumerate(week_dates):
                if date == 0:
                    continue
                x = grid_left + d * col_width + col_width / 2
                y = y_top - 1.3 - (week + 1) * row_height
                is_weekend = d >= 5
                ax.text(x, y, str(date), ha='center', va='center',
                        fontsize=9, color=WEEKEND if is_weekend else DARK,
                        fontweight='bold' if is_weekend else 'normal')

        # Separator line
        if i < 2:
            ax.plot([1.0, 9.0], [y_bottom, y_bottom], color=LIGHT_GRAY, linewidth=0.5)

    # === BIG NUMBER ZONE (thin strip) ===
    num_y = 2.0
    num_bg = FancyBboxPatch((1.0, 1.2), 8.0, 1.1, boxstyle="round,pad=0.1",
                            facecolor=color, edgecolor='none', alpha=0.1)
    ax.add_patch(num_bg)
    ax.text(2.5, 1.75, big_number, ha='center', va='center',
            fontsize=22, color=color, fontweight='bold')
    ax.text(5.5, 1.75, big_number_label, ha='left', va='center',
            fontsize=9, color=DARK)

    # === BOTTOM PANEL: Slogan ===
    bottom_bg = Rectangle((0.2, 0.2), 9.6, 0.8, facecolor=DARK, edgecolor='none')
    ax.add_patch(bottom_bg)

    ax.text(1.5, 0.6, 'EMPILS', ha='left', va='center',
            fontsize=10, color=WHITE, fontweight='bold')
    ax.text(5, 0.6, slogan, ha='center', va='center',
            fontsize=8, color=WHITE, fontweight='bold')
    ax.text(9, 0.6, quarter, ha='right', va='center',
            fontsize=8, color=GRAY)

    fig.savefig(os.path.join(OUT_DIR, filename), dpi=150, bbox_inches='tight',
                facecolor=WHITE)
    plt.close(fig)
    print(f'  Done: {filename}')


# ===== SYMBOL DRAWING FUNCTIONS =====

def draw_shield(ax, color):
    """Draw a shield symbol for protection."""
    # Shield shape
    shield = Polygon([(4, 15.2), (6, 15.2), (6, 17.5), (5, 18.5), (4, 17.5)],
                     closed=True, facecolor=WHITE, edgecolor=WHITE, linewidth=2, alpha=0.9)
    ax.add_patch(shield)
    # Inner shield
    shield2 = Polygon([(4.3, 15.5), (5.7, 15.5), (5.7, 17.3), (5, 18.0), (4.3, 17.3)],
                      closed=True, facecolor='none', edgecolor=color, linewidth=2)
    ax.add_patch(shield2)
    # "10" inside
    ax.text(5, 16.8, '10', ha='center', va='center', fontsize=20,
            color=color, fontweight='bold')
    # "лет" below
    ax.text(5, 16.0, 'лет', ha='center', va='center', fontsize=8, color=color)


def draw_tree_rings(ax, color):
    """Draw tree rings symbol for wood protection."""
    cx, cy = 5, 16.8
    for r in [0.4, 0.7, 1.0, 1.3, 1.6]:
        circle = Circle((cx, cy), r, facecolor='none', edgecolor=WHITE, linewidth=1.5, alpha=0.8)
        ax.add_patch(circle)
    ax.text(cx, cy, '7', ha='center', va='center', fontsize=16, color=WHITE, fontweight='bold')
    ax.text(cx, cy - 2.3, 'лет защиты', ha='center', va='center', fontsize=9, color=WHITE, alpha=0.8)


def draw_one_layer(ax, color):
    """Draw single layer symbol for interior paints."""
    # Single thick line = 1 layer
    rect = FancyBboxPatch((3.5, 16.5), 3, 0.6, boxstyle="round,pad=0.05",
                          facecolor=WHITE, edgecolor=WHITE, linewidth=1)
    ax.add_patch(rect)
    ax.text(5, 16.8, '1 слой', ha='center', va='center', fontsize=12, color=color, fontweight='bold')
    # Crossed out 3 layers below
    for i in range(3):
        y = 15.8 - i * 0.25
        line = Rectangle((3.5, y), 3, 0.15, facecolor=WHITE, alpha=0.3)
        ax.add_patch(line)
    # Red X over 3 layers
    ax.plot([3.3, 6.7], [15.3, 16.0], color=WHITE, linewidth=2, alpha=0.5)
    ax.plot([3.3, 6.7], [16.0, 15.3], color=WHITE, linewidth=2, alpha=0.5)
    ax.text(5, 15.0, 'вместо 3', ha='center', va='center', fontsize=8, color=WHITE, alpha=0.7)


def draw_savings(ax, color):
    """Draw savings symbol - percentage."""
    # Big -61%
    ax.text(5, 17.0, '-61%', ha='center', va='center', fontsize=32,
            color=WHITE, fontweight='bold')
    ax.text(5, 16.0, 'экономия за 10 лет', ha='center', va='center',
            fontsize=9, color=WHITE, alpha=0.8)
    # Small bars
    bar1 = Rectangle((3.5, 15.0), 3, 0.3, facecolor=WHITE, alpha=0.3)
    ax.add_patch(bar1)
    bar2 = Rectangle((3.5, 14.7), 1.2, 0.3, facecolor=WHITE, alpha=0.7)
    ax.add_patch(bar2)


# ===== COVER PAGE =====
def draw_cover():
    fig, ax = plt.subplots(figsize=(8, 16), constrained_layout=True)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 20)
    ax.set_aspect('equal')
    ax.axis('off')

    # Background - dark
    bg = Rectangle((0, 0), 10, 20, facecolor=DARK, edgecolor='none')
    ax.add_patch(bg)

    # Red accent top
    top = Rectangle((0, 17), 10, 3, facecolor=RED, edgecolor='none')
    ax.add_patch(top)

    # Spiral
    for x in [1.5, 3.5, 5, 6.5, 8.5]:
        circle = Circle((x, 19.5), 0.15, facecolor=WHITE, edgecolor=GRAY, linewidth=0.5)
        ax.add_patch(circle)

    # EMPILS logo area
    ax.text(5, 18.5, 'EMPILS', ha='center', va='center', fontsize=24,
            color=WHITE, fontweight='bold')
    ax.text(5, 17.5, '2027', ha='center', va='center', fontsize=36,
            color=WHITE, fontweight='bold')

    # Shield symbol
    shield = Polygon([(4, 13), (6, 13), (6, 15.5), (5, 16.5), (4, 15.5)],
                     closed=True, facecolor=RED, edgecolor=WHITE, linewidth=2)
    ax.add_patch(shield)
    ax.text(5, 15.0, '10', ha='center', va='center', fontsize=28,
            color=WHITE, fontweight='bold')
    ax.text(5, 14.2, 'лет', ha='center', va='center', fontsize=10, color=WHITE)

    # Main slogan
    ax.text(5, 11.5, 'Защити один раз —', ha='center', va='center',
            fontsize=18, color=WHITE, fontweight='bold')
    ax.text(5, 10.5, 'не перекрашивай 10 лет', ha='center', va='center',
            fontsize=18, color=RED, fontweight='bold')

    # 6 brands
    brands = ['Ореол', 'TITAN', 'Profiwood', 'Newtone', 'Расцвет', 'AvantGarde']
    brand_colors = [GRAY, RED, GREEN, BLUE, ORANGE, PURPLE]
    for i, (brand, bc) in enumerate(zip(brands, brand_colors)):
        y = 8.5 - i * 0.6
        ax.text(5, y, brand, ha='center', va='center', fontsize=11,
                color=bc, fontweight='bold')

    # Bottom
    ax.text(5, 3.5, '6 брендов . 12 месяцев . 1 результат', ha='center', va='center',
            fontsize=10, color=GRAY)
    ax.text(5, 2.5, 'Защита металла . Дерева . Интерьера . Премиум', ha='center', va='center',
            fontsize=8, color=GRAY, alpha=0.7)

    # Bottom panel
    bottom = Rectangle((0, 0), 10, 1.2, facecolor=RED, edgecolor='none')
    ax.add_patch(bottom)
    ax.text(5, 0.6, 'EMPILS . 2027', ha='center', va='center',
            fontsize=10, color=WHITE, fontweight='bold')

    fig.savefig(os.path.join(OUT_DIR, '0_cover.png'), dpi=150, bbox_inches='tight',
                facecolor=DARK)
    plt.close(fig)
    print('  Done: 0_cover.png')


# ===== GENERATE ALL PAGES =====
print('Generating calendar mockups...')

# Cover
draw_cover()

# Q1 - Metal (Red)
draw_calendar_page(
    '1_q1_metal.png', 'Q1', ['1|Январь', '2|Февраль', '3|Март'],
    RED, 'Щит защиты',
    'Защити металл один раз — забудь на 10 лет',
    '1 = 3', '1 банка TITAN вместо 3 продуктов',
    symbol_draw_fn=draw_shield
)

# Q2 - Wood (Green)
draw_calendar_page(
    '2_q2_wood.png', 'Q2', ['4|Апрель', '5|Май', '6|Июнь'],
    GREEN, '7 лет защиты',
    'Защити дерево — не меняй доски 7 лет',
    '7 лет', 'Profiwood защищает дерево от гниения',
    symbol_draw_fn=draw_tree_rings
)

# Q3 - Interior (Blue)
draw_calendar_page(
    '3_q3_interior.png', 'Q3', ['7|Июль', '8|Август', '9|Сентябрь'],
    BLUE, '1 слой вместо 3',
    'Покрась стены за выходные — без запаха и разводов',
    '1 слой', 'Newtone — вместо 3 слоёв обычной краски',
    symbol_draw_fn=draw_one_layer
)

# Q4 - Premium (Purple)
draw_calendar_page(
    '4_q4_premium.png', 'Q4', ['10|Октябрь', '11|Ноябрь', '12|Декабрь'],
    PURPLE, 'Экономия 61%',
    'Заплатил больше сейчас — сэкономил на годы',
    '-61%', 'Экономия TITAN за 10 лет vs ПФ-115',
    symbol_draw_fn=draw_savings
)

print(f'\nAll mockups saved to: {OUT_DIR}')
print(f'Files: {sorted(os.listdir(OUT_DIR))}')
