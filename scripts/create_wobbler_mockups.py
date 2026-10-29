#!/usr/bin/env python3
"""
Generate wobbler mockups for TITAN 3в1 paint can.
3 designs: "1=3", "Экономия 61%", "Дороже банка — дешевле покраска"
Plus a photo-realistic mockup showing wobbler on a paint can on a shelf.
"""

import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.font_manager as fm
fm.fontManager.addfont('/usr/share/fonts/truetype/chinese/LiberationSans-Regular.ttf')
fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, Rectangle, Circle, Polygon, FancyArrowPatch
import numpy as np

plt.rcParams['font.sans-serif'] = ['Liberation Sans', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

OUT_DIR = '/home/z/my-project/download/wobbler_mockups'
os.makedirs(OUT_DIR, exist_ok=True)

RED = '#C62828'
DARK_RED = '#8E0000'
GREEN = '#16A34A'
DARK = '#1A1F36'
WHITE = '#FFFFFF'
GRAY = '#9CA3AF'
LIGHT_GRAY = '#E5E7EB'
ORANGE = '#EA580C'
YELLOW = '#FBBF24'
SHELF_BROWN = '#8B6F47'


def draw_wobbler(filename, design_num, title, main_text, sub_text, 
                 accent_color, bg_color, extra_elements=None):
    """
    Draw a wobbler design (the card part, without the stem).
    Standard wobbler size: ~10x6 cm (ratio ~1.67:1)
    """
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6), constrained_layout=True,
                                    gridspec_kw={'width_ratios': [1, 1.2]})
    
    # ===== LEFT: Wobbler design =====
    ax1.set_xlim(0, 10)
    ax1.set_ylim(0, 6)
    ax1.set_aspect('equal')
    ax1.axis('off')
    
    # Wobbler card (rounded rectangle)
    card = FancyBboxPatch((0.3, 0.3), 9.4, 5.4, boxstyle="round,pad=0.15",
                          facecolor=bg_color, edgecolor=accent_color, linewidth=2.5)
    ax1.add_patch(card)
    
    # Stem (пластиковая ножка) - visual hint
    stem = Polygon([(4.5, 0.3), (5.0, 0.3), (5.2, -0.3), (4.8, -0.3)],
                   closed=True, facecolor=GRAY, edgecolor=GRAY, alpha=0.5)
    ax1.add_patch(stem)
    ax1.text(5, -0.6, '(гибкая ножка)', ha='center', fontsize=7, color=GRAY, style='italic')
    
    # Design-specific content
    if design_num == 1:
        # "1 = 3" design
        ax1.text(2.5, 3.5, '1', ha='center', va='center', fontsize=42,
                 color=WHITE, fontweight='bold')
        ax1.text(4, 3.5, '=', ha='center', va='center', fontsize=28,
                 color=GREEN, fontweight='bold')
        ax1.text(5.5, 3.5, '3', ha='center', va='center', fontsize=42,
                 color=WHITE, fontweight='bold')
        
        ax1.text(5, 1.8, '1 банка TITAN 3в1', ha='center', fontsize=9,
                 color=WHITE, fontweight='bold')
        ax1.text(5, 1.1, 'вместо 3 продуктов', ha='center', fontsize=8,
                 color=WHITE, alpha=0.8)
        
        # TITAN logo
        ax1.text(8.5, 5.0, 'TITAN', ha='center', va='center', fontsize=10,
                 color=WHITE, fontweight='bold')
        ax1.text(8.5, 4.5, '3в1', ha='center', va='center', fontsize=8,
                 color=WHITE)
    
    elif design_num == 2:
        # "Экономия 61%" design
        ax1.text(5, 3.8, '61%', ha='center', va='center', fontsize=44,
                 color=WHITE, fontweight='bold')
        ax1.text(5, 2.5, 'экономия за 10 лет', ha='center', fontsize=10,
                 color=WHITE, fontweight='bold')
        
        # Mini comparison bars
        bar1 = FancyBboxPatch((1.5, 1.0), 4.5, 0.5, boxstyle="round,pad=0.05",
                              facecolor='#EF4444', edgecolor='none', alpha=0.7)
        ax1.add_patch(bar1)
        ax1.text(1.7, 1.25, 'ПФ-115 + грунт', fontsize=6, color=WHITE, fontweight='bold')
        
        bar2 = FancyBboxPatch((1.5, 0.5), 1.8, 0.5, boxstyle="round,pad=0.05",
                              facecolor=GREEN, edgecolor='none', alpha=0.8)
        ax1.add_patch(bar2)
        ax1.text(1.7, 0.75, 'TITAN 3в1', fontsize=6, color=WHITE, fontweight='bold')
        
        ax1.text(7.5, 0.75, '5 320 руб.', fontsize=7, color=WHITE, fontweight='bold')
        
        # TITAN logo
        ax1.text(8.5, 5.0, 'TITAN', ha='center', va='center', fontsize=10,
                 color=WHITE, fontweight='bold')
    
    elif design_num == 3:
        # Slogan design
        ax1.text(5, 4.0, 'Дороже банка —', ha='center', va='center', fontsize=13,
                 color=WHITE, fontweight='bold')
        ax1.text(5, 3.0, 'дешевле покраска', ha='center', va='center', fontsize=13,
                 color=WHITE, fontweight='bold')
        
        # Separator
        ax1.plot([2, 8], [2.2, 2.2], color=WHITE, linewidth=1, alpha=0.3)
        
        ax1.text(5, 1.5, 'TITAN 3в1: 1 банка вместо 3', ha='center', fontsize=8,
                 color=WHITE, alpha=0.9)
        ax1.text(5, 0.9, '10 лет службы · по ржавчине', ha='center', fontsize=7,
                 color=WHITE, alpha=0.7)
        
        # TITAN logo
        ax1.text(8.5, 5.0, 'TITAN', ha='center', va='center', fontsize=10,
                 color=WHITE, fontweight='bold')
    
    # Title
    ax1.text(5, 5.7, title, ha='center', va='center', fontsize=10,
             color=accent_color, fontweight='bold', style='italic')
    
    # ===== RIGHT: Wobbler on a shelf with paint can =====
    ax2.set_xlim(0, 12)
    ax2.set_ylim(0, 8)
    ax2.set_aspect('equal')
    ax2.axis('off')
    
    ax2.text(6, 7.5, 'Как выглядит на полке', ha='center', fontsize=11,
             color=DARK, fontweight='bold', style='italic')
    
    # Shelf
    shelf = Rectangle((0.5, 0.5), 11, 0.4, facecolor=SHELF_BROWN, edgecolor='#6B5635', linewidth=1)
    ax2.add_patch(shelf)
    
    # Paint cans (3 cans)
    can_colors = [RED, DARK, '#2E5C8A']
    can_labels = ['TITAN\n3в1', 'TITAN\n3в1', 'TITAN\n3в1']
    can_ral = ['RAL 9005', 'RAL 9010', 'RAL 6005']
    
    for i, (cc, cl, ral) in enumerate(zip(can_colors, can_labels, can_ral)):
        cx = 2 + i * 3.5
        
        # Can body
        can = FancyBboxPatch((cx, 1.0), 2.5, 5.0, boxstyle="round,pad=0.1",
                             facecolor=cc, edgecolor=DARK, linewidth=1)
        ax2.add_patch(can)
        
        # Can lid
        lid = Rectangle((cx + 0.1, 5.8), 2.3, 0.3, facecolor='#444', edgecolor=DARK, linewidth=0.5)
        ax2.add_patch(lid)
        
        # Label text
        ax2.text(cx + 1.25, 4.8, cl, ha='center', va='center', fontsize=8,
                 color=WHITE, fontweight='bold')
        ax2.text(cx + 1.25, 3.8, ral, ha='center', va='center', fontsize=6,
                 color=WHITE, alpha=0.8)
        ax2.text(cx + 1.25, 2.5, '0.9 кг', ha='center', va='center', fontsize=6,
                 color=WHITE, alpha=0.7)
        
        # Wobbler on first can only
        if i == 0:
            # Stem
            stem = Polygon([(cx + 1.0, 5.0), (cx + 1.5, 5.0), (cx + 2.0, 6.5), (cx + 1.5, 6.5)],
                          closed=True, facecolor='none', edgecolor=GRAY, linewidth=1, linestyle='--')
            ax2.add_patch(stem)
            
            # Wobbler card (angled)
            wobbler = FancyBboxPatch((cx + 0.3, 6.2), 3.5, 1.3, boxstyle="round,pad=0.1",
                                     facecolor=bg_color, edgecolor=accent_color, linewidth=1.5,
                                     alpha=0.95)
            ax2.add_patch(wobbler)
            
            # Wobbler text (scaled down)
            if design_num == 1:
                ax2.text(cx + 1.2, 7.1, '1', ha='center', va='center', fontsize=14,
                         color=WHITE, fontweight='bold')
                ax2.text(cx + 1.6, 7.1, '=', ha='center', va='center', fontsize=10,
                         color=GREEN, fontweight='bold')
                ax2.text(cx + 2.0, 7.1, '3', ha='center', va='center', fontsize=14,
                         color=WHITE, fontweight='bold')
                ax2.text(cx + 2.5, 6.6, 'банка = продуктов', fontsize=6, color=WHITE)
            elif design_num == 2:
                ax2.text(cx + 1.5, 7.1, '61%', ha='center', va='center', fontsize=14,
                         color=WHITE, fontweight='bold')
                ax2.text(cx + 2.5, 6.6, 'экономия', fontsize=6, color=WHITE)
            elif design_num == 3:
                ax2.text(cx + 2.0, 7.1, 'Дороже банка —', ha='center', va='center', fontsize=7,
                         color=WHITE, fontweight='bold')
                ax2.text(cx + 2.0, 6.6, 'дешевле покраска', ha='center', va='center', fontsize=7,
                         color=WHITE, fontweight='bold')
    
    # Price tags
    for i in range(3):
        cx = 2 + i * 3.5
        price_bg = FancyBboxPatch((cx + 0.3, 0.0), 1.8, 0.4, boxstyle="round,pad=0.05",
                                  facecolor=YELLOW, edgecolor='none')
        ax2.add_patch(price_bg)
        ax2.text(cx + 1.2, 0.2, '5 720 руб.', ha='center', va='center', fontsize=6,
                 color=DARK, fontweight='bold')
    
    # Arrow pointing to wobbler
    ax2.annotate('Воблер\nна ножке', xy=(3.5, 6.5), xytext=(5.5, 7.5),
                 arrowprops=dict(arrowstyle='->', color=RED, lw=2),
                 fontsize=8, color=RED, fontweight='bold')
    
    # Size note
    ax2.text(6, -0.3, 'Размер воблера: 10×6 см · Материал: картон 300 г/м² + гибкая ножка',
             ha='center', fontsize=7, color=GRAY, style='italic')
    
    fig.savefig(os.path.join(OUT_DIR, filename), dpi=150, bbox_inches='tight',
                facecolor=WHITE)
    plt.close(fig)
    print(f'  Done: {filename}')


# ===== Generate all 3 wobbler designs =====
print('Generating wobbler mockups...')

draw_wobbler(
    'wobbler_1_one_equals_three.png', 1,
    'Вариант 1: «1 = 3»',
    '1 = 3', '1 банка вместо 3 продуктов',
    RED, RED
)

draw_wobbler(
    'wobbler_2_savings.png', 2,
    'Вариант 2: «61% экономии»',
    '61%', 'экономия за 10 лет',
    GREEN, RED
)

draw_wobbler(
    'wobbler_3_slogan.png', 3,
    'Вариант 3: «Дороже банка — дешевле покраска»',
    'Дороже банка — дешевле покраска', 'Слоган ТСО',
    RED, DARK_RED
)

print(f'\nAll wobbler mockups saved to: {OUT_DIR}')
print(f'Files: {sorted(os.listdir(OUT_DIR))}')
