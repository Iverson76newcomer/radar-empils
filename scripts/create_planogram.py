#!/usr/bin/env python3
"""
Generate TITAN planogram visual + DOCX with analysis.
Based on: existing Profiwood planogram + TITAN product presentation.
3 product lines: 3в1 по ржавчине, быстросохнущая матовая, молотковая.
"""

import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.font_manager as fm
fm.fontManager.addfont('/usr/share/fonts/truetype/chinese/LiberationSans-Regular.ttf')
fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, Rectangle, Circle, Polygon
import numpy as np

plt.rcParams['font.sans-serif'] = ['Liberation Sans', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

OUT_DIR = '/home/z/my-project/download/planogram_titan'
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
METAL_GRAY = '#C0C0C0'

# RAL colors for cans
RAL_COLORS = {
    'RAL 9005': '#1A1A1A',  # Черный
    'RAL 9010': '#F5F5DC',  # Белый
    'RAL 3009': '#7B3F00',  # Красно-коричневый
    'RAL 6005': '#2D5016',  # Зеленый
    'RAL 7004': '#B0B0B0',  # Серый
    'RAL 7024': '#404040',  # Графитовый серый
    'RAL 8017': '#3D2817',  # Шоколадно-коричневый
    'RAL 1021': '#E6B800',  # Желтый
    'RAL 5005': '#0047AB',  # Синий
    'RAL 1014': '#E8D4A0',  # Слоновая кость
    'RAL 9006': '#A0A0A0',  # Серебро
    'Золотистый': '#B8860B',
    'Серебристый': '#C0C0C0',
    'Медный': '#B87333',
    'Темно-зеленый': '#355E3B',
}


def draw_can(ax, x, y, width, height, color, label, ral, volume, effect='gloss'):
    """Draw a paint can."""
    # Can body
    can = FancyBboxPatch((x, y), width, height, boxstyle="round,pad=0.05",
                         facecolor=color, edgecolor=DARK, linewidth=1)
    ax.add_patch(can)
    
    # Effect indicator
    if effect == 'matte':
        # Matte - slightly different shade
        overlay = FancyBboxPatch((x, y), width, height, boxstyle="round,pad=0.05",
                                  facecolor='none', edgecolor='#666', linewidth=0.5, linestyle='--')
        ax.add_patch(overlay)
    elif effect == 'hammer':
        # Hammer pattern dots
        for _ in range(8):
            dx = x + np.random.uniform(0.2, width-0.2)
            dy = y + np.random.uniform(0.2, height-0.2)
            ax.add_patch(Circle((dx, dy), 0.06, facecolor='#FFFFFF', alpha=0.2))
    
    # Label
    ax.text(x + width/2, y + height*0.75, 'TITAN', ha='center', va='center',
            fontsize=5, color=WHITE if color != RAL_COLORS['RAL 9010'] and color != RAL_COLORS['RAL 1014'] and color != RAL_COLORS['RAL 9006'] else DARK,
            fontweight='bold')
    ax.text(x + width/2, y + height*0.55, label, ha='center', va='center',
            fontsize=3.5, color=WHITE if color != RAL_COLORS['RAL 9010'] and color != RAL_COLORS['RAL 1014'] and color != RAL_COLORS['RAL 9006'] else DARK)
    ax.text(x + width/2, y + height*0.35, ral, ha='center', va='center',
            fontsize=3, color=WHITE if color != RAL_COLORS['RAL 9010'] and color != RAL_COLORS['RAL 1014'] and color != RAL_COLORS['RAL 9006'] else DARK,
            alpha=0.8)
    ax.text(x + width/2, y + height*0.15, volume, ha='center', va='center',
            fontsize=2.5, color=WHITE if color != RAL_COLORS['RAL 9010'] and color != RAL_COLORS['RAL 1014'] and color != RAL_COLORS['RAL 9006'] else DARK,
            alpha=0.7)


def draw_planogram():
    """Draw the full TITAN planogram - shelf display."""
    fig, ax = plt.subplots(figsize=(20, 14), constrained_layout=True)
    ax.set_xlim(0, 40)
    ax.set_ylim(0, 28)
    ax.set_aspect('equal')
    ax.axis('off')
    
    # Title
    ax.text(20, 27, 'ПЛАНОГРАММА ВЫКЛАДКИ TITAN', ha='center', va='center',
            fontsize=18, fontweight='bold', color=RED)
    ax.text(20, 26, '3 линейки продукции · 3 полки · ТСО-зона · POS-материалы',
            ha='center', va='center', fontsize=10, color=GRAY, style='italic')
    
    # === STAND/STELLAzh ===
    stand_bg = FancyBboxPatch((1, 1), 38, 24, boxstyle="round,pad=0.3",
                              facecolor=LIGHT_GRAY, edgecolor=GRAY, linewidth=2)
    ax.add_patch(stand_bg)
    
    # Topper
    topper = FancyBboxPatch((2, 23), 36, 1.8, boxstyle="round,pad=0.15",
                            facecolor=RED, edgecolor=DARK_RED, linewidth=2)
    ax.add_patch(topper)
    ax.text(20, 23.9, 'TITAN — 10 лет дешевле', ha='center', va='center',
            fontsize=14, color=WHITE, fontweight='bold')
    ax.text(20, 23.2, '1 банка вместо 3 · Экономия 61%', ha='center', va='center',
            fontsize=8, color=WHITE, alpha=0.9)
    
    # === SHELF 1 (TOP): Грунт-эмаль 3в1 по ржавчине (GLOSSY) ===
    shelf1_y = 17.5
    shelf1 = Rectangle((2, shelf1_y), 36, 0.3, facecolor=SHELF_BROWN, edgecolor='#6B5635')
    ax.add_patch(shelf1)
    
    ax.text(3, 22.5, 'ПОЛКА 1: Грунт-эмаль 3в1 по ржавчине (глянцевая)',
            fontsize=9, color=RED, fontweight='bold')
    ax.text(3, 22.0, '9 цветов RAL · 10 лет гарантии · по ржавчине без грунта',
            fontsize=7, color=GRAY)
    
    # 9 colors of 3в1, 0.75L
    colors_3v1 = ['RAL 9005', 'RAL 9010', 'RAL 3009', 'RAL 6005', 'RAL 7004', 
                  'RAL 7024', 'RAL 8017', 'RAL 1021', 'RAL 5005']
    can_w = 1.2
    can_h = 3.2
    for i, ral in enumerate(colors_3v1):
        x = 4 + i * 3.7
        color = RAL_COLORS[ral]
        draw_can(ax, x, shelf1_y + 0.3, can_w, can_h, color, '3в1', ral, '0.75 л', 'gloss')
    
    # === SHELF 2 (MIDDLE): Быстросохнущая матовая ===
    shelf2_y = 11.5
    shelf2 = Rectangle((2, shelf2_y), 36, 0.3, facecolor=SHELF_BROWN, edgecolor='#6B5635')
    ax.add_patch(shelf2)
    
    ax.text(3, 16.5, 'ПОЛКА 2: Грунт-эмаль быстросохнущая (матовая)',
            fontsize=9, color=DARK, fontweight='bold')
    ax.text(3, 16.0, '7 цветов RAL · сохнет 30 мин · до -15°C',
            fontsize=7, color=GRAY)
    
    # 7 colors of быстросохнущая
    colors_fast = ['RAL 9010', 'RAL 9005', 'RAL 8017', 'RAL 7024', 'RAL 7004', 'RAL 1021', 'RAL 6005']
    for i, ral in enumerate(colors_fast):
        x = 5 + i * 4.5
        color = RAL_COLORS[ral]
        draw_can(ax, x, shelf2_y + 0.3, can_w, can_h, color, 'матовая', ral, '0.75 л', 'matte')
    
    # === SHELF 3 (BOTTOM): Молотковый эффект ===
    shelf3_y = 5.5
    shelf3 = Rectangle((2, shelf3_y), 36, 0.3, facecolor=SHELF_BROWN, edgecolor='#6B5635')
    ax.add_patch(shelf3)
    
    ax.text(3, 10.5, 'ПОЛКА 3: Эмаль с молотковым эффектом',
            fontsize=9, color='#B8860B', fontweight='bold')
    ax.text(3, 10.0, '6 цветов · декоративный эффект · скрывает дефекты · 10 лет',
            fontsize=7, color=GRAY)
    
    # 6 colors of молотковая
    colors_hammer = [('Золотистый', 'Золотистый'), ('Серебристый', 'Серебристый'), ('Коричневый', 'RAL 8017'), ('Черный', 'RAL 9005'), ('Медный', 'Медный'), ('Темно-зеленый', 'Темно-зеленый')]
    for i, (name, ral_key) in enumerate(colors_hammer):
        x = 6 + i * 5
        color = RAL_COLORS[ral_key]
        draw_can(ax, x, shelf3_y + 0.3, can_w, can_h, color, 'молот.', name, '0.8 л', 'hammer')
    
    # === TCO ZONE (bottom of stand) ===
    tco_bg = FancyBboxPatch((2, 1.5), 20, 3.5, boxstyle="round,pad=0.2",
                            facecolor=GREEN, edgecolor='none', alpha=0.1)
    ax.add_patch(tco_bg)
    ax.text(12, 4.5, 'ЗОНА ТСО', ha='center', va='center',
            fontsize=10, color=GREEN, fontweight='bold')
    ax.text(12, 3.8, 'TITAN 3в1: 1 банка = 3 продукта', ha='center', va='center',
            fontsize=8, color=DARK)
    ax.text(12, 3.2, 'Экономия 61% за 10 лет', ha='center', va='center',
            fontsize=8, color=GREEN, fontweight='bold')
    ax.text(12, 2.5, 'Выкрасы: TITAN vs ПФ-115', ha='center', va='center',
            fontsize=7, color=GRAY)
    ax.text(12, 1.9, 'QR -> калькулятор ТСО', ha='center', va='center',
            fontsize=7, color=RED)
    
    # QR code placeholder
    np.random.seed(42)
    qr_data = np.random.randint(0, 2, (7, 7))
    for i in range(7):
        for j in range(7):
            if qr_data[i, j]:
                ax.add_patch(Rectangle((23.5 + j*0.25, 2.0 + i*0.25), 0.25, 0.25,
                                       facecolor=DARK, edgecolor='none'))
    ax.text(24.5, 1.6, 'QR', ha='center', fontsize=6, color=GRAY)
    
    # Sample cards zone
    ax.text(30, 4.0, 'Выкрасы', ha='center', va='center',
            fontsize=8, color=DARK, fontweight='bold')
    ax.text(30, 3.3, 'TITAN vs ПФ-115', ha='center', va='center',
            fontsize=7, color=GRAY)
    ax.text(30, 2.5, 'до/после', ha='center', va='center',
            fontsize=7, color=GRAY)
    ax.text(30, 1.8, 'через 1 год', ha='center', va='center',
            fontsize=7, color=ORANGE)
    
    # Comparison cards
    for i, (label, color) in enumerate([('TITAN', GREEN), ('ПФ-115', RED)]):
        x = 27 + i * 2.5
        card = FancyBboxPatch((x, 2.0), 1.8, 1.5, boxstyle="round,pad=0.1",
                              facecolor=color, edgecolor='none', alpha=0.8)
        ax.add_patch(card)
        ax.text(x + 0.9, 2.75, label, ha='center', va='center',
                fontsize=6, color=WHITE, fontweight='bold')
    
    # === POS MATERIALS annotations ===
    # Topper annotation
    ax.annotate('Топпер с УТП\n«10 лет дешевле»', xy=(20, 23.5), xytext=(38, 25),
                arrowprops=dict(arrowstyle='->', color=RED, lw=1.5),
                fontsize=7, color=RED, fontweight='bold',
                bbox=dict(boxstyle='round,pad=0.3', facecolor=WHITE, edgecolor=RED))
    
    # Wobbler annotation
    ax.annotate('Воблеры\nна каждой полке', xy=(5, 21.3), xytext=(0.5, 22),
                arrowprops=dict(arrowstyle='->', color=GREEN, lw=1.5),
                fontsize=7, color=GREEN, fontweight='bold',
                bbox=dict(boxstyle='round,pad=0.3', facecolor=WHITE, edgecolor=GREEN))
    
    # TCO zone annotation
    ax.annotate('Зона ТСО:\nкалькулятор + выкрасы + QR', xy=(12, 3.5), xytext=(0.5, 2),
                arrowprops=dict(arrowstyle='->', color=GREEN, lw=1.5),
                fontsize=7, color=GREEN, fontweight='bold',
                bbox=dict(boxstyle='round,pad=0.3', facecolor=WHITE, edgecolor=GREEN))
    
    # Footer
    ax.text(20, 0.5, 'Empils · TITAN · Планограмма выкладки · 2026',
            ha='center', fontsize=8, color=GRAY, style='italic')
    
    fig.savefig(os.path.join(OUT_DIR, 'planogram_titan_full.png'), dpi=150, bbox_inches='tight')
    plt.close(fig)
    print('✅ Full planogram saved')


# Generate
print('Generating TITAN planogram...')
draw_planogram()
print(f'Saved to: {OUT_DIR}')
