#!/usr/bin/env python3
"""
Generate a simple visual instruction for TITAN 3в1 application.
5 steps with pictograms, simple language, clear icons.
Format: A4 portrait, 1 page, easy to understand.
"""

import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.font_manager as fm
fm.fontManager.addfont('/usr/share/fonts/truetype/chinese/LiberationSans-Regular.ttf')
fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, Rectangle, Circle, Polygon, FancyArrowPatch, Wedge
import numpy as np

plt.rcParams['font.sans-serif'] = ['Liberation Sans', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

OUT_DIR = '/home/z/my-project/download/titan_instruction'
os.makedirs(OUT_DIR, exist_ok=True)

RED = '#C62828'
DARK_RED = '#8E0000'
GREEN = '#16A34A'
DARK = '#1A1F36'
WHITE = '#FFFFFF'
GRAY = '#9CA3AF'
LIGHT_GRAY = '#F5F5F5'
ORANGE = '#EA580C'
BLUE = '#2563EB'
YELLOW = '#FBBF24'
BROWN = '#8B6F47'
METAL = '#C0C0C0'
RUST = '#8B4513'


def draw_pictogram(ax, cx, cy, size, step_num, draw_fn):
    """Draw a circle with step number and pictogram inside."""
    # Background circle
    circle = Circle((cx, cy), size, facecolor=RED, edgecolor=DARK_RED, linewidth=2, zorder=3)
    ax.add_patch(circle)
    
    # Step number
    ax.text(cx, cy + size * 0.7, str(step_num), ha='center', va='center',
            fontsize=16, color=WHITE, fontweight='bold', zorder=4)
    
    # Pictogram
    draw_fn(ax, cx, cy, size * 0.8)


# Pictogram drawing functions
def draw_icon_clean(ax, cx, cy, s):
    """Metal surface with brush - cleaning."""
    # Metal plate
    plate = FancyBboxPatch((cx - s*0.6, cy - s*0.4), s*1.2, s*0.5,
                           boxstyle="round,pad=0.05", facecolor=METAL, edgecolor=DARK, linewidth=1, zorder=4)
    ax.add_patch(plate)
    # Rust spots
    for _ in range(5):
        dx = cx + np.random.uniform(-s*0.4, s*0.4)
        dy = cy + np.random.uniform(-s*0.2, s*0.1)
        ax.add_patch(Circle((dx, dy), s*0.05, facecolor=RUST, edgecolor='none', zorder=5))
    # Brush
    ax.add_patch(Rectangle((cx + s*0.3, cy + s*0.1), s*0.15, s*0.3, facecolor='#8B4513', zorder=6))
    ax.add_patch(Rectangle((cx + s*0.28, cy + s*0.05), s*0.2, s*0.08, facecolor=YELLOW, zorder=6))


def draw_icon_degrease(ax, cx, cy, s):
    """Spray bottle - degreasing."""
    # Surface
    plate = FancyBboxPatch((cx - s*0.6, cy - s*0.5), s*1.2, s*0.2,
                           boxstyle="round,pad=0.03", facecolor=METAL, edgecolor=DARK, linewidth=1, zorder=4)
    ax.add_patch(plate)
    # Spray bottle
    ax.add_patch(Rectangle((cx - s*0.15, cy - s*0.1), s*0.3, s*0.5, facecolor=BLUE, edgecolor=DARK, linewidth=1, zorder=5))
    ax.add_patch(Rectangle((cx + s*0.1, cy + s*0.3), s*0.05, s*0.15, facecolor=DARK, zorder=6))
    # Spray drops
    for dy in [0.45, 0.5, 0.55]:
        ax.add_patch(Circle((cx + s*0.3, cy + s*dy), s*0.04, facecolor=BLUE, alpha=0.6, zorder=6))
    # Sparkle
    ax.text(cx - s*0.3, cy + s*0.3, '*', fontsize=14, color=YELLOW, fontweight='bold', zorder=6)


def draw_icon_stir(ax, cx, cy, s):
    """Paint can with stir stick."""
    # Can
    can = FancyBboxPatch((cx - s*0.35, cy - s*0.5), s*0.7, s*0.8,
                         boxstyle="round,pad=0.05", facecolor=RED, edgecolor=DARK_RED, linewidth=1, zorder=4)
    ax.add_patch(can)
    # TITAN text
    ax.text(cx, cy - s*0.1, 'TITAN', ha='center', va='center', fontsize=6, color=WHITE, fontweight='bold', zorder=5)
    # Stir stick
    ax.plot([cx, cx + s*0.4], [cy + s*0.1, cy + s*0.5], color='#8B4513', linewidth=3, zorder=6)
    # Swirl
    theta = np.linspace(0, 2*np.pi, 20)
    r = s * 0.15
    ax.plot(cx + r*np.cos(theta), cy + r*np.sin(theta) + s*0.2, color=WHITE, linewidth=1, zorder=6)


def draw_icon_paint(ax, cx, cy, s):
    """Brush painting metal."""
    # Metal surface
    plate = FancyBboxPatch((cx - s*0.6, cy - s*0.4), s*1.2, s*0.3,
                           boxstyle="round,pad=0.03", facecolor=GREEN, edgecolor=DARK, linewidth=1, zorder=4)
    ax.add_patch(plate)
    # Unpainted part (left)
    plate2 = FancyBboxPatch((cx - s*0.6, cy - s*0.4), s*0.5, s*0.3,
                            boxstyle="round,pad=0.03", facecolor=METAL, edgecolor=DARK, linewidth=1, zorder=5)
    ax.add_patch(plate2)
    # Brush
    ax.add_patch(Rectangle((cx - s*0.15, cy + s*0.1), s*0.1, s*0.4, facecolor='#8B4513', zorder=6))
    ax.add_patch(Rectangle((cx - s*0.18, cy - s*0.05), s*0.16, s*0.15, facecolor=DARK, zorder=6))
    # Arrow showing motion
    ax.annotate('', xy=(cx + s*0.3, cy - s*0.25), xytext=(cx - s*0.1, cy - s*0.25),
                arrowprops=dict(arrowstyle='->', color=RED, lw=2), zorder=7)


def draw_icon_done(ax, cx, cy, s):
    """Checkmark - done, 10 years."""
    # Shield
    shield = Polygon([(cx - s*0.4, cy - s*0.5), (cx + s*0.4, cy - s*0.5),
                      (cx + s*0.4, cy + s*0.2), (cx, cy + s*0.5), (cx - s*0.4, cy + s*0.2)],
                     closed=True, facecolor=GREEN, edgecolor='none', zorder=4)
    ax.add_patch(shield)
    # Checkmark
    ax.plot([cx - s*0.2, cx - s*0.05, cx + s*0.25],
            [cy, cy - s*0.15, cy + s*0.2],
            color=WHITE, linewidth=3, zorder=5)
    # 10 years text
    ax.text(cx, cy - s*0.65, '10 лет', ha='center', va='center', fontsize=8, color=GREEN, fontweight='bold', zorder=5)


def create_instruction():
    """Create the visual instruction poster."""
    fig, ax = plt.subplots(figsize=(12, 17), constrained_layout=True)
    ax.set_xlim(0, 12)
    ax.set_ylim(0, 17)
    ax.set_aspect('equal')
    ax.axis('off')

    # === HEADER ===
    header = FancyBboxPatch((0.3, 15), 11.4, 1.8, boxstyle="round,pad=0.15",
                            facecolor=RED, edgecolor='none')
    ax.add_patch(header)
    ax.text(6, 16.3, 'TITAN 3в1', ha='center', va='center',
            fontsize=22, color=WHITE, fontweight='bold')
    ax.text(6, 15.5, 'Инструкция за 5 шагов — 10 лет защиты', ha='center', va='center',
            fontsize=11, color=WHITE, alpha=0.9)

    # === INTRO ===
    ax.text(6, 14.5, 'Соблюдайте 5 простых шагов — и покрытие прослужит 10 лет без перекраски',
            ha='center', va='center', fontsize=9, color=DARK, style='italic')

    # === STEP 1: Очистить ===
    y1 = 13.0
    draw_pictogram(ax, 2, y1, 0.7, 1, draw_icon_clean)
    
    # Step text
    ax.text(3.5, y1 + 0.4, 'ОЧИСТИТЬ', ha='left', va='center',
            fontsize=13, color=RED, fontweight='bold')
    ax.text(3.5, y1 - 0.05, 'Удалите рыхлую ржавчину и старую', ha='left', va='center',
            fontsize=9, color=DARK)
    ax.text(3.5, y1 - 0.35, 'краску металлической щёткой или', ha='left', va='center',
            fontsize=9, color=DARK)
    ax.text(3.5, y1 - 0.65, 'наждачкой (Р40-Р80).', ha='left', va='center',
            fontsize=9, color=DARK)
    
    # Tip
    tip1 = FancyBboxPatch((3.5, y1 - 1.1), 7.5, 0.35, boxstyle="round,pad=0.05",
                          facecolor=YELLOW, edgecolor='none', alpha=0.3)
    ax.add_patch(tip1)
    ax.text(3.7, y1 - 0.92, '! Плотная ржавчина — можно оставить. TITAN красит по ржавчине.',
            ha='left', va='center', fontsize=7, color=DARK)

    # Arrow
    ax.annotate('', xy=(2, y1 - 1.35), xytext=(2, y1 - 1.2),
                arrowprops=dict(arrowstyle='->', color=GRAY, lw=2))

    # === STEP 2: Обезжирить ===
    y2 = 10.8
    draw_pictogram(ax, 2, y2, 0.7, 2, draw_icon_degrease)
    
    ax.text(3.5, y2 + 0.4, 'ОБЕЗЖИРИТЬ', ha='left', va='center',
            fontsize=13, color=RED, fontweight='bold')
    ax.text(3.5, y2 - 0.05, 'Протрите поверхность тряпкой с', ha='left', va='center',
            fontsize=9, color=DARK)
    ax.text(3.5, y2 - 0.35, 'растворителем (ацетон, ксилол).', ha='left', va='center',
            fontsize=9, color=DARK)
    ax.text(3.5, y2 - 0.65, 'Поверхность должна быть сухой.', ha='left', va='center',
            fontsize=9, color=DARK)
    
    tip2 = FancyBboxPatch((3.5, y2 - 1.1), 7.5, 0.35, boxstyle="round,pad=0.05",
                          facecolor=YELLOW, edgecolor='none', alpha=0.3)
    ax.add_patch(tip2)
    ax.text(3.7, y2 - 0.92, '! Не используйте уайт-спирит. После обезжиривания не трогайте руками.',
            ha='left', va='center', fontsize=7, color=DARK)

    ax.annotate('', xy=(2, y2 - 1.35), xytext=(2, y2 - 1.2),
                arrowprops=dict(arrowstyle='->', color=GRAY, lw=2))

    # === STEP 3: Перемешать ===
    y3 = 8.6
    draw_pictogram(ax, 2, y3, 0.7, 3, draw_icon_stir)
    
    ax.text(3.5, y3 + 0.4, 'ПЕРЕМЕШАТЬ', ha='left', va='center',
            fontsize=13, color=RED, fontweight='bold')
    ax.text(3.5, y3 - 0.05, 'Открыть банку и тщательно перемешать', ha='left', va='center',
            fontsize=9, color=DARK)
    ax.text(3.5, y3 - 0.35, 'до однородного цвета. Со дна вверх.', ha='left', va='center',
            fontsize=9, color=DARK)
    ax.text(3.5, y3 - 0.65, 'Готово к нанесению — грунтовать не нужно!', ha='left', va='center',
            fontsize=9, color=GREEN, fontweight='bold')
    
    tip3 = FancyBboxPatch((3.5, y3 - 1.1), 7.5, 0.35, boxstyle="round,pad=0.05",
                          facecolor=YELLOW, edgecolor='none', alpha=0.3)
    ax.add_patch(tip3)
    ax.text(3.7, y3 - 0.92, '! TITAN 3в1 = грунт + эмаль + преобразователь ржавчины в одной банке.',
            ha='left', va='center', fontsize=7, color=DARK)

    ax.annotate('', xy=(2, y3 - 1.35), xytext=(2, y3 - 1.2),
                arrowprops=dict(arrowstyle='->', color=GRAY, lw=2))

    # === STEP 4: Покрасить ===
    y4 = 6.0
    draw_pictogram(ax, 2, y4, 0.7, 4, draw_icon_paint)
    
    ax.text(3.5, y4 + 0.4, 'ПОКРАСИТЬ', ha='left', va='center',
            fontsize=13, color=RED, fontweight='bold')
    ax.text(3.5, y4 - 0.05, 'Нанесите 2-3 слоя кистью, валиком', ha='left', va='center',
            fontsize=9, color=DARK)
    ax.text(3.5, y4 - 0.35, 'или распылителем. Расход: 80-100 г/м2.', ha='left', va='center',
            fontsize=9, color=DARK)
    ax.text(3.5, y4 - 0.65, 'Между слоями: 5 часов при +20 C.', ha='left', va='center',
            fontsize=9, color=DARK)
    
    # Mini table
    table_bg = FancyBboxPatch((3.5, y4 - 1.3), 7.5, 0.55, boxstyle="round,pad=0.05",
                              facecolor=LIGHT_GRAY, edgecolor=GRAY, linewidth=0.5)
    ax.add_patch(table_bg)
    ax.text(3.8, y4 - 0.9, 'Кисть/валик: 5 ч', ha='left', va='center', fontsize=7, color=DARK)
    ax.text(6.5, y4 - 0.9, 'Распылитель: 8 ч', ha='left', va='center', fontsize=7, color=DARK)
    ax.text(9.2, y4 - 0.9, 'T: +5...+30 C', ha='left', va='center', fontsize=7, color=DARK)

    ax.annotate('', xy=(2, y4 - 1.55), xytext=(2, y4 - 1.4),
                arrowprops=dict(arrowstyle='->', color=GRAY, lw=2))

    # === STEP 5: Готово ===
    y5 = 3.5
    draw_pictogram(ax, 2, y5, 0.7, 5, draw_icon_done)
    
    ax.text(3.5, y5 + 0.4, 'ГОТОВО! 10 ЛЕТ', ha='left', va='center',
            fontsize=13, color=GREEN, fontweight='bold')
    ax.text(3.5, y5 - 0.05, 'Покрытие сохнет 5 часов. Не кантовать', ha='left', va='center',
            fontsize=9, color=DARK)
    ax.text(3.5, y5 - 0.35, 'до полного высыхания. Эксплуатация —', ha='left', va='center',
            fontsize=9, color=DARK)
    ax.text(3.5, y5 - 0.65, 'через 7 дней. Срок службы: 10 лет.', ha='left', va='center',
            fontsize=9, color=GREEN, fontweight='bold')

    # === TSO BLOCK ===
    tso_bg = FancyBboxPatch((0.5, 1.0), 11, 1.5, boxstyle="round,pad=0.15",
                            facecolor=GREEN, edgecolor='none', alpha=0.1)
    ax.add_patch(tso_bg)
    
    ax.text(6, 2.15, 'TITAN 3в1 — 10 лет дешевле', ha='center', va='center',
            fontsize=12, color=GREEN, fontweight='bold')
    ax.text(6, 1.65, '1 банка вместо 3 (грунт + эмаль + преобразователь)', ha='center', va='center',
            fontsize=9, color=DARK)
    ax.text(6, 1.25, 'Экономия 61% за 10 лет. Покрасил 1 раз — забыл на 10 лет.', ha='center', va='center',
            fontsize=9, color=GREEN, fontweight='bold')

    # === FOOTER ===
    ax.text(6, 0.5, 'Empils . TITAN . Инструкция за 5 шагов . 2026',
            ha='center', va='center', fontsize=8, color=GRAY, style='italic')
    
    # QR code placeholder
    np.random.seed(42)
    qr_data = np.random.randint(0, 2, (6, 6))
    for i in range(6):
        for j in range(6):
            if qr_data[i, j]:
                ax.add_patch(Rectangle((10.5 + j*0.15, 0.3 + i*0.15), 0.15, 0.15,
                                       facecolor=DARK, edgecolor='none'))
    ax.text(11.0, 0.15, 'Видео', ha='center', fontsize=6, color=GRAY)

    fig.savefig(os.path.join(OUT_DIR, 'instruction_5_steps.png'), dpi=150, bbox_inches='tight',
                facecolor=WHITE)
    plt.close(fig)
    print('Instruction saved!')


print('Generating visual instruction...')
create_instruction()
print(f'Saved to: {OUT_DIR}')
