#!/usr/bin/env python3
"""
Generate visual mockups for TCO elements on paint can labels.
7 mockups as PNG images showing how TCO info can be placed on the can.
"""

import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.font_manager as fm
fm.fontManager.addfont('/usr/share/fonts/truetype/chinese/LiberationSans-Regular.ttf')
fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, Rectangle, Circle, FancyArrowPatch
import numpy as np

plt.rcParams['font.sans-serif'] = ['Liberation Sans', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

OUT_DIR = '/home/z/my-project/download/tco_mockups'
os.makedirs(OUT_DIR, exist_ok=True)

# Brand colors
RED = '#C62828'
DARK_RED = '#8E0000'
GREEN = '#16A34A'
BLUE = '#2563EB'
ORANGE = '#EA580C'
DARK = '#1A1F36'
LIGHT = '#F5F5F5'
WHITE = '#FFFFFF'
GRAY = '#9CA3AF'

def draw_can_outline(ax, title="ЭТИКЕТКА БАНКИ"):
    """Draw a simple can outline for context."""
    ax.set_xlim(0, 20)
    ax.set_ylim(0, 12)
    ax.set_aspect('equal')
    ax.axis('off')
    ax.text(10, 11.5, title, ha='center', va='center', fontsize=10, color=GRAY, style='italic')


# ===== 1. QR-код блок =====
def mockup_qr():
    fig, ax = plt.subplots(figsize=(8, 6), constrained_layout=True)
    ax.set_xlim(0, 20)
    ax.set_ylim(0, 12)
    ax.set_aspect('equal')
    ax.axis('off')
    
    # Title
    ax.text(10, 11.3, 'Макет 1: QR-код → калькулятор ТСО', 
            ha='center', fontsize=14, fontweight='bold', color=DARK)
    ax.text(10, 10.5, 'Размещение: обратная сторона этикетки, внизу', 
            ha='center', fontsize=10, color=GRAY, style='italic')
    
    # Label area
    label = FancyBboxPatch((1, 1), 18, 8.5, boxstyle="round,pad=0.2",
                           facecolor=LIGHT, edgecolor=GRAY, linewidth=1)
    ax.add_patch(label)
    ax.text(10, 8.8, 'ОБОРОТ ЭТИКЕТКИ', ha='center', fontsize=9, color=GRAY)
    
    # Fake instruction text lines
    for i, y in enumerate([8.2, 7.8, 7.4, 7.0, 6.6, 6.2, 5.8, 5.4]):
        ax.plot([2, 18], [y, y], color='#E0E0E0', linewidth=0.5)
    
    ax.text(2, 5.0, 'ИНСТРУКЦИЯ ПО НАНЕСЕНИЮ', fontsize=8, color=GRAY, fontweight='bold')
    ax.text(2, 4.6, 'Подготовка поверхности: очистить от рыхлой ржавчины...', fontsize=7, color=GRAY)
    ax.text(2, 4.2, 'Нанесение: 2-3 слоя кистью, валиком или краскопультом...', fontsize=7, color=GRAY)
    ax.text(2, 3.8, 'Время высыхания: 5 часов при +20°C...', fontsize=7, color=GRAY)
    
    # QR block
    qr_bg = FancyBboxPatch((13, 1.5), 5, 4, boxstyle="round,pad=0.15",
                           facecolor=WHITE, edgecolor=RED, linewidth=2)
    ax.add_patch(qr_bg)
    
    # Fake QR pattern
    np.random.seed(42)
    qr_data = np.random.randint(0, 2, (8, 8))
    for i in range(8):
        for j in range(8):
            if qr_data[i, j]:
                ax.add_patch(Rectangle((13.5 + j*0.4, 3.5 + i*0.4), 0.4, 0.4, 
                                       facecolor=DARK, edgecolor='none'))
    # QR corners
    for cx, cy in [(13.5, 4.7), (16.3, 4.7), (13.5, 3.5)]:
        ax.add_patch(Rectangle((cx, cy), 0.8, 0.8, facecolor=DARK, edgecolor='none'))
        ax.add_patch(Rectangle((cx+0.15, cy+0.15), 0.5, 0.5, facecolor=WHITE, edgecolor='none'))
        ax.add_patch(Rectangle((cx+0.3, cy+0.3), 0.2, 0.2, facecolor=DARK, edgecolor='none'))
    
    # Call to action
    ax.text(15.5, 3.0, 'Сканируй и', ha='center', fontsize=8, color=DARK)
    ax.text(15.5, 2.6, 'посчитай экономию', ha='center', fontsize=8, color=RED, fontweight='bold')
    
    # Arrow pointing to QR
    ax.annotate('', xy=(13, 3.5), xytext=(10, 4),
                arrowprops=dict(arrowstyle='->', color=RED, lw=2))
    ax.text(10, 4.3, '2×2 см', fontsize=8, color=RED, fontweight='bold')
    
    fig.savefig(os.path.join(OUT_DIR, '1_qr_block.png'), dpi=150, bbox_inches='tight')
    plt.close(fig)
    print('✅ Mockup 1: QR block')


# ===== 2. "1=3" визуальная формула =====
def mockup_one_equals_three():
    fig, ax = plt.subplots(figsize=(10, 5), constrained_layout=True)
    ax.set_xlim(0, 20)
    ax.set_ylim(0, 10)
    ax.set_aspect('equal')
    ax.axis('off')
    
    ax.text(10, 9.3, 'Макет 2: «1 банка = 3 продукта»', 
            ha='center', fontsize=14, fontweight='bold', color=DARK)
    ax.text(10, 8.5, 'Размещение: лицевая сторона, под названием продукта', 
            ha='center', fontsize=10, color=GRAY, style='italic')
    
    # Left: 1 can
    can_bg = FancyBboxPatch((1, 2), 4, 5, boxstyle="round,pad=0.2",
                            facecolor=RED, edgecolor=DARK_RED, linewidth=2)
    ax.add_patch(can_bg)
    ax.text(3, 5.5, 'TITAN', ha='center', fontsize=11, color=WHITE, fontweight='bold')
    ax.text(3, 4.8, '3в1', ha='center', fontsize=14, color=WHITE, fontweight='bold')
    ax.text(3, 3.0, '1 банка', ha='center', fontsize=10, color=WHITE)
    ax.text(3, 1.2, '1 продукт', ha='center', fontsize=9, color=DARK, fontweight='bold')
    
    # Equals sign
    ax.text(6.5, 4.5, '=', ha='center', va='center', fontsize=30, color=GREEN, fontweight='bold')
    ax.text(6.5, 2.5, 'ЗАМЕНЯЕТ', ha='center', fontsize=8, color=GREEN, fontweight='bold')
    
    # Right: 3 cans
    for i, (label, color) in enumerate([
        ('ГФ-021\nгрунт', '#6B7280'),
        ('ПФ-115\nэмаль', '#9CA3AF'),
        ('Преобр.\nржавч.', '#D1D5DB')
    ]):
        x = 8.5 + i * 3.5
        can = FancyBboxPatch((x, 2.5), 2.5, 4, boxstyle="round,pad=0.15",
                             facecolor=color, edgecolor=GRAY, linewidth=1.5)
        ax.add_patch(can)
        ax.text(x + 1.25, 4.5, label, ha='center', va='center', fontsize=8, 
                color=DARK, fontweight='bold')
    
    ax.text(13.75, 1.2, '3 продукта', ha='center', fontsize=9, color=GRAY, fontweight='bold')
    
    # Savings badge
    badge = Circle((17.5, 5.5), 1.3, facecolor=GREEN, edgecolor='none')
    ax.add_patch(badge)
    ax.text(17.5, 5.8, '61%', ha='center', va='center', fontsize=12, color=WHITE, fontweight='bold')
    ax.text(17.5, 4.8, 'экономия', ha='center', va='center', fontsize=8, color=WHITE)
    
    # Bottom message
    msg_bg = FancyBboxPatch((1, 0.2), 18, 0.8, boxstyle="round,pad=0.1",
                            facecolor=GREEN, edgecolor='none', alpha=0.15)
    ax.add_patch(msg_bg)
    ax.text(10, 0.6, '1 банка TITAN 3в1 = грунт + эмаль + преобразователь ржавчины', 
            ha='center', fontsize=9, color=GREEN, fontweight='bold')
    
    fig.savefig(os.path.join(OUT_DIR, '2_one_equals_three.png'), dpi=150, bbox_inches='tight')
    plt.close(fig)
    print('✅ Mockup 2: 1=3')


# ===== 3. Цепочка нанесения =====
def mockup_chain():
    fig, ax = plt.subplots(figsize=(12, 6), constrained_layout=True)
    ax.set_xlim(0, 24)
    ax.set_ylim(0, 12)
    ax.set_aspect('equal')
    ax.axis('off')
    
    ax.text(12, 11.3, 'Макет 3: Цепочка нанесения (сравнение)', 
            ha='center', fontsize=14, fontweight='bold', color=DARK)
    ax.text(12, 10.5, 'Размещение: оборот этикетки', 
            ha='center', fontsize=10, color=GRAY, style='italic')
    
    # Variant 1: Classic
    ax.text(6, 9.2, 'Классическая схема', ha='center', fontsize=10, color=GRAY, fontweight='bold')
    
    steps1 = [
        ('Ржавчина', '#EF4444'),
        ('ГФ-021\nгрунт', '#6B7280'),
        ('ПФ-115\nэмаль', '#9CA3AF'),
    ]
    for i, (label, color) in enumerate(steps1):
        x = 1.5 + i * 3.5
        box = FancyBboxPatch((x, 6.5), 2.5, 2, boxstyle="round,pad=0.15",
                             facecolor=color, edgecolor='none', alpha=0.8)
        ax.add_patch(box)
        ax.text(x + 1.25, 7.5, label, ha='center', va='center', fontsize=8, 
                color=WHITE, fontweight='bold')
        if i < len(steps1) - 1:
            ax.annotate('', xy=(x + 3.3, 7.5), xytext=(x + 2.7, 7.5),
                        arrowprops=dict(arrowstyle='->', color=GRAY, lw=2))
    
    ax.text(6, 5.8, '3 продукта • 3 слоя • 2 дня', ha='center', fontsize=8, color=GRAY)
    ax.text(6, 5.2, 'Стоимость: 13 860 руб. за 10 лет', ha='center', fontsize=9, color='#EF4444', fontweight='bold')
    
    # Divider
    ax.plot([12, 12], [5, 9.5], color=GRAY, linewidth=1, linestyle='--')
    
    # Variant 2: TITAN 3в1
    ax.text(18, 9.2, 'TITAN 3в1', ha='center', fontsize=10, color=RED, fontweight='bold')
    
    steps2 = [
        ('Ржавчина', '#EF4444'),
        ('TITAN\n3в1', RED),
    ]
    for i, (label, color) in enumerate(steps2):
        x = 14 + i * 4
        box = FancyBboxPatch((x, 6.5), 3, 2, boxstyle="round,pad=0.15",
                             facecolor=color, edgecolor='none', alpha=0.85)
        ax.add_patch(box)
        ax.text(x + 1.5, 7.5, label, ha='center', va='center', fontsize=9, 
                color=WHITE, fontweight='bold')
        if i < len(steps2) - 1:
            ax.annotate('', xy=(x + 3.8, 7.5), xytext=(x + 3.2, 7.5),
                        arrowprops=dict(arrowstyle='->', color=GREEN, lw=2.5))
    
    ax.text(18, 5.8, '1 продукт • 2 слоя • 1 день', ha='center', fontsize=8, color=GREEN, fontweight='bold')
    ax.text(18, 5.2, 'Стоимость: 5 320 руб. за 10 лет', ha='center', fontsize=9, color=GREEN, fontweight='bold')
    
    # Savings bar
    savings_bg = FancyBboxPatch((1, 2.5), 22, 1.8, boxstyle="round,pad=0.2",
                                facecolor=GREEN, edgecolor='none', alpha=0.1)
    ax.add_patch(savings_bg)
    
    ax.text(12, 3.8, 'ЭКОНОМИЯ: 8 540 руб. (61%) + 1 целый день времени', 
            ha='center', fontsize=12, color=GREEN, fontweight='bold')
    ax.text(12, 3.0, 'TITAN 3в1 красит прямо по ржавчине — не нужно зачищать до металла', 
            ha='center', fontsize=9, color=DARK)
    
    # Size note
    ax.text(12, 1.2, 'Размер на этикетке: ~5×4 см', ha='center', fontsize=8, color=GRAY, style='italic')
    
    fig.savefig(os.path.join(OUT_DIR, '3_chain.png'), dpi=150, bbox_inches='tight')
    plt.close(fig)
    print('✅ Mockup 3: Chain')


# ===== 4. Цена за год службы =====
def mockup_price_per_year():
    fig, ax = plt.subplots(figsize=(8, 6), constrained_layout=True)
    ax.set_xlim(0, 16)
    ax.set_ylim(0, 10)
    ax.set_aspect('equal')
    ax.axis('off')
    
    ax.text(8, 9.3, 'Макет 4: «Цена за год службы»', 
            ha='center', fontsize=14, fontweight='bold', color=DARK)
    ax.text(8, 8.5, 'Размещение: лицевая сторона, рядом с ценником', 
            ha='center', fontsize=10, color=GRAY, style='italic')
    
    # Price block (standard)
    price_bg = FancyBboxPatch((2, 4.5), 5, 2.5, boxstyle="round,pad=0.15",
                              facecolor=LIGHT, edgecolor=GRAY, linewidth=1)
    ax.add_patch(price_bg)
    ax.text(4.5, 6.2, 'Цена', ha='center', fontsize=9, color=GRAY)
    ax.text(4.5, 5.2, '5 720 руб.', ha='center', fontsize=16, color=DARK, fontweight='bold')
    
    # Arrow
    ax.annotate('', xy=(8.5, 5.5), xytext=(7.2, 5.5),
                arrowprops=dict(arrowstyle='->', color=GREEN, lw=2.5))
    
    # Price per year block (highlighted)
    ppy_bg = FancyBboxPatch((8.5, 3.5), 5.5, 3.5, boxstyle="round,pad=0.2",
                            facecolor=GREEN, edgecolor='none', alpha=0.9)
    ax.add_patch(ppy_bg)
    ax.text(11.25, 6.0, '532 руб.', ha='center', fontsize=22, color=WHITE, fontweight='bold')
    ax.text(11.25, 4.8, 'за год службы', ha='center', fontsize=10, color=WHITE)
    ax.text(11.25, 4.0, '(10 лет гарантии)', ha='center', fontsize=8, color=WHITE, alpha=0.8)
    
    # Comparison with ПФ-115
    comp_bg = FancyBboxPatch((2, 1), 12, 1.8, boxstyle="round,pad=0.15",
                             facecolor=LIGHT, edgecolor=GRAY, linewidth=0.5)
    ax.add_patch(comp_bg)
    ax.text(8, 2.4, 'Для сравнения: ПФ-115 + грунт = 1 386 руб./год (служит 2-3 года)', 
            ha='center', fontsize=8, color=GRAY)
    ax.text(8, 1.5, 'TITAN 3в1 дешевле на 62% в пересчёте на год службы', 
            ha='center', fontsize=9, color=GREEN, fontweight='bold')
    
    fig.savefig(os.path.join(OUT_DIR, '4_price_per_year.png'), dpi=150, bbox_inches='tight')
    plt.close(fig)
    print('✅ Mockup 4: Price per year')


# ===== 5. Иконки ТСО =====
def mockup_icons():
    fig, ax = plt.subplots(figsize=(10, 4), constrained_layout=True)
    ax.set_xlim(0, 20)
    ax.set_ylim(0, 8)
    ax.set_aspect('equal')
    ax.axis('off')
    
    ax.text(10, 7.3, 'Макет 5: Иконки ТСО (компактный блок)', 
            ha='center', fontsize=14, fontweight='bold', color=DARK)
    ax.text(10, 6.5, 'Размещение: лицевая сторона, сбоку. Размер: 4×1.5 см', 
            ha='center', fontsize=10, color=GRAY, style='italic')
    
    # Three icon blocks
    icons = [
        ('1', 'банка', RED, 'Вместо 3-х'),
        ('2', 'слоя', BLUE, 'Вместо 3-х'),
        ('10', 'лет', GREEN, 'Вместо 2-3'),
    ]
    
    for i, (num, label, color, sub) in enumerate(icons):
        x = 3 + i * 5.5
        
        # Circle with number
        circle = Circle((x, 3.5), 1.2, facecolor=color, edgecolor='none')
        ax.add_patch(circle)
        ax.text(x, 3.5, num, ha='center', va='center', fontsize=20, color=WHITE, fontweight='bold')
        
        # Label below
        ax.text(x, 1.8, label, ha='center', fontsize=10, color=DARK, fontweight='bold')
        ax.text(x, 1.0, sub, ha='center', fontsize=8, color=GRAY)
    
    # Bottom bar
    bar_bg = FancyBboxPatch((1, 0.1), 18, 0.5, boxstyle="round,pad=0.1",
                            facecolor=GREEN, edgecolor='none', alpha=0.1)
    ax.add_patch(bar_bg)
    ax.text(10, 0.35, '1 банка TITAN 3в1 = 2 слоя = 10 лет службы', 
            ha='center', fontsize=9, color=GREEN, fontweight='bold')
    
    fig.savefig(os.path.join(OUT_DIR, '5_icons.png'), dpi=150, bbox_inches='tight')
    plt.close(fig)
    print('✅ Mockup 5: Icons')


# ===== 6. Шкала ТСО =====
def mockup_scale():
    fig, ax = plt.subplots(figsize=(10, 5), constrained_layout=True)
    ax.set_xlim(0, 20)
    ax.set_ylim(0, 10)
    ax.set_aspect('equal')
    ax.axis('off')
    
    ax.text(10, 9.3, 'Макет 6: Шкала ТСО (сравнение за 10 лет)', 
            ha='center', fontsize=14, fontweight='bold', color=DARK)
    ax.text(10, 8.5, 'Размещение: оборот этикетки. Размер: 6×3 см', 
            ha='center', fontsize=10, color=GRAY, style='italic')
    
    # Title
    ax.text(2, 7.2, 'Стоимость владения за 10 лет:', fontsize=10, color=DARK, fontweight='bold')
    
    # Bar 1: ПФ-115
    bar1 = FancyBboxPatch((2, 4.5), 14, 1.5, boxstyle="round,pad=0.1",
                          facecolor='#EF4444', edgecolor='none', alpha=0.8)
    ax.add_patch(bar1)
    ax.text(2.5, 5.25, 'ПФ-115 + грунт', fontsize=9, color=WHITE, fontweight='bold')
    ax.text(15.5, 5.25, '13 860 руб.', ha='right', fontsize=11, color=WHITE, fontweight='bold')
    
    # Bar 2: TITAN
    bar2 = FancyBboxPatch((2, 2.0), 5.4, 1.5, boxstyle="round,pad=0.1",
                          facecolor=GREEN, edgecolor='none', alpha=0.85)
    ax.add_patch(bar2)
    ax.text(2.5, 2.75, 'TITAN 3в1', fontsize=9, color=WHITE, fontweight='bold')
    ax.text(7.1, 2.75, '5 320 руб.', ha='right', fontsize=11, color=WHITE, fontweight='bold')
    
    # Savings label
    savings_bg = FancyBboxPatch((8, 2.0), 4, 1.5, boxstyle="round,pad=0.1",
                                facecolor=GREEN, edgecolor='none', alpha=0.15)
    ax.add_patch(savings_bg)
    ax.text(10, 2.75, 'Экономия: 8 540 руб.', ha='center', fontsize=9, color=GREEN, fontweight='bold')
    
    # Percentage
    ax.text(16, 2.75, '−61%', ha='center', fontsize=14, color=GREEN, fontweight='bold')
    
    # QR note
    ax.text(10, 0.8, '> Сканируй QR для расчёта своего объекта', 
            ha='center', fontsize=9, color=RED)
    
    fig.savefig(os.path.join(OUT_DIR, '6_scale.png'), dpi=150, bbox_inches='tight')
    plt.close(fig)
    print('✅ Mockup 6: Scale')


# ===== 7. Общий макет банки (все элементы вместе) =====
def mockup_full_can():
    fig, axes = plt.subplots(1, 2, figsize=(16, 8), constrained_layout=True)
    
    # ===== FRONT =====
    ax = axes[0]
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 16)
    ax.set_aspect('equal')
    ax.axis('off')
    ax.set_title('ЛИЦЕВАЯ СТОРОНА', fontsize=11, color=GRAY)
    
    # Can body
    can = FancyBboxPatch((1, 1), 8, 14, boxstyle="round,pad=0.1",
                         facecolor=RED, edgecolor=DARK_RED, linewidth=2)
    ax.add_patch(can)
    
    # Brand
    ax.text(5, 13.5, 'TITAN', ha='center', fontsize=20, color=WHITE, fontweight='bold')
    ax.text(5, 12.5, 'Грунт-эмаль 3в1', ha='center', fontsize=11, color=WHITE)
    
    # 1=3 formula (element 2)
    ax.text(2, 10.5, '1', fontsize=16, color=WHITE, fontweight='bold')
    ax.text(2.5, 10.5, '=', fontsize=14, color=GREEN, fontweight='bold')
    ax.text(3.2, 10.5, '3', fontsize=16, color=WHITE, fontweight='bold')
    ax.text(4, 10.5, 'банка = продукта', fontsize=8, color=WHITE)
    
    # Icons (element 5)
    for i, (num, label, color) in enumerate([
        ('1', 'банка', WHITE),
        ('2', 'слоя', WHITE),
        ('10', 'лет', WHITE)
    ]):
        x = 1.5 + i * 2.5
        circle = Circle((x, 8.5), 0.7, facecolor='none', edgecolor=WHITE, linewidth=2)
        ax.add_patch(circle)
        ax.text(x, 8.5, num, ha='center', va='center', fontsize=12, color=WHITE, fontweight='bold')
        ax.text(x, 7.3, label, ha='center', fontsize=7, color=WHITE)
    
    # Price per year (element 4)
    ppy_bg = FancyBboxPatch((1.5, 4.5), 4, 1.5, boxstyle="round,pad=0.1",
                            facecolor=GREEN, edgecolor='none', alpha=0.9)
    ax.add_patch(ppy_bg)
    ax.text(3.5, 5.5, '532 руб.', ha='center', fontsize=12, color=WHITE, fontweight='bold')
    ax.text(3.5, 4.8, 'за год службы', ha='center', fontsize=7, color=WHITE)
    
    # Color RAL
    ax.text(7, 5.0, 'RAL 9005', ha='center', fontsize=9, color=WHITE, fontweight='bold')
    ax.text(7, 4.5, 'Чёрный', ha='center', fontsize=8, color=WHITE)
    color_swatch = Rectangle((6.2, 2.5), 1.6, 1.5, facecolor='#1A1A1A', edgecolor=WHITE, linewidth=1)
    ax.add_patch(color_swatch)
    
    # Volume
    ax.text(5, 1.8, '0.9 кг', ha='center', fontsize=9, color=WHITE)
    
    # ===== BACK =====
    ax = axes[1]
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 16)
    ax.set_aspect('equal')
    ax.axis('off')
    ax.set_title('ОБОРОТ', fontsize=11, color=GRAY)
    
    # Label area
    label = FancyBboxPatch((1, 1), 8, 14, boxstyle="round,pad=0.1",
                           facecolor=LIGHT, edgecolor=GRAY, linewidth=1)
    ax.add_patch(label)
    
    # TCO Scale (element 6)
    ax.text(5, 13.5, 'Стоимость за 10 лет:', ha='center', fontsize=8, color=DARK, fontweight='bold')
    
    bar1 = FancyBboxPatch((1.5, 11.5), 7, 1, boxstyle="round,pad=0.05",
                          facecolor='#EF4444', edgecolor='none', alpha=0.8)
    ax.add_patch(bar1)
    ax.text(1.8, 12.0, 'ПФ-115+грунт', fontsize=7, color=WHITE, fontweight='bold')
    ax.text(8.2, 12.0, '13 860 руб.', ha='right', fontsize=8, color=WHITE, fontweight='bold')
    
    bar2 = FancyBboxPatch((1.5, 9.8), 2.7, 1, boxstyle="round,pad=0.05",
                          facecolor=GREEN, edgecolor='none', alpha=0.85)
    ax.add_patch(bar2)
    ax.text(1.8, 10.3, 'TITAN 3в1', fontsize=7, color=WHITE, fontweight='bold')
    ax.text(4.0, 10.3, '5 320 руб.', ha='right', fontsize=8, color=WHITE, fontweight='bold')
    
    ax.text(6, 10.3, '−61%', fontsize=10, color=GREEN, fontweight='bold')
    
    # Chain (element 3) - compact
    ax.text(5, 8.8, 'Цепочка нанесения:', ha='center', fontsize=8, color=DARK, fontweight='bold')
    
    # Classic chain
    for i, (label, color) in enumerate([('Ржавчина', '#EF4444'), ('ГФ-021', '#6B7280'), ('ПФ-115', '#9CA3AF')]):
        x = 1.5 + i * 1.8
        box = FancyBboxPatch((x, 7.5), 1.3, 0.8, boxstyle="round,pad=0.05",
                             facecolor=color, edgecolor='none', alpha=0.8)
        ax.add_patch(box)
        ax.text(x + 0.65, 7.9, label, ha='center', va='center', fontsize=5, color=WHITE, fontweight='bold')
        if i < 2:
            ax.annotate('', xy=(x + 1.6, 7.9), xytext=(x + 1.35, 7.9),
                        arrowprops=dict(arrowstyle='->', color=GRAY, lw=1))
    
    ax.text(5, 6.8, '3 продукта • 2 дня', ha='center', fontsize=6, color=GRAY)
    
    # TITAN chain
    box = FancyBboxPatch((2.5, 5.5), 2.5, 0.8, boxstyle="round,pad=0.05",
                         facecolor=RED, edgecolor='none', alpha=0.85)
    ax.add_patch(box)
    ax.text(3.75, 5.9, 'TITAN 3в1', ha='center', va='center', fontsize=6, color=WHITE, fontweight='bold')
    ax.annotate('', xy=(5.2, 5.9), xytext=(5.0, 5.9),
                arrowprops=dict(arrowstyle='->', color=GREEN, lw=1.5))
    ax.text(6.5, 5.9, '1 день', ha='center', fontsize=7, color=GREEN, fontweight='bold')
    
    ax.text(5, 4.8, '1 продукт • 1 день', ha='center', fontsize=6, color=GREEN, fontweight='bold')
    
    # Instruction lines
    for y in [4.0, 3.6, 3.2, 2.8]:
        ax.plot([1.5, 8.5], [y, y], color='#E0E0E0', linewidth=0.5)
    ax.text(1.5, 2.2, 'ИНСТРУКЦИЯ: подготовить поверхность...', fontsize=6, color=GRAY)
    
    # QR code (element 1)
    np.random.seed(42)
    qr_data = np.random.randint(0, 2, (6, 6))
    for i in range(6):
        for j in range(6):
            if qr_data[i, j]:
                ax.add_patch(Rectangle((6.5 + j*0.25, 1.2 + i*0.25), 0.25, 0.25, 
                                       facecolor=DARK, edgecolor='none'))
    ax.text(7.25, 0.7, 'Сканируй', ha='center', fontsize=6, color=DARK)
    ax.text(7.25, 0.4, '→ калькулятор', ha='center', fontsize=6, color=RED, fontweight='bold')
    
    fig.savefig(os.path.join(OUT_DIR, '7_full_can.png'), dpi=150, bbox_inches='tight')
    plt.close(fig)
    print('✅ Mockup 7: Full can (front + back)')


# Generate all mockups
print('Generating TCO mockups...')
mockup_qr()
mockup_one_equals_three()
mockup_chain()
mockup_price_per_year()
mockup_icons()
mockup_scale()
mockup_full_can()
print(f'\nAll mockups saved to: {OUT_DIR}')
print(f'Files: {len(os.listdir(OUT_DIR))}')
