#!/usr/bin/env python3
"""
Generate the comprehensive TCO implementation program for TITAN distribution.
Landscape (horizontal) orientation, with text + example images.
Three stages: Primary → Secondary → Tertiary sales.
"""

import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.font_manager as fm
fm.fontManager.addfont('/usr/share/fonts/truetype/chinese/LiberationSans-Regular.ttf')
fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch
import numpy as np

plt.rcParams['font.sans-serif'] = ['Liberation Sans', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

CHARTS_DIR = '/tmp/tco_program_charts'
os.makedirs(CHARTS_DIR, exist_ok=True)
OUT_PATH = '/home/z/my-project/download/Программа_внедрения_ТСО_TITAN_3_этапа.docx'

P = {
    'primary': RGBColor(0xC6, 0x28, 0x28),
    'dark': RGBColor(0x1A, 0x1F, 0x36),
    'body': RGBColor(0x2D, 0x31, 0x42),
    'secondary': RGBColor(0x5A, 0x60, 0x80),
    'accent': RGBColor(0x16, 0xA3, 0x4A),
    'warn': RGBColor(0xEA, 0x58, 0x0C),
    'blue': RGBColor(0x25, 0x63, 0xEB),
    'purple': RGBColor(0x7C, 0x3A, 0xED),
}

# Colors for charts
C_RED = '#C62828'
C_BLUE = '#2563eb'
C_GREEN = '#16a34a'
C_ORANGE = '#ea580c'
C_PURPLE = '#7c3aed'
C_DARK = '#1A1F36'
C_LIGHT = '#F5F5F5'


# ===================== CHARTS =====================

def make_funnel_chart():
    """Create a 3-stage sales funnel diagram."""
    fig, ax = plt.subplots(figsize=(14, 5), constrained_layout=True)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 6)
    ax.axis('off')

    # Three trapezoids for funnel
    stages = [
        {'y': 4, 'h': 1.8, 'w_top': 9, 'w_bot': 6.5, 'color': C_RED, 'label': 'ЭТАП 1\nПервичные продажи', 'sub': 'Завод → Дистрибьютор'},
        {'y': 2.1, 'h': 1.8, 'w_top': 6.5, 'w_bot': 4, 'color': C_BLUE, 'label': 'ЭТАП 2\nВторичные продажи', 'sub': 'Дистрибьютор → Розница'},
        {'y': 0.2, 'h': 1.8, 'w_top': 4, 'w_bot': 1.5, 'color': C_GREEN, 'label': 'ЭТАП 3\nТретичные продажи', 'sub': 'Розница → Покупатель'},
    ]

    for s in stages:
        x_top_left = (10 - s['w_top']) / 2
        x_top_right = (10 + s['w_top']) / 2
        x_bot_left = (10 - s['w_bot']) / 2
        x_bot_right = (10 + s['w_bot']) / 2

        polygon = plt.Polygon([
            (x_top_left, s['y'] + s['h']),
            (x_top_right, s['y'] + s['h']),
            (x_bot_right, s['y']),
            (x_bot_left, s['y'])
        ], color=s['color'], alpha=0.85, edgecolor='white', linewidth=2)
        ax.add_patch(polygon)

        ax.text(5, s['y'] + s['h'] / 2 + 0.15, s['label'],
                ha='center', va='center', fontsize=12, fontweight='bold', color='white')
        ax.text(5, s['y'] + s['h'] / 2 - 0.35, s['sub'],
                ha='center', va='center', fontsize=9, color='white', alpha=0.9)

    ax.text(5, 5.7, 'Воронка продаж TITAN 3в1 по методологии ТСО',
            ha='center', va='center', fontsize=14, fontweight='bold', color=C_DARK)

    fig.savefig(os.path.join(CHARTS_DIR, 'funnel.png'), dpi=150, bbox_inches='tight')
    plt.close(fig)


def make_timeline_chart():
    """Create a Gantt-style timeline for program implementation."""
    fig, ax = plt.subplots(figsize=(14, 5), constrained_layout=True)

    tasks = [
        ('Подготовка материалов', 0, 2, C_RED),
        ('Пилот: Центр Красок', 1, 4, C_BLUE),
        ('Обучение торговой команды', 2, 3, C_GREEN),
        ('Запуск POS-материалов', 3, 2, C_ORANGE),
        ('Калькулятор ТСО онлайн', 3, 2, C_PURPLE),
        ('Масштабирование на 3-5 дистр.', 5, 4, C_BLUE),
        ('Анализ результатов пилота', 5, 2, C_RED),
        ('Программа мотивации продавцов', 6, 3, C_GREEN),
        ('Оценка программы, корректировка', 9, 3, C_DARK),
    ]

    for i, (name, start, duration, color) in enumerate(tasks):
        ax.barh(i, duration, left=start, height=0.6, color=color, alpha=0.85, edgecolor='white')
        ax.text(start + duration / 2, i, name, ha='center', va='center',
                fontsize=8, fontweight='bold', color='white')

    ax.set_yticks(range(len(tasks)))
    ax.set_yticklabels([''] * len(tasks))
    ax.set_xticks(range(0, 13))
    ax.set_xticklabels([f'{i}' for i in range(13)], fontsize=9)
    ax.set_xlabel('Месяцы с момента старта', fontsize=10)
    ax.set_title('Таймлайн внедрения программы ТСО (12 месяцев)', fontsize=13, fontweight='bold', pad=10)
    ax.invert_yaxis()
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.grid(axis='x', alpha=0.2)

    fig.savefig(os.path.join(CHARTS_DIR, 'timeline.png'), dpi=150)
    plt.close(fig)


def make_stage_diagram(stage_num, title, color, items):
    """Create a stage detail diagram with practices."""
    fig, ax = plt.subplots(figsize=(13, 4), constrained_layout=True)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 5)
    ax.axis('off')

    # Title bar
    bar = FancyBboxPatch((0.3, 4), 9.4, 0.8, boxstyle="round,pad=0.1",
                         facecolor=color, edgecolor='none')
    ax.add_patch(bar)
    ax.text(5, 4.4, f'ЭТАП {stage_num}: {title}',
            ha='center', va='center', fontsize=13, fontweight='bold', color='white')

    # Items as boxes
    n = len(items)
    box_w = 9.0 / n - 0.2
    for i, (icon, label, desc) in enumerate(items):
        x = 0.5 + i * (box_w + 0.25)
        box = FancyBboxPatch((x, 0.3), box_w, 3.3, boxstyle="round,pad=0.15",
                             facecolor='white', edgecolor=color, linewidth=2)
        ax.add_patch(box)
        ax.text(x + box_w / 2, 3.2, icon, ha='center', va='center', fontsize=24)
        ax.text(x + box_w / 2, 2.5, label, ha='center', va='center',
                fontsize=10, fontweight='bold', color=color)
        ax.text(x + box_w / 2, 1.3, desc, ha='center', va='center',
                fontsize=8, color='#444', wrap=True)

    fig.savefig(os.path.join(CHARTS_DIR, f'stage{stage_num}.png'), dpi=150, bbox_inches='tight')
    plt.close(fig)


def make_kpi_chart():
    """Create KPI dashboard chart."""
    fig, axes = plt.subplots(1, 3, figsize=(14, 4), constrained_layout=True)

    stages = ['Этап 1\nПервичные', 'Этап 2\nВторичные', 'Этап 3\nТретичные']
    colors = [C_RED, C_BLUE, C_GREEN]

    # KPI 1: Volume growth
    ax = axes[0]
    values = [30, 25, 20]
    bars = ax.bar(stages, values, color=colors, alpha=0.85, edgecolor='white')
    ax.set_title('Рост объёмов продаж TITAN 3в1\n(целевой прирост %)', fontsize=11, fontweight='bold')
    ax.set_ylabel('% прироста', fontsize=9)
    for bar, v in zip(bars, values):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 1,
                f'+{v}%', ha='center', fontsize=10, fontweight='bold')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.set_ylim(0, 40)

    # KPI 2: Training coverage
    ax = axes[1]
    values = [100, 80, 60]
    bars = ax.bar(stages, values, color=colors, alpha=0.85, edgecolor='white')
    ax.set_title('Охват обучением ТСО\n(% целевой аудитории)', fontsize=11, fontweight='bold')
    ax.set_ylabel('% охвата', fontsize=9)
    for bar, v in zip(bars, values):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 2,
                f'{v}%', ha='center', fontsize=10, fontweight='bold')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.set_ylim(0, 120)

    # KPI 3: TCO adoption
    ax = axes[2]
    values = [90, 75, 50]
    bars = ax.bar(stages, values, color=colors, alpha=0.85, edgecolor='white')
    ax.set_title('Применение ТСО в продажах\n(% использующих методику)', fontsize=11, fontweight='bold')
    ax.set_ylabel('% применения', fontsize=9)
    for bar, v in zip(bars, values):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 2,
                f'{v}%', ha='center', fontsize=10, fontweight='bold')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.set_ylim(0, 110)

    fig.savefig(os.path.join(CHARTS_DIR, 'kpi.png'), dpi=150)
    plt.close(fig)


# Generate all charts
print('Generating charts...')
make_funnel_chart()
make_timeline_chart()
make_stage_diagram(1, 'Первичные продажи (Завод → Дистрибьютор)', C_RED, [
    ('$', 'Бонусы', '8% ежемесячно +\n3% полугодовой'),
    ('D', 'Отсрочка', '60 дней вместо\nстандартных 30'),
    ('T', 'Обучение', 'Выездное обучение\nкоманды дистрибьютора'),
    ('F', 'Экскурсия', 'Поездка на завод\nЭмпилс'),
])
make_stage_diagram(2, 'Вторичные продажи (Дистрибьютор → Розница)', C_BLUE, [
    ('+', 'Мотивация ТП', 'Бонусы для торговых\nпредставителей'),
    ('*', 'Инструменты', 'Калькулятор ТСО,\nпрезентации, образцы'),
    ('%', 'Программы', 'Розничные программы\nлояльности'),
    ('K', 'KPI', 'Метрики охвата\nторговых точек'),
])
make_stage_diagram(3, 'Третичные продажи (Розница → Покупатель)', C_GREEN, [
    ('Q', 'QR на банке', 'Видео-инструкция\n+ калькулятор ТСО'),
    ('P', 'POS-материалы', 'Шпаргалки, выкрасы,\nсравнительные таблицы'),
    ('L', 'Обучение', 'Продавцы умеют\nобъяснить ТСО за 2 мин'),
    ('=', 'Калькулятор', 'Онлайн-расчёт ТСО\nдля покупателя'),
])
make_kpi_chart()
print('Charts generated!')


# ===================== DOCX =====================

def set_cell_bg(cell, color_hex):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tc_pr.append(shd)


def set_cell_borders(cell, color='CCCCCC', size='4'):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = OxmlElement('w:tcBorders')
    for side in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{side}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), size)
        border.set(qn('w:color'), color)
        tc_borders.append(border)
    tc_pr.append(tc_borders)


def add_heading(doc, text, level=1, color=None):
    if color is None:
        color = P['primary'] if level == 1 else P['dark']
    h = doc.add_heading('', level=level)
    run = h.add_run(text)
    run.font.color.rgb = color
    run.font.name = 'Calibri'
    return h


def add_para(doc, text, size=11, bold=False, italic=False, color=None, indent=False):
    p = doc.add_paragraph()
    if indent:
        p.paragraph_format.first_line_indent = Cm(0.5)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.25
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = 'Calibri'
    if color:
        run.font.color.rgb = color
    return p


def add_bullet(doc, text):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.2
    run = p.add_run(text)
    run.font.size = Pt(10)
    run.font.name = 'Calibri'


def add_image(doc, path, width_cm=24):
    if os.path.exists(path):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(path, width=Cm(width_cm))


def make_table(doc, headers, rows, col_widths=None, header_color='C62828'):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    if col_widths:
        for i, width in enumerate(col_widths):
            for cell in table.columns[i].cells:
                cell.width = Cm(width)
    header_cells = table.rows[0].cells
    for i, header in enumerate(headers):
        cell = header_cells[i]
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(header)
        run.font.bold = True
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        run.font.name = 'Calibri'
        set_cell_bg(cell, header_color)
        set_cell_borders(cell)
    for r_idx, row_data in enumerate(rows):
        row_cells = table.rows[r_idx + 1].cells
        for c_idx, cell_text in enumerate(row_data):
            cell = row_cells[c_idx]
            cell.text = ''
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            run = p.add_run(str(cell_text))
            run.font.size = Pt(9)
            run.font.name = 'Calibri'
            set_cell_borders(cell)
            if r_idx % 2 == 0:
                set_cell_bg(cell, 'FAF5F5')
    return table


# ====== Build Document ======
doc = Document()
style = doc.styles['Normal']
style.font.name = 'Calibri'
style.font.size = Pt(11)
style.paragraph_format.line_spacing = 1.25

# Set LANDSCAPE orientation
section = doc.sections[0]
section.orientation = WD_ORIENT.LANDSCAPE
section.page_width = Cm(29.7)   # A4 landscape width
section.page_height = Cm(21.0)  # A4 landscape height
section.left_margin = Cm(2.0)
section.right_margin = Cm(2.0)
section.top_margin = Cm(1.5)
section.bottom_margin = Cm(1.5)

# ===== COVER =====
cover_table = doc.add_table(rows=1, cols=1)
cover_cell = cover_table.rows[0].cells[0]
set_cell_bg(cover_cell, 'C62828')
cover_cell.text = ''
p = cover_cell.paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(30)
p.paragraph_format.space_after = Pt(10)
run = p.add_run('TITAN | Эмпилс')
run.font.size = Pt(36)
run.font.bold = True
run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

p2 = cover_cell.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
p2.paragraph_format.space_after = Pt(20)
run = p2.add_run('Сводная программа внедрения методологии ТСО\nв сеть дистрибуции')
run.font.size = Pt(22)
run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

p3 = cover_cell.add_paragraph()
p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
p3.paragraph_format.space_after = Pt(30)
run = p3.add_run('Совокупная стоимость владения как инструмент продаж TITAN 3в1\n'
                  'на трёх этапах: первичные → вторичные → третичные продажи')
run.font.size = Pt(14)
run.font.italic = True
run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Три этапа программы')
run.font.size = Pt(16)
run.font.bold = True
run.font.color.rgb = P['dark']

stages_summary = doc.add_table(rows=1, cols=3)
stages_summary.alignment = WD_TABLE_ALIGNMENT.CENTER
stage_info = [
    ('ЭТАП 1', 'Первичные продажи', 'Завод → Дистрибьютор', 'C62828'),
    ('ЭТАП 2', 'Вторичные продажи', 'Дистрибьютор → Розница', '2563eb'),
    ('ЭТАП 3', 'Третичные продажи', 'Розница → Покупатель', '16a34a'),
]
for i, (num, title, sub, color) in enumerate(stage_info):
    cell = stages_summary.rows[0].cells[i]
    cell.text = ''
    set_cell_bg(cell, color)
    set_cell_borders(cell)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(15)
    p.paragraph_format.space_after = Pt(5)
    run = p.add_run(num)
    run.font.size = Pt(20)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    p2 = cell.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p2.add_run(title)
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    p3 = cell.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.paragraph_format.space_after = Pt(15)
    run = p3.add_run(sub)
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Empils • TITAN • 2026')
run.font.size = Pt(12)
run.font.color.rgb = P['secondary']

doc.add_page_break()

# ===== 1. КОНЦЕПЦИЯ =====
add_heading(doc, '1. Концепция программы', level=1)

add_para(doc,
    'Методология ТСО (Total Cost of Ownership — Совокупная стоимость владения) — это подход '
    'к продаже лакокрасочных материалов, при котором покупателю демонстрируется не цена банки, '
    'а полная стоимость покраски: материалы + грунт + труд + время + повторные покрытия. '
    'Расчёт ТСО показывает, что грунт-эмаль TITAN 3в1, несмотря на более высокую цену банки, '
    'выгоднее дешёвых аналогов (ПФ-115 + грунт) за счёт меньшего количества слоёв, экономии '
    'времени и большего срока службы покрытия (до 10 лет).', indent=True)

add_para(doc,
    'Программа внедряет ТСО-методологию на трёх уровнях цепочки продаж: от завода-производителя '
    'через дистрибьютора и розничную точку до конечного покупателя. Каждый уровень требует '
    'своих инструментов, мотивации и обучения. Успех программы зависит от синхронной работы '
    'всех трёх этапов — провал на любом из них обнуляет усилия остальных.', indent=True)

# Funnel diagram
add_image(doc, os.path.join(CHARTS_DIR, 'funnel.png'), width_cm=24)

add_para(doc,
    'Программа рассчитана на 12 месяцев и охватывает пилотных дистрибьюторов (1–2 на старте, '
    'масштабирование на 3–5 с 5-го месяца). Базовые пилотные партнёры — «Центр Красок» '
    '(Краснодар, ЮФО) и «Лемана ПРО» (Ростов-на-Дону).', indent=True, italic=True,
    color=P['secondary'])

doc.add_page_break()

# ===== 2. ЭТАП 1: ПЕРВИЧНЫЕ ПРОДАЖИ =====
add_heading(doc, '2. Этап 1. Первичные продажи: Завод → Дистрибьютор', level=1, color=P['primary'])

add_para(doc,
    'Задача этапа — стимулировать дистрибьютора закупать TITAN 3в1 у «Эмпилс» в необходимых '
    'объёмах и поддерживать ассортимент на складе. Дистрибьютор должен видеть прямую '
    'финансовую выгоду от работы с TITAN, а также получать поддержку, компенсирующую '
    'маркетинговые инвестиции конкурентов (AkzoNobel/Dulux, Hammerite).', indent=True)

# Stage 1 diagram
add_image(doc, os.path.join(CHARTS_DIR, 'stage1.png'), width_cm=25)

add_heading(doc, '2.1. Финансовые стимулы', level=2, color=P['dark'])

fin1_data = [
    ['Ежемесячный бонус', '8% от объёма закупок TITAN 3в1', 'Безоговорочный, выплата в течение 10 раб. дней'],
    ['Полугодовой бонус', '3% от объёма закупок за полугодие', 'Выплата по итогам полугодия'],
    ['Отсрочка платежа', '60 календарных дней (вместо 30)', 'При соблюдении графика платежей'],
    ['Программа объёмных скидок', 'Прогрессивная шкала: 3% / 5% / 7%', 'При превышении квартального минимума закупок'],
    ['Субсидирование маркетинга', 'До 50% затрат на POS и промо', 'По согласованию, с отчётностью'],
]
make_table(doc, ['Инструмент', 'Условие', 'Порядок'], fin1_data, col_widths=[5, 8, 10])

add_heading(doc, '2.2. Нефинансовые стимулы', level=2, color=P['dark'])

add_para(doc, 'Обучение команды дистрибьютора:', bold=True, color=P['primary'])
add_bullet(doc, 'Выездное обучение торговой команды на территории дистрибьютора (офлайн, 8 часов)')
add_bullet(doc, 'Модули: о компании Эмпилс → продукты TITAN (упор на 3в1) → методика ТСО → практикум с выкрасами → работа с возражениями → тестирование')
add_bullet(doc, 'Материалы: шпаргалка продавца, сравнительная таблица TITAN vs Hammerite vs ПФ-115, калькулятор ТСО, сертификаты участников')

add_para(doc, 'Экскурсия на завод «Эмпилс»:', bold=True, color=P['primary'])
add_bullet(doc, 'Делегация до 10 человек (руководители, торговые представители, ключевые продавцы)')
add_bullet(doc, 'Программа: экскурсия по производству, лаборатория качества, дегустация продуктов, круглый стол с руководством')
add_bullet(doc, 'Цель: формирование лояльности и глубокого понимания продукта')

add_para(doc, 'Партнёрский статус:', bold=True, color=P['primary'])
add_bullet(doc, 'Присвоение статуса «Официальный партнёр ТСО» с сертификатом')
add_bullet(doc, 'Приоритет в поставках новинок и ограниченных серий')
add_bullet(doc, 'Прямой доступ к техническому специалисту Эмпилс для консультаций')

add_heading(doc, '2.3. Пример: Программа для «Центр Красок» (Краснодар)', level=2, color=P['dark'])

add_para(doc,
    'На основе протокола встречи от 01.07.2026 разработана индивидуальная программа: '
    'бонус 8% + 3%, отсрочка 60 дней, выездное обучение в июле—августе, экскурсия на завод '
    'в сентябре—октябре, обеспечение выкрасами. Ожидаемый результат — рост закупок TITAN 3в1 '
    'на 20–30% за полугодие и восстановление сокращённого ассортимента.',
    indent=True, italic=True, color=P['secondary'])

doc.add_page_break()

# ===== 3. ЭТАП 2: ВТОРИЧНЫЕ ПРОДАЖИ =====
add_heading(doc, '3. Этап 2. Вторичные продажи: Дистрибьютор → Розничные точки', level=1, color=P['blue'])

add_para(doc,
    'Задача этапа — вооружить торговых представителей (ТП) дистрибьютора инструментами '
    'и мотивацией для активного предложения TITAN 3в1 в розничные точки. ТП — это главный '
    'канал коммуникации с розницей: от их убедительности зависит, попадёт ли TITAN на полку '
    'и в каком объёме. Методология ТСО даёт ТП сильный аргумент: «TITAN 3в1 выгоднее для '
    'вашего покупателя, а значит — выше оборот и меньше возвратов».', indent=True)

# Stage 2 diagram
add_image(doc, os.path.join(CHARTS_DIR, 'stage2.png'), width_cm=25)

add_heading(doc, '3.1. Мотивация торговых представителей', level=2, color=P['dark'])

motiv_data = [
    ['Бонус за новую точку', '500–1000 руб. за каждую новую торговую точку, взявшую TITAN 3в1', 'Выплата ежемесячно, при подтверждении закупки точкой'],
    ['Бонус за оборот', '2–3% от объёма продаж TITAN 3в1 в зоне ответственности ТП', 'Ежемесячно, при достижении минимального порога'],
    ['Квартальный конкурс', 'Топ-3 ТП по продажам TITAN 3в1 — премия 10 000 / 7 000 / 5 000 руб.', 'Ежеквартально, публичное награждение'],
    ['Сертификация ТСО', 'Бонус 3 000 руб. за сданный экзамен по ТСО-методологии', 'Однократно, после обучения'],
]
make_table(doc, ['Инструмент', 'Условие', 'Порядок'], motiv_data, col_widths=[5, 9, 9])

add_heading(doc, '3.2. Инструменты для торговых представителей', level=2, color=P['dark'])

tools_data = [
    ['1', 'Калькулятор ТСО (мобильный)', 'Онлайн-калькулятор на телефоне: ввод площади → расчёт стоимости TITAN vs ПФ-115+грунт. Показывает экономию в рублях и процентах.'],
    ['2', 'Шпаргалка продавца TITAN 3в1', 'Карточка А5 с 5 ключевыми аргументами ТСО, сравнением с конкурентами, ответами на топ-5 возражений. Ламинированная — для кармана.'],
    ['3', 'Планшетная презентация', '5-слайдовая презентация для показа закупщику розничной точки: TITAN 3в1 — выгода для точки (маржа, оборачиваемость, меньше возвратов).'],
    ['4', 'Набор выкрасов', 'Портативный набор: TITAN 3в1 (5 цветов RAL) + сравнение с Hammerite и ПФ-115. Для демонстрации прямо в торговой точке.'],
    ['5', 'Сравнительная таблица', 'A4: TITAN 3в1 vs ПФ-115+ГФ-021 vs Hammerite — по 10 параметрам: цена, слои, время, расход, срок службы, ТСО за 5 лет.'],
    ['6', 'Кейсы ТСО для розницы', '3 готовых расчёта: «Ворота 6 м²», «Гараж 20 м²», «Забор 50 м²» — для показа закупщику выгоды для его покупателя.'],
]
make_table(doc, ['№', 'Инструмент', 'Описание'], tools_data, col_widths=[1, 5, 17])

add_heading(doc, '3.3. Программы для розничных точек', level=2, color=P['dark'])

add_para(doc, 'Программы, которые ТП предлагают закупщикам розничных точек:', indent=True)
add_bullet(doc, '🎁 Стартовый набор TITAN: при первой закупке 3в1 — бесплатные выкрасы + шпаргалка + QR-стенд для полки')
add_bullet(doc, '💰 Скидка за ассортимент: при заказе 3+ цветов RAL — дополнительная скидка 3%')
add_bullet(doc, '📈 Бонус за оборот: при достижении квартального минимума продаж 3в1 — бонус 2% товаром')
add_bullet(doc, '🔄 Возврат неликвида: возможность вернуть нереализованный TITAN 3в1 в течение 90 дней (в пределах срока годности)')
add_bullet(doc, '📱 QR-стенд на полку: бесплатный стенд с QR-кодом → видео-инструкция + калькулятор ТСО для покупателя')

add_heading(doc, '3.4. Пример: Работа ТП с новой точкой', level=2, color=P['dark'])

add_para(doc,
    'Торговый представитель заходит в новый магазин строительных материалов. Показывает '
    'на планшете 5-слайдовую презентацию: «TITAN 3в1 — почему ваш покупатель выберет '
    'именно его». Достаёт шпаргалку с расчётом ТСО для ворот 6 м²: ПФ-115 + грунт = '
    '4 200 руб. и 2 дня работы, TITAN 3в1 = 3 800 руб. и 1 день. Показывает выкрасы — '
    'TITAN vs ПФ-115. Предлагает стартовый набор (бесплатные выкрасы + QR-стенд). '
    'Закупщик видит: выше маржа, меньше возвратов, покупатель доволен. Оформляет первую '
    'закупку на 2 цвета RAL. ТП получает бонус 500 руб. за новую точку.',
    indent=True, italic=True, color=P['secondary'])

doc.add_page_break()

# ===== 4. ЭТАП 3: ТРЕТИЧНЫЕ ПРОДАЖИ =====
add_heading(doc, '4. Этап 3. Третичные продажи: Розница → Покупатель', level=1, color=P['accent'])

add_para(doc,
    'Это ключевой этап — момент, когда конечный покупатель стоит у полки и выбирает между '
    'TITAN 3в1, ПФ-115 и Hammerite. От того, насколько хорошо продавец-консультант объяснит '
    'выгоду ТСО, зависит решение покупателя. На этом этапе сконцентрирована основная часть '
    'ранее разработанных материалов Эмпилс: DIY-гайды, инфографика, QR-коды, калькулятор ТСО.',
    indent=True)

# Stage 3 diagram
add_image(doc, os.path.join(CHARTS_DIR, 'stage3.png'), width_cm=25)

add_heading(doc, '4.1. Материалы для полки', level=2, color=P['dark'])

shelf_data = [
    ['1', 'QR-код на банке TITAN 3в1', 'Покупатель сканирует → видео 2 мин «Как наносить 3в1» + калькулятор ТСО. Весь цикл — от сканирования до расчёта — 30 секунд.'],
    ['2', 'Шпаргалка продавца (на полке)', 'Карточка А6: 5 шагов продажи TITAN 3в1 через ТСО. Вопросы покупателю → расчёт → сравнение → закрытие сделки. Крепится к полке.'],
    ['3', 'Сравнительная таблица (воблер)', 'Воблер на полке: «TITAN 3в1: 1 банка вместо 2 + экономия времени. Полная стоимость покраски — дешевле!»'],
    ['4', 'Выкрасы (до/после, через год)', 'Набор образцов на полке: TITAN 3в1 vs ПФ-115 — наглядное сравнение покрытия. Покупатель видит разницу своими глазами.'],
    ['5', 'Листовка ТСО для покупателя', 'А5: «Почему TITAN 3в1 выгоднее? Расчёт на примере ворот, гаража, забора». Покупатель забирает домой.'],
    ['6', 'Видео-экран (в крупных точках)', 'Видео 30 сек на экране у полки: ускоренная покраска ворот TITAN 3в1 vs ПФ-115+грунт.'],
]
make_table(doc, ['№', 'Материал', 'Описание и механика'], shelf_data, col_widths=[1, 5, 17])

add_heading(doc, '4.2. Обучение продавцов-консультантов', level=2, color=P['dark'])

add_para(doc, 'Формат обучения продавцов розничных точек:', indent=True)
add_bullet(doc, 'Офлайн-сессии 45–60 мин в магазине (выезд тренера Эмпилс или тренера дистрибьютора)')
add_bullet(doc, 'Онлайн-платформа: 10 видеоуроков по 3–5 мин + тест (по аналогии с AkzoNobel Ice Free Lion)')
add_bullet(doc, 'Мотивация: +3 000 руб. за сданный тест, −3 000 руб. за несданный (в зачёт премии продавца)')
add_bullet(doc, 'Сертификат «Продавец ТСО» — подтверждение квалификации, действует 1 год')
add_bullet(doc, 'Регулярность: обучение 1 раз в год + обновление при выходе новинок')

add_para(doc, 'Ключевой навык продавца — «объяснить ТСО за 2 минуты»:', bold=True, color=P['primary'])
add_para(doc,
    '1) Вопрос: «Что красите?» → 2) Расчёт площади → 3) Сравнение TITAN 3в1 vs ПФ-115+грунт '
    'на калькуляторе → 4) Показ выкраса → 5) Аргумент: «Полная стоимость покраски дешевле, '
    'а хватит на 10 лет». Весь цикл — 2 минуты.', indent=True)

add_heading(doc, '4.3. Калькулятор ТСО для покупателя', level=2, color=P['dark'])

add_para(doc,
    'Онлайн-калькулятор на сайте titanpaints.ru и по QR-коду на банке. Покупатель вводит: '
    'тип поверхности (ворота, гараж, забор, трубы), площадь, текущее состояние (ржавчина/ '
    'чистый металл). Калькулятор выдаёт: расход TITAN 3в1, количество банок, стоимость '
    'материалов, сравнение с ПФ-115+грунт, экономию в рублях и времени, срок службы покрытия. '
    'Результат можно распечатать или отправить в WhatsApp.', indent=True)

add_heading(doc, '4.4. Пример: Покупатель в магазине', level=2, color=P['dark'])

add_para(doc,
    'Покупатель заходит в магазин: «Нужно покрасить ржавые ворота на даче, 6 квадратных метров». '
    'Продавец: «Возьмите грунт-эмаль TITAN 3в1 — красит прямо по ржавчине, не нужно '
    'отдельно грунтовать. Давайте посчитаем». Достаёт калькулятор: TITAN 3в1 — 1 банка, '
    '2 слоя, 1 день работы, 3 800 руб. ПФ-115 + грунт ГФ-021 — 2 банки, 3 слоя, 2 дня, '
    '4 200 руб. «TITAN дешевле на 400 рублей, сэкономит вам целый день, а держаться будет '
    '10 лет. Вот, посмотрите образец». Показывает выкрас. Покупатель: «Беру TITAN».',
    indent=True, italic=True, color=P['secondary'])

doc.add_page_break()

# ===== 5. МАТРИЦА ПРАКТИК =====
add_heading(doc, '5. Сводная матрица практик по этапам', level=1)

matrix_data = [
    ['1', 'Бонус 8% + 3%', '—', '—', 'Стимулирование закупки дистрибьютором'],
    ['1', 'Отсрочка 60 дней', '—', '—', 'Снижение финансовой нагрузки на дистрибьютора'],
    ['1', 'Выездное обучение команды', 'Мотивация ТП (бонусы)', 'Обучение продавцов розницы', 'Знания и навыки на всех уровнях'],
    ['1', 'Экскурсия на завод', '—', '—', 'Лояльность руководства дистрибьютора'],
    ['1', 'Субсидирование маркетинга', 'Инструменты для ТП', 'POS-материалы на полку', 'Финансовая поддержка на всех уровнях'],
    ['—', '—', 'Бонус за новую точку', 'QR-стенд на полку', 'Расширение присутствия TITAN'],
    ['—', '—', 'Калькулятор ТСО для ТП', 'Калькулятор ТСО для покупателя', 'ТСО-аргумент на двух уровнях'],
    ['—', '—', 'Квартальный конкурс ТП', 'Конкурс «Продавец ТСО»', 'Мотивация через соревнование'],
    ['—', '—', 'Набор выкрасов для ТП', 'Выкрасы на полке (до/после)', 'Наглядная демонстрация качества'],
    ['—', '—', 'Презентация для закупщика', 'Шпаргалка продавца', 'Инструмент убеждения'],
    ['—', '—', '—', 'Сертификат «Продавец ТСО»', 'Подтверждение квалификации'],
    ['—', '—', '—', 'Листовка ТСО для покупателя', 'ТСО-аргумент забирается домой'],
]
make_table(doc, ['Этап', 'Первичные (Завод→Дистр.)', 'Вторичные (Дистр.→Розница)', 'Третичные (Розница→Покупатель)', 'Цель'],
           matrix_data, col_widths=[1.2, 5, 5.5, 5.5, 5])

doc.add_page_break()

# ===== 6. ТАЙМЛАЙН =====
add_heading(doc, '6. Таймлайн внедрения (12 месяцев)', level=1)

add_image(doc, os.path.join(CHARTS_DIR, 'timeline.png'), width_cm=24)

add_para(doc, 'Ключевые вехи:', bold=True, color=P['primary'])
add_bullet(doc, 'Месяц 0–2: подготовка всех материалов (калькулятор, шпаргалки, выкрасы, QR-коды, обучающие видео)')
add_bullet(doc, 'Месяц 1–4: пилот на 1–2 дистрибьюторах (Центр Красок + Лемана ПРО)')
add_bullet(doc, 'Месяц 2–4: выездное обучение торговых команд пилотных дистрибьюторов')
add_bullet(doc, 'Месяц 3–4: запуск POS-материалов и калькулятора ТСО онлайн')
add_bullet(doc, 'Месяц 5–8: масштабирование на 3–5 дистрибьюторов')
add_bullet(doc, 'Месяц 6–8: программа мотивации продавцов розницы (сертификация ТСО)')
add_bullet(doc, 'Месяц 9–12: анализ результатов, корректировка, принятие решения о полном масштабировании')

doc.add_page_break()

# ===== 7. KPI =====
add_heading(doc, '7. KPI и метрики по этапам', level=1)

add_image(doc, os.path.join(CHARTS_DIR, 'kpi.png'), width_cm=24)

kpi_data = [
    ['1', 'Объём закупок TITAN 3в1 дистрибьютором', '+30% к базовому периоду', 'Ежемесячно', 'Отчёты Эмпилс'],
    ['1', 'Доля TITAN 3в1 в ассортименте дистрибьютора', 'Не менее 5 цветов RAL в наличии', 'Ежеквартально', 'Аудит склада'],
    ['1', 'Соблюдение отсрочки платежа', '0 просрочек > 60 дней', 'Ежемесячно', 'Бухгалтерия'],
    ['2', 'Количество активных розничных точек с TITAN 3в1', '+25% к базовому периоду', 'Ежемесячно', 'Отчёты ТП'],
    ['2', 'Охват торговых представителей обучением ТСО', '100% ТП пилотного дистрибьютора', 'К мес. 3', 'Журнал обучения'],
    ['2', 'Доля новых точек от TITAN (от всех новых)', 'Не менее 20%', 'Ежемесячно', 'CRM дистрибьютора'],
    ['3', 'Количество розничных точек с POS-материалами', '80% точек с TITAN на полке', 'К мес. 6', 'Аудит полок'],
    ['3', 'Количество сертифицированных продавцов ТСО', 'Не менее 20 на пилотного дистрибьютора', 'К мес. 6', 'Журнал сертификации'],
    ['3', 'Конверсия в покупку TITAN (из тех, кто спросил ЛКМ)', 'Не менее 30%', 'Ежеквартально', 'Опрос точек'],
    ['3', 'Сканирования QR-кода на банке', '500+ сканов/мес на пилотного дистрибьютора', 'Ежемесячно', 'Аналитика QR'],
    ['3', 'Снижение возвратов TITAN 3в1', '−30% к базовому периоду', 'Ежеквартально', 'Отчёты дистрибьютора'],
]
make_table(doc, ['Этап', 'Метрика', 'Целевое значение', 'Периодичность', 'Источник данных'],
           kpi_data, col_widths=[1.2, 6, 5.5, 3, 5])

doc.add_page_break()

# ===== 8. БЮДЖЕТ =====
add_heading(doc, '8. Ориентировочный бюджет программы (пилот, 12 месяцев)', level=1)

budget_data = [
    ['1', 'Финансовые бонусы дистрибьютору (8%+3%)', 'Зависит от объёма. Пример: 10 млн/мес → 660 000 за полугодие', 'В счёт будущих поставок'],
    ['1', 'Командировочные тренера (выездное обучение)', '50 000 – 80 000 на 2 пилотных дистрибьютора', 'Проезд, проживание, питание'],
    ['1', 'Экскурсия на завод (2 делегации)', '100 000 – 150 000 (программа, обед, сувениры)', 'Проезд 50/50 с партнёром'],
    ['2', 'Мотивация ТП (бонусы, конкурсы)', '200 000 – 300 000 на 2 пилота за год', 'При достижении KPI'],
    ['2', 'Инструменты для ТП (шпаргалки, выкрасы, планшеты)', '100 000 – 150 000', 'Тиражирование материалов'],
    ['3', 'POS-материалы (QR-стенды, воблеры, листовки)', '150 000 – 250 000', 'Тираж на пилотные точки'],
    ['3', 'Разработка калькулятора ТСО (онлайн)', '150 000 – 250 000 (разовый)', 'Разработка + хостинг'],
    ['3', 'Обучение продавцов (видео + платформа)', '100 000 – 200 000', 'Производство видео + платформа'],
    ['3', 'QR-коды на этикетках', 'Включено в стоимость этикетки', 'Переутверждение макета'],
    ['—', 'Аналитика и контроль (аудит полок, опросы)', '50 000 – 80 000', 'Сбор данных, отчёты'],
]
make_table(doc, ['Этап', 'Статья', 'Сумма (руб.)', 'Примечание'], budget_data, col_widths=[1.2, 7, 7, 8])

add_para(doc, '')
add_para(doc,
    'Итого ориентировочный бюджет пилотной программы на 2 дистрибьютора: '
    '1 500 000 – 2 200 000 руб. на 12 месяцев (без учёта бонусов дистрибьютору, '
    'которые выплачиваются из маржи и зависят от объёма).',
    bold=True, color=P['primary'])

add_para(doc,
    'ROI программы: при росте продаж TITAN 3в1 на 25% (минимальный сценарий) '
    'дополнительная выручка Эмпилс окупает бюджет программы за 6–8 месяцев.',
    italic=True, color=P['secondary'])

doc.add_page_break()

# ===== 9. РЕКОМЕНДАЦИИ =====
add_heading(doc, '9. Рекомендации и следующие шаги', level=1)

add_heading(doc, '9.1. Ключевые факторы успеха', level=2, color=P['dark'])
add_bullet(doc, '✅ Синхронность: все три этапа запускаются одновременно — провал на одном обнуляет остальные')
add_bullet(doc, '✅ Калькулятор ТСО — это центральный инструмент, работает на всех трёх уровнях (для ТП, для продавца, для покупателя)')
add_bullet(doc, '✅ Обучение — не разовое мероприятие, а непрерывный процесс с тестированием и сертификацией')
add_bullet(doc, '✅ Маркетинговая поддержка должна быть не меньше, чем у конкурентов (AkzoNobel инвестирует агрессивно)')
add_bullet(doc, '✅ Метрики должны собираться регулярно — без данных невозможно корректировать программу')

add_heading(doc, '9.2. Следующие шаги', level=2, color=P['dark'])
add_bullet(doc, '1. Утвердить программу и бюджет (июль 2026)')
add_bullet(doc, '2. Разработать калькулятор ТСО онлайн (июль—август 2026)')
add_bullet(doc, '3. Произвести POS-материалы и обучающие видео (июль—август 2026)')
add_bullet(doc, '4. Запустить пилот на «Центр Красок» и «Лемана ПРО» (август 2026)')
add_bullet(doc, '5. Провести выездное обучение торговых команд (август—сентябрь 2026)')
add_bullet(doc, '6. Первый аудит результатов — ноябрь 2026, корректировка программы')
add_bullet(doc, '7. Решение о масштабировании — январь 2027')

# Final note
doc.add_paragraph()
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(15)
run = p.add_run('Empils • TITAN • Программа внедрения ТСО-методологии в дистрибуцию • 2026')
run.font.size = Pt(10)
run.font.italic = True
run.font.color.rgb = P['secondary']

# Save
doc.save(OUT_PATH)
print(f'Document saved: {OUT_PATH}')
print(f'Size: {os.path.getsize(OUT_PATH)} bytes')
