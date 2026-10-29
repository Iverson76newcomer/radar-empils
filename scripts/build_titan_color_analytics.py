"""
Аналитика по цветам TITAN — таблицы процентов + визуализация.
Читает Excel, считает %, создаёт 6 matplotlib диаграмм, собирает Word документ.
"""
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import numpy as np
from openpyxl import load_workbook
from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

# ---------- Шрифты ----------
fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf')
fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf')
plt.rcParams.update({
    'font.sans-serif': ['DejaVu Sans'],
    'axes.unicode_minus': False,
    'figure.facecolor': '#FFFFFF',
    'axes.facecolor': '#FFFFFF',
    'axes.edgecolor': '#E5E7EB',
    'axes.linewidth': 0.8,
    'axes.spines.top': False,
    'axes.spines.right': False,
    'axes.grid': False,
    'xtick.major.size': 0,
    'ytick.major.size': 0,
    'xtick.labelsize': 10,
    'ytick.labelsize': 10,
    'axes.labelsize': 11,
    'axes.titlesize': 16,
    'axes.titleweight': 'bold',
    'axes.titlepad': 20,
    'legend.frameon': False,
    'legend.fontsize': 10,
    'figure.dpi': 200,
    'savefig.dpi': 200,
    'savefig.bbox': 'tight',
    'savefig.facecolor': '#FFFFFF',
    'savefig.pad_inches': 0.3,
})

# ---------- Цвета ----------
ORANGE = '#F47629'
ORANGE_DARK = '#D45F1A'
ORANGE_LIGHT = '#FBA668'
ORANGE_PALE = '#FDD9B8'
GREY = '#D1D5DB'
GREY_DARK = '#6B7280'
DARK = '#313131'
WHITE = '#FFFFFF'

# Палитра для bar charts (топ-5 оранжевых, остальные серые)
TOP5_COLORS = [ORANGE_DARK, ORANGE, ORANGE_LIGHT, '#F5C896', GREY]
ALL_GREY = [GREY] * 10

# ---------- Чтение данных ----------
wb = load_workbook('/home/z/my-project/upload/Analitika_Titan_tsveta_15.07.2026.xlsx', data_only=True)
ws = wb['Общая по цветам']

# Парсим данные — 3 категории
categories = {}
current_cat = None
for row in ws.iter_rows(min_row=3, values_only=True):
    cat, color, qty, kg, akb = row
    if cat and cat != 'Итог':
        current_cat = cat
        categories[current_cat] = []
    if color and cat != 'Итог' and not (isinstance(cat, str) and cat == 'Итог'):
        categories[current_cat].append({
            'color': color,
            'qty': qty or 0,
            'kg': kg or 0,
            'akb': akb or 0,
        })

# Считаем итоги
for cat, items in categories.items():
    total_qty = sum(i['qty'] for i in items)
    total_kg = sum(i['kg'] for i in items)
    total_akb = sum(i['akb'] for i in items)
    for i in items:
        i['pct_qty'] = (i['qty'] / total_qty * 100) if total_qty else 0
        i['pct_kg'] = (i['kg'] / total_kg * 100) if total_kg else 0
        i['pct_akb'] = (i['akb'] / total_akb * 100) if total_akb else 0
    cat_data = {
        'items': items,
        'total_qty': total_qty,
        'total_kg': total_kg,
        'total_akb': total_akb,
    }
    categories[cat] = cat_data

print('Данные загружены:')
for cat, d in categories.items():
    print(f'  {cat}: {len(d["items"])} цветов, {d["total_qty"]} шт, {d["total_kg"]:.0f} кг, {d["total_akb"]} АКБ')

# ---------- Генерация диаграмм ----------
CHART_DIR = '/home/z/my-project/scripts/charts_titan'
os.makedirs(CHART_DIR, exist_ok=True)

chart_files = {}

for cat_name, data in categories.items():
    items = data['items']
    # Сортируем по убыванию qty
    items_sorted = sorted(items, key=lambda x: x['qty'], reverse=True)
    labels = [i['color'] for i in items_sorted]
    pcts = [i['pct_qty'] for i in items_sorted]
    qtys = [i['qty'] for i in items_sorted]

    # --- Bar chart (горизонтальный) ---
    fig, ax = plt.subplots(figsize=(10, 6), constrained_layout=True)
    y_pos = np.arange(len(labels))
    # Топ-5 — оранжевый, остальные — серый
    bar_colors = [ORANGE if i < 5 else GREY for i in range(len(labels))]
    bars = ax.barh(y_pos, pcts, color=bar_colors, height=0.6, edgecolor='white', linewidth=0.5)
    ax.set_yticks(y_pos)
    ax.set_yticklabels(labels, fontsize=10)
    ax.invert_yaxis()  # топ-1 сверху
    ax.set_xlabel('Доля продаж, %', fontsize=11, color=DARK)
    ax.set_xlim(0, max(pcts) * 1.25)

    # Подписи процентов на барах
    for bar, pct, qty in zip(bars, pcts, qtys):
        w = bar.get_width()
        ax.text(w + 0.5, bar.get_y() + bar.get_height()/2,
                f'{pct:.1f}%  ({qty:,} шт)'.replace(',', ' '),
                va='center', ha='left', fontsize=9, color=DARK)

    ax.set_title(f'{cat_name}\nДоля цвета в продажах (по штукам)', fontsize=14, fontweight='bold', color=DARK, pad=16)
    ax.spines['left'].set_color('#E5E7EB')
    ax.spines['bottom'].set_color('#E5E7EB')

    bar_file = os.path.join(CHART_DIR, f'bar_{cat_name[:10]}.png')
    fig.savefig(bar_file, dpi=200, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    print(f'  Saved: {bar_file}')

    # --- Pie chart (топ-5 + остальные) ---
    top5 = items_sorted[:5]
    other_pct = sum(i['pct_qty'] for i in items_sorted[5:])
    other_qty = sum(i['qty'] for i in items_sorted[5:])

    pie_labels = [i['color'] for i in top5]
    pie_pcts = [i['pct_qty'] for i in top5]
    if other_pct > 0:
        pie_labels.append(f'Остальные ({len(items_sorted)-5})')
        pie_pcts.append(other_pct)

    pie_colors_pie = [ORANGE_DARK, ORANGE, ORANGE_LIGHT, '#F5C896', GREY, '#E5E7EB'][:len(pie_labels)]

    fig2, ax2 = plt.subplots(figsize=(8, 8), constrained_layout=True)
    wedges, texts, autotexts = ax2.pie(
        pie_pcts, labels=None, autopct='%1.1f%%',
        colors=pie_colors_pie, startangle=90,
        pctdistance=0.75, wedgeprops={'edgecolor': 'white', 'linewidth': 2},
        textprops={'fontsize': 11, 'color': WHITE, 'fontweight': 'bold'}
    )
    # Легенда справа
    legend_labels = [f'{l} — {p:.1f}%' for l, p in zip(pie_labels, pie_pcts)]
    ax2.legend(wedges, legend_labels, loc='center left', bbox_to_anchor=(1.0, 0.5), fontsize=10)
    ax2.set_title(f'{cat_name}\nТоп-5 цветов по доле продаж', fontsize=14, fontweight='bold', color=DARK, pad=16)

    pie_file = os.path.join(CHART_DIR, f'pie_{cat_name[:10]}.png')
    fig2.savefig(pie_file, dpi=200, bbox_inches='tight', facecolor='white')
    plt.close(fig2)
    print(f'  Saved: {pie_file}')

    chart_files[cat_name] = {'bar': bar_file, 'pie': pie_file}

print('\nВсе диаграммы созданы')

# ============================================================
# WORD ДОКУМЕНТ
# ============================================================
ORANGE_W = RGBColor(0xF4, 0x76, 0x29)
DARK_W   = RGBColor(0x31, 0x31, 0x31)
GREY_W   = RGBColor(0x6B, 0x72, 0x80)
WHITE_W  = RGBColor(0xFF, 0xFF, 0xFF)
RED_W    = RGBColor(0xC0, 0x39, 0x2B)
GREEN_W  = RGBColor(0x16, 0xA3, 0x4A)

doc = Document()
style = doc.styles['Normal']
style.font.name = 'Arial'
style.font.size = Pt(11)
style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')

for section in doc.sections:
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)

def set_spacing(p, before=0, after=6, line=1.3):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line

def add_run(p, text, size=11, bold=False, color=None, italic=False):
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.bold = bold
    r.italic = italic
    if color:
        r.font.color.rgb = color
    r.font.name = 'Arial'
    return r

def add_para(text, size=11, bold=False, color=None, italic=False, align='left', after=6, before=0):
    p = doc.add_paragraph()
    set_spacing(p, before=before, after=after, line=1.3)
    p.alignment = {'center': WD_ALIGN_PARAGRAPH.CENTER, 'right': WD_ALIGN_PARAGRAPH.RIGHT,
                   'left': WD_ALIGN_PARAGRAPH.LEFT, 'justify': WD_ALIGN_PARAGRAPH.JUSTIFY}.get(align, WD_ALIGN_PARAGRAPH.LEFT)
    add_run(p, text, size=size, bold=bold, color=color, italic=italic)
    return p

def add_h1(text, color=DARK_W, size=20, before=0, after=8, orange_bar=True):
    p = doc.add_paragraph()
    set_spacing(p, before=before, after=after, line=1.1)
    add_run(p, text, size=size, bold=True, color=color)
    if orange_bar:
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        left = OxmlElement('w:left')
        left.set(qn('w:val'), 'single')
        left.set(qn('w:sz'), '24')
        left.set(qn('w:space'), '8')
        left.set(qn('w:color'), 'F47629')
        pBdr.append(left)
        pPr.append(pBdr)
        p.paragraph_format.left_indent = Cm(0.3)
    return p

def add_h2(text, color=ORANGE_W, size=14, before=10, after=4):
    p = doc.add_paragraph()
    set_spacing(p, before=before, after=after, line=1.2)
    add_run(p, text, size=size, bold=True, color=color)
    return p

def set_cell_bg(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def set_cell_borders(cell, color='D0D0D0', sz='4'):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), sz)
        el.set(qn('w:color'), color)
        tcBorders.append(el)
    tcPr.append(tcBorders)

def add_table(headers, rows, col_widths, font_size=10, header_bg='313131'):
    tbl = doc.add_table(rows=1 + len(rows), cols=len(headers))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    for i, w in enumerate(col_widths):
        for cell in tbl.columns[i].cells:
            cell.width = Cm(w)
    for i, h in enumerate(headers):
        c = tbl.rows[0].cells[i]
        set_cell_bg(c, header_bg)
        set_cell_borders(c, color='313131', sz='4')
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_spacing(p, before=2, after=2, line=1.2)
        add_run(p, h, size=font_size, bold=True, color=WHITE_W)
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            c = tbl.rows[ri + 1].cells[ci]
            if ri % 2 == 1:
                set_cell_bg(c, 'F8F4F0')
            set_cell_borders(c, color='D0D0D0', sz='4')
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = c.paragraphs[0]
            set_spacing(p, before=2, after=2, line=1.25)
            add_run(p, str(val), size=font_size)
            if ci == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.runs[0].bold = True
                p.runs[0].font.color.rgb = ORANGE_W
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc.add_paragraph()
    set_spacing(p, before=0, after=4, line=1.0)
    return tbl

WHITE_W = RGBColor(0xFF, 0xFF, 0xFF)

def add_alert_box(title, body, color='green'):
    color_map = {
        'red':    ('FEE2E2', 'C0392B'),
        'amber':  ('FEF3C7', 'B8860B'),
        'green':  ('D1FAE5', '166534'),
        'blue':   ('DBEAFE', '1E3A5F'),
    }
    bg, bar = color_map.get(color, color_map['green'])
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    cell = tbl.cell(0, 0)
    cell.width = Cm(17.0)
    set_cell_bg(cell, bg)
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge, c in (('top', bg), ('right', bg), ('bottom', bg), ('left', bar)):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), '36' if edge == 'left' else '4')
        el.set(qn('w:color'), c)
        tcBorders.append(el)
    tcPr.append(tcBorders)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    if title:
        p = cell.paragraphs[0]
        set_spacing(p, before=2, after=2, line=1.2)
        add_run(p, title, size=12, bold=True, color=RGBColor.from_string(bar))
    if body:
        p2 = cell.add_paragraph()
        set_spacing(p2, before=0, after=2, line=1.3)
        add_run(p2, body, size=11, color=DARK_W)
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for edge, val in (('top', '120'), ('left', '200'), ('bottom', '120'), ('right', '200')):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:w'), val)
        el.set(qn('w:type'), 'dxa')
        tcMar.append(el)
    tcPr.append(tcMar)
    p = doc.add_paragraph()
    set_spacing(p, before=0, after=4, line=1.0)
    return tbl

def page_break():
    p = doc.add_paragraph()
    r = p.add_run()
    r.add_break(WD_BREAK.PAGE)

def hr():
    p = doc.add_paragraph()
    set_spacing(p, before=2, after=6, line=1.0)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:color'), 'F47629')
    pBdr.append(bottom)
    pPr.append(pBdr)

def add_image(path, width_cm=16):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_spacing(p, before=6, after=6, line=1.0)
    r = p.add_run()
    r.add_picture(path, width=Cm(width_cm))

# ============================================================
# ЗАГОЛОВОК
# ============================================================
add_h1('Аналитика по цветам TITAN', size=22, before=0, after=2)
add_para('Вторичные продажи по цветам в 3 категориях продукции',
         size=14, bold=True, color=DARK_W, after=2)
add_para('Период: 15.07.2026  ·  Источник: вторичные продажи партнёров',
         size=11, italic=True, color=GREY_W, after=10)

hr()

add_para(
    'Аналитика показывает популярность каждого цвета RAL в трёх категориях '
    'продукции Ореол TITAN: грунт-эмаль 3в1, грунт-эмаль быстросохнущая и '
    'эмаль с молотковым эффектом. Данные основаны на вторичных продажах '
    '(от дистрибьюторов в розницу). Для каждого цвета рассчитан процент от '
    'общего объёма продаж в штуках и килограммах, а также количество активных '
    'клиентов (АКБ).',
    after=8
)

# Сводная таблица по 3 категориям
add_h2('Сводная статистика по категориям', before=4, after=4)

total_all_qty = sum(d['total_qty'] for d in categories.values())
total_all_kg = sum(d['total_kg'] for d in categories.values())
total_all_akb = sum(d['total_akb'] for d in categories.values())

add_table(
    headers=['Категория', 'Кол-во цветов', 'Продажи, шт', 'Продажи, кг', 'АКБ'],
    rows=[
        ['Грунт-эмаль 3в1',           str(len(categories['Грунт-эмаль']['items'])),
         f"{categories['Грунт-эмаль']['total_qty']:,}".replace(',', ' '),
         f"{categories['Грунт-эмаль']['total_kg']:.0f}",
         str(categories['Грунт-эмаль']['total_akb'])],
        ['Быстросохнущая',             str(len(categories['Грунт-эмаль быстросохнущая']['items'])),
         f"{categories['Грунт-эмаль быстросохнущая']['total_qty']:,}".replace(',', ' '),
         f"{categories['Грунт-эмаль быстросохнущая']['total_kg']:.0f}",
         str(categories['Грунт-эмаль быстросохнущая']['total_akb'])],
        ['Молотковая эмаль',           str(len(categories['Эмаль с молотковым эффектом']['items'])),
         f"{categories['Эмаль с молотковым эффектом']['total_qty']:,}".replace(',', ' '),
         f"{categories['Эмаль с молотковым эффектом']['total_kg']:.0f}",
         str(categories['Эмаль с молотковым эффектом']['total_akb'])],
        ['ИТОГО',                       '—',
         f"{total_all_qty:,}".replace(',', ' '),
         f"{total_all_kg:.0f}",
         str(total_all_akb)],
    ],
    col_widths=[5.0, 3.0, 3.5, 3.0, 2.5],
    font_size=10
)

add_alert_box('Главная находка',
    f'Топ-5 цветов дают ~70% продаж во всех 3 категориях. '
    f'Чёрный RAL 9005 — безусловный лидер во всех категориях: '
    f'24,3% в грунт-эмали, 41,6% в быстросохнущей, 30,5% в молотковой. '
    f'Это подтверждает правило выкладки: топ-5 цветов — приоритет на полке.',
    color='green')

page_break()

# ============================================================
# КАТЕГОРИЯ 1 — ГРУНТ-ЭМАЛЬ 3в1
# ============================================================
cat1 = categories['Грунт-эмаль']
items1 = sorted(cat1['items'], key=lambda x: x['qty'], reverse=True)

add_h1('Категория 1. Грунт-эмаль 3в1', size=18, before=0, after=4)
add_para(f'Всего: {cat1["total_qty"]:,} шт  ·  {cat1["total_kg"]:.0f} кг  ·  {cat1["total_akb"]} АКБ  ·  {len(items1)} цветов'.replace(',', ' '),
         size=12, bold=True, color=ORANGE_W, after=8)

# Таблица процентов
add_h2('Таблица популярности цветов', before=4, after=4)

add_table(
    headers=['№', 'Цвет', 'Продажи, шт', '% шт', 'Продажи, кг', '% кг', 'АКБ'],
    rows=[
        [str(i+1), item['color'],
         f"{item['qty']:,}".replace(',', ' '),
         f"{item['pct_qty']:.1f}%",
         f"{item['kg']:.0f}",
         f"{item['pct_kg']:.1f}%",
         str(item['akb'])]
        for i, item in enumerate(items1)
    ],
    col_widths=[0.8, 5.0, 2.5, 1.5, 2.5, 1.5, 1.5],
    font_size=9
)

# Диаграмма bar
add_h2('Диаграмма: доля цвета в продажах', before=10, after=4)
add_image(chart_files['Грунт-эмаль']['bar'], width_cm=16)

page_break()

# Диаграмма pie
add_h2('Топ-5 цветов', before=0, after=4)
add_image(chart_files['Грунт-эмаль']['pie'], width_cm=14)

add_alert_box('Вывод по грунт-эмали 3в1',
    'Топ-5 цветов (чёрный, графитовый, серый, шоколад, белый) дают 80,5% продаж в штуках. '
    'Чёрный RAL 9005 — лидер с 24,3%. Цвета «серый» и «красно-коричневый» (без RAL) '
    'нуждаются в стандартизации под RAL для упрощения выкладки и учёта.',
    color='blue')

page_break()

# ============================================================
# КАТЕГОРИЯ 2 — БЫСТРОСОХНУЩАЯ
# ============================================================
cat2 = categories['Грунт-эмаль быстросохнущая']
items2 = sorted(cat2['items'], key=lambda x: x['qty'], reverse=True)

add_h1('Категория 2. Грунт-эмаль быстросохнущая', size=18, before=0, after=4)
add_para(f'Всего: {cat2["total_qty"]:,} шт  ·  {cat2["total_kg"]:.0f} кг  ·  {cat2["total_akb"]} АКБ  ·  {len(items2)} цветов'.replace(',', ' '),
         size=12, bold=True, color=ORANGE_W, after=8)

add_h2('Таблица популярности цветов', before=4, after=4)

add_table(
    headers=['№', 'Цвет', 'Продажи, шт', '% шт', 'Продажи, кг', '% кг', 'АКБ'],
    rows=[
        [str(i+1), item['color'],
         f"{item['qty']:,}".replace(',', ' '),
         f"{item['pct_qty']:.1f}%",
         f"{item['kg']:.0f}",
         f"{item['pct_kg']:.1f}%",
         str(item['akb'])]
        for i, item in enumerate(items2)
    ],
    col_widths=[0.8, 5.0, 2.5, 1.5, 2.5, 1.5, 1.5],
    font_size=9
)

add_h2('Диаграмма: доля цвета в продажах', before=10, after=4)
add_image(chart_files['Грунт-эмаль быстросохнущая']['bar'], width_cm=16)

page_break()

add_h2('Топ-5 цветов', before=0, after=4)
add_image(chart_files['Грунт-эмаль быстросохнущая']['pie'], width_cm=14)

add_alert_box('Вывод по быстросохнущей',
    'Чёрный RAL 9005 доминирует с 41,6% — почти половина всех продаж. '
    'Топ-3 (чёрный, серый, шоколадно-коричневый) дают 74,3%. '
    'Категория более концентрированная, чем грунт-эмаль 3в1 — меньше цветов, '
    'но чёткий лидер. Два варианта «шоколадно-коричневый RAL 8017» нужно объединить '
    'в один SKU для корректной аналитики.',
    color='blue')

page_break()

# ============================================================
# КАТЕГОРИЯ 3 — МОЛОТКОВАЯ ЭМАЛЬ
# ============================================================
cat3 = categories['Эмаль с молотковым эффектом']
items3 = sorted(cat3['items'], key=lambda x: x['qty'], reverse=True)

add_h1('Категория 3. Эмаль с молотковым эффектом', size=18, before=0, after=4)
add_para(f'Всего: {cat3["total_qty"]:,} шт  ·  {cat3["total_kg"]:.0f} кг  ·  {cat3["total_akb"]} АКБ  ·  {len(items3)} цветов'.replace(',', ' '),
         size=12, bold=True, color=ORANGE_W, after=8)

add_h2('Таблица популярности цветов', before=4, after=4)

add_table(
    headers=['№', 'Цвет', 'Продажи, шт', '% шт', 'Продажи, кг', '% кг', 'АКБ'],
    rows=[
        [str(i+1), item['color'],
         f"{item['qty']:,}".replace(',', ' '),
         f"{item['pct_qty']:.1f}%",
         f"{item['kg']:.0f}",
         f"{item['pct_kg']:.1f}%",
         str(item['akb'])]
        for i, item in enumerate(items3)
    ],
    col_widths=[0.8, 5.0, 2.5, 1.5, 2.5, 1.5, 1.5],
    font_size=9
)

add_h2('Диаграмма: доля цвета в продажах', before=10, after=4)
add_image(chart_files['Эмаль с молотковым эффектом']['bar'], width_cm=16)

page_break()

add_h2('Топ-5 цветов (все 6 цветов)', before=0, after=4)
add_image(chart_files['Эмаль с молотковым эффектом']['pie'], width_cm=14)

add_alert_box('Вывод по молотковой эмали',
    'Чёрный — лидер с 30,5%, но распределение более равномерное, чем в других '
    'категориях. Топ-3 (чёрный, коричневый, серебристый) дают 76,6%. '
    'Все 6 цветов активно продаются — нет «мёртвых» цветов. '
    'Рекомендация: сохранять полную матрицу 6 цветов во всех точках.',
    color='blue')

page_break()

# ============================================================
# РЕКОМЕНДАЦИИ
# ============================================================
add_h1('Рекомендации по выкладке и ассортименту', size=18, before=0, after=6)

add_h2('1. Правило топ-5 цветов', before=4, after=4)
add_para(
    'Во всех 3 категориях топ-5 цветов дают 70–80% продаж. Это подтверждает '
    'правило выкладки из Sales Guide TITAN: топ-5 цветов — приоритет на полке, '
    'остальные — по запросу. Для точек класса C и D достаточно только топ-5.',
    after=8
)

add_h2('2. Чёрный RAL 9005 — безусловный лидер', before=8, after=4)
add_para(
    'Чёрный — №1 во всех 3 категориях: 24,3% (грунт-эмаль), 41,6% (быстросохнущая), '
    '30,5% (молотковая). Рекомендация: 2 фейса на чёрный во всех точках, '
    'глубина 6 шт. на фейс (класс A/B), 4 шт. (класс C), 2 шт. (класс D).',
    after=8
)

add_h2('3. Стандартизация номенклатуры', before=8, after=4)
add_para(
    'В грунт-эмали есть дублирующие цвета: «серый» и «серый RAL 7004», '
    '«красно-коричневый» и «красно-коричневый RAL 3009». В быстросохнущей — '
    'два варианта «шоколадно-коричневый RAL 8017». Это затрудняет аналитику и '
    'управление запасами. Рекомендация: стандартизировать все цвета под RAL.',
    after=8
)

add_h2('4. Расширение палитры — потенциал', before=8, after=4)
add_para(
    'Цвета с долей < 5% (жёлтый, красно-коричневый, синий, серебро, слоновая кость) '
    'дают суммарно 10–15% продаж. Это нишевые цвета — не убирать, но и не расширять '
    'выкладку. Достаточно 1 фейс на каждый.',
    after=8
)

# ---------- Финал ----------
hr()
add_para('TITAN  ·  Аналитика по цветам  ·  15.07.2026',
         size=10, italic=True, color=GREY_W, align='center', after=2)
add_para('Источник: вторичные продажи партнёров. Расчёт процентов — по штукам и килограммам.',
         size=10, italic=True, color=GREY_W, align='center')

# ---------- Сохранение ----------
out = '/home/z/my-project/download/TITAN_Аналитика_по_цветам.docx'
doc.save(out)
print(f'\nOK saved: {out}')
print(f'Size: {os.path.getsize(out)} bytes')
