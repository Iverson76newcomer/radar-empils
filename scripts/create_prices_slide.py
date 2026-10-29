#!/usr/bin/env python3
"""
PPT-слайд: Ценовое сравнение TITAN vs конкуренты — было/стало.
Источник: TITAN_конкуренты_цены_ВсеИнструменты.xlsx
"""

import openpyxl
from pptx import Presentation
from pptx.util import Inches, Pt, Cm, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import os

# Регистрация шрифтов для matplotlib (с поддержкой кириллицы)
fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf')
fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf')
fm.fontManager.addfont('/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf')
fm.fontManager.addfont('/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf')
plt.rcParams['font.sans-serif'] = ['Liberation Sans', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

EXCEL_PATH = '/home/z/my-project/download/TITAN_конкуренты_цены_ВсеИнструменты.xlsx'
CHART_PATH = '/home/z/my-project/scripts/prices_chart.png'
OUT = '/home/z/my-project/download/TITAN_ценовое_сравнение.pptx'

# Цвета
COL_PRIMARY = RGBColor(0x1E, 0x3A, 0x5F)
COL_ACCENT = RGBColor(0xF4, 0x79, 0x20)
COL_RED = RGBColor(0xC6, 0x28, 0x28)
COL_GREEN = RGBColor(0x16, 0xA3, 0x4A)
COL_BODY = RGBColor(0x1A, 0x1F, 0x36)
COL_GRAY = RGBColor(0x6B, 0x72, 0x80)
COL_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COL_TITAN = RGBColor(0xF4, 0x79, 0x20)  # оранжевый для TITAN


# === 1. ЧТЕНИЕ ДАННЫХ ИЗ EXCEL ===
wb = openpyxl.load_workbook(EXCEL_PATH, data_only=False)
ws = wb['Цены было-стало']

products = []
for ri in range(5, 21):
    n = ws.cell(ri, 1).value
    cat = ws.cell(ri, 2).value
    brand = ws.cell(ri, 3).value
    product = ws.cell(ri, 4).value
    vol = ws.cell(ri, 5).value
    price_old = ws.cell(ri, 7).value
    price_new = ws.cell(ri, 8).value
    is_titan = brand == 'TITAN'
    if price_new:
        delta = price_new - price_old
        delta_pct = delta / price_old * 100
    else:
        delta = None
        delta_pct = None
    products.append({
        'n': n, 'cat': cat, 'brand': brand, 'product': product,
        'vol': vol, 'price_old': price_old, 'price_new': price_new,
        'delta': delta, 'delta_pct': delta_pct, 'is_titan': is_titan,
    })

# Разделение на категории
cats_3v1 = [p for p in products if '3в1' in p['cat']]
cats_fast = [p for p in products if 'Быстро' in p['cat']]
cats_hammer = [p for p in products if 'Молот' in p['cat']]

# === 2. ГЕНЕРАЦИЯ ДИАГРАММЫ ===
# Сортируем все 16 по delta_pct (от падения к росту)
# Сначала те, у кого нет цены — в самый низ
sorted_products = sorted(
    [p for p in products],
    key=lambda x: (x['delta_pct'] is None, x['delta_pct'] if x['delta_pct'] is not None else -999)
)

fig, ax = plt.subplots(figsize=(11, 6.5), constrained_layout=True)

# Данные
labels = []
values = []
colors_bar = []
for p in sorted_products:
    # Короткая подпись
    cat_short = '3в1' if '3в1' in p['cat'] else ('Быстр.' if 'Быстро' in p['cat'] else 'Молот.')
    label = f"{p['brand']} ({cat_short})"
    labels.append(label)
    if p['delta_pct'] is not None:
        values.append(p['delta_pct'])
    else:
        values.append(0)  # для отсутствующей цены
    
    # Цвета
    if p['is_titan']:
        colors_bar.append('#F47920')  # оранжевый — TITAN
    elif p['delta_pct'] is None:
        colors_bar.append('#9CA3AF')  # серый — нет данных
    elif p['delta_pct'] < 0:
        colors_bar.append('#16A34A')  # зелёный — упали (хорошо для покупателя, плохо для бренда)
    else:
        colors_bar.append('#C62828')  # красный — выросли

# Рисуем горизонтальную диаграмму
y_pos = range(len(labels))
bars = ax.barh(y_pos, values, color=colors_bar, edgecolor='white', linewidth=0.8)

# Подписи на столбцах
for i, (bar, p) in enumerate(zip(bars, sorted_products)):
    width = bar.get_width()
    if p['delta_pct'] is not None:
        text = f"{p['delta_pct']:+.0f}%  ({p['price_old']}->{p['price_new']} руб)"
    else:
        text = "нет данных"
    
    # Размещаем подпись справа от столбца (или слева, если ушёл в минус)
    if width >= 0:
        ax.text(width + 1.5, bar.get_y() + bar.get_height()/2, text,
                ha='left', va='center', fontsize=8.5, color='#1A1F36',
                fontweight='bold' if p['is_titan'] else 'normal')
    else:
        ax.text(width - 1.5, bar.get_y() + bar.get_height()/2, text,
                ha='right', va='center', fontsize=8.5, color='#1A1F36',
                fontweight='bold' if p['is_titan'] else 'normal')

# Настройка осей
ax.set_yticks(y_pos)
ax.set_yticklabels(labels, fontsize=9)
# Жирным TITAN
for tick_label, p in zip(ax.get_yticklabels(), sorted_products):
    if p['is_titan']:
        tick_label.set_fontweight('bold')
        tick_label.set_color('#F47920')

ax.axvline(0, color='#374151', linewidth=1.2)
ax.set_xlim(-35, 70)
ax.set_xlabel('Изменение цены, %', fontsize=10, color='#1A1F36')
ax.set_title('Δ цены TITAN и конкурентов (ВсеИнструменты.ру, 10.06.2026 → актуально)',
             fontsize=12, fontweight='bold', color='#1E3A5F', pad=12)

# Сетка
ax.grid(axis='x', linestyle='--', alpha=0.3)
ax.set_axisbelow(True)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['left'].set_color('#9CA3AF')
ax.spines['bottom'].set_color('#9CA3AF')

# Легенда
from matplotlib.patches import Patch
legend_elements = [
    Patch(facecolor='#F47920', label='TITAN'),
    Patch(facecolor='#C62828', label='Подорожал'),
    Patch(facecolor='#16A34A', label='Подешевел'),
    Patch(facecolor='#9CA3AF', label='Нет данных'),
]
ax.legend(handles=legend_elements, loc='lower right', fontsize=9, frameon=False)

plt.savefig(CHART_PATH, dpi=180, facecolor='white')
plt.close()
print(f'OK: график сохранён {CHART_PATH}')


# === 3. СОЗДАНИЕ PPT-СЛАЙДА ===
prs = Presentation()
prs.slide_width = Cm(33.867)  # 16:9
prs.slide_height = Cm(19.05)

slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

# === Заголовок ===
title_box = slide.shapes.add_textbox(Cm(1.0), Cm(0.5), Cm(31), Cm(1.0))
tf = title_box.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.LEFT
r = p.add_run()
r.text = 'TITAN vs конкуренты — ценовое сравнение'
r.font.name = 'Arial'
r.font.size = Pt(24)
r.font.bold = True
r.font.color.rgb = COL_PRIMARY

# Подзаголовок
sub_box = slide.shapes.add_textbox(Cm(1.0), Cm(1.6), Cm(31), Cm(0.6))
tf = sub_box.text_frame
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.LEFT
r = p.add_run()
r.text = 'Источник: ВсеИнструменты.ру. Цены 10.06.2026 (из презентации TITAN) vs актуальные на сегодня. 16 SKU в 3 категориях.'
r.font.name = 'Arial'
r.font.size = Pt(11)
r.font.color.rgb = COL_GRAY
r.font.italic = True

# === ГРАФИК (слева, занимает большую часть слайда) ===
slide.shapes.add_picture(CHART_PATH, Cm(0.3), Cm(2.5), width=Cm(22.5), height=Cm(13.5))

# === 3 ИТОГОВЫХ БЛОКА (справа) ===
# Блок 1: TITAN итог
titan_box = slide.shapes.add_textbox(Cm(23.5), Cm(2.5), Cm(9.5), Cm(3.0))
tf = titan_box.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.LEFT
r = p.add_run()
r.text = 'TITAN (3 SKU)'
r.font.name = 'Arial'
r.font.size = Pt(13)
r.font.bold = True
r.font.color.rgb = COL_ACCENT

p2 = tf.add_paragraph()
p2.alignment = PP_ALIGN.LEFT
r = p2.add_run()
r.text = 'Средняя цена: 982 → 1 052 ₽'
r.font.name = 'Arial'
r.font.size = Pt(11)
r.font.color.rgb = COL_BODY
r2 = p2.add_run()
r2.text = '  (+7%)'
r2.font.name = 'Arial'
r2.font.size = Pt(11)
r2.font.bold = True
r2.font.color.rgb = COL_RED

for line in ['• 3в1: 872 → 905 (+4%)', '• Быстр.: 765 → 956 (+25%) ↑↑', '• Молот.: 1 308 → 1 294 (−1%)']:
    p3 = tf.add_paragraph()
    p3.alignment = PP_ALIGN.LEFT
    r = p3.add_run()
    r.text = line
    r.font.name = 'Arial'
    r.font.size = Pt(10)
    r.font.color.rgb = COL_BODY

# Блок 2: Кто демпингует
dump_box = slide.shapes.add_textbox(Cm(23.5), Cm(6.0), Cm(9.5), Cm(4.5))
tf = dump_box.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
r = p.add_run()
r.text = '⚠️ Конкуренты демпингуют'
r.font.name = 'Arial'
r.font.size = Pt(13)
r.font.bold = True
r.font.color.rgb = COL_GREEN

for line in [
    '• HUSKY 3в1: −17% (1 153→953)',
    '• Tikkurila METALLISTA: −23%',
    '  (2 085→1 599)',
    '• Dufa HAMMERLACK: −27%',
    '  (1 653→1 210)',
    '• Лакра молотковая: −15%',
]:
    p2 = tf.add_paragraph()
    r = p2.add_run()
    r.text = line
    r.font.name = 'Arial'
    r.font.size = Pt(10)
    r.font.color.rgb = COL_BODY

# Блок 3: Кто подорожал
grow_box = slide.shapes.add_textbox(Cm(23.5), Cm(11.0), Cm(9.5), Cm(4.5))
tf = grow_box.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
r = p.add_run()
r.text = '📈 Подорожали вместе с TITAN'
r.font.name = 'Arial'
r.font.size = Pt(13)
r.font.bold = True
r.font.color.rgb = COL_RED

for line in [
    '• Новбытхим: +53% (765→1 168)',
    '• HUSKY молотковая: +44%',
    '  (1 196→1 719)',
    '• Elcon молотковая: +23%',
    '  (1 097→1 344)',
    '• Dali: +8-9%',
]:
    p2 = tf.add_paragraph()
    r = p2.add_run()
    r.text = line
    r.font.name = 'Arial'
    r.font.size = Pt(10)
    r.font.color.rgb = COL_BODY

# === Главный вывод (внизу слайда) ===
concl_box = slide.shapes.add_textbox(Cm(1.0), Cm(16.5), Cm(31.5), Cm(1.8))
tf = concl_box.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.LEFT
r = p.add_run()
r.text = 'ВЫВОД: '
r.font.name = 'Arial'
r.font.size = Pt(13)
r.font.bold = True
r.font.color.rgb = COL_ACCENT
r2 = p.add_run()
r2.text = ('TITAN подорожал на +7% (быстрее рынка +3%) — теряет ценовое преимущество. '
           'HUSKY, Tikkurila и Dufa снизили цены на 15-27% — теперь дешевле TITAN в своих категориях. '
           'Рекомендация: точечные акции против HUSKY и Tikkurila, удержание цены на 3в1 (рост всего +4%).')
r2.font.name = 'Arial'
r2.font.size = Pt(11)
r2.font.color.rgb = COL_BODY

# === Подвал ===
footer_box = slide.shapes.add_textbox(Cm(1.0), Cm(18.3), Cm(31), Cm(0.5))
tf = footer_box.text_frame
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.LEFT
r = p.add_run()
r.text = '15 из 16 SKU найдены на vseinstrumenti.ru. SKU 10 (Лакра PROF IT быстросохнущая) — нет в наличии или цена не отображается.'
r.font.name = 'Arial'
r.font.size = Pt(8)
r.font.color.rgb = COL_GRAY
r.font.italic = True

prs.save(OUT)
print(f'OK: {OUT}')
print(f'Размер: {os.path.getsize(OUT) / 1024:.1f} KB')
