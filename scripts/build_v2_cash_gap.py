"""
Обновлённая модель кассового разрыва:
1. Лист "Объяснение простыми словами" — аналогии и термины
2. Улучшенные визуализации: каскад (waterfall), накопительный график, схема "бассейна"
"""
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, BarChart, Reference, AreaChart, BarChart3D
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.chart.marker import DataPoint as ChartDataPoint
from openpyxl.drawing.fill import PatternFillProperties, ColorChoice
import json
import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import numpy as np

# === Загрузить модель ===
with open('/home/z/my-project/scripts/cash_gap_model.json', 'r', encoding='utf-8') as f:
    model = json.load(f)

cash_flow = model['cash_flow']
months = [cf['month'] for cf in cash_flow]
primary = [cf['primary'] for cf in cash_flow]
secondary = [cf['secondary'] for cf in cash_flow]
pay_empils = [cf['pay_empils'] for cf in cash_flow]
receive_rtt = [cf['receive_rtt'] for cf in cash_flow]
net = [cf['net'] for cf in cash_flow]
cumulative = [cf['cumulative'] for cf in cash_flow]

# === Matplotlib setup for Cyrillic ===
fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf')
fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf')
fm.fontManager.addfont('/usr/share/fonts/truetype/lxgw-wenkai/LXGWWenKai-Regular.ttf')
plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'LXGW WenKai']
plt.rcParams['axes.unicode_minus'] = False

OUT_PATH = "/home/z/my-project/download/Кассовый_разрыв_Олимп_модель_v2.xlsx"
IMG_DIR = "/home/z/my-project/scripts/cash_gap_images"
os.makedirs(IMG_DIR, exist_ok=True)

# === Стили ===
HEAD_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
HEAD_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
HEAD_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
TITLE_FONT = Font(name="Calibri", size=14, bold=True, color="1F4E79")
SUBTITLE_FONT = Font(name="Calibri", size=10, italic=True, color="595959")
BODY_FONT = Font(name="Calibri", size=10)
BODY_ALIGN = Alignment(horizontal="left", vertical="center", wrap_text=True)
NUM_ALIGN = Alignment(horizontal="right", vertical="center")
CENTER_ALIGN = Alignment(horizontal="center", vertical="center")
GAP_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
WARN_FILL = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
POSITIVE_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
POSITIVE_FONT = Font(name="Calibri", size=10, bold=True, color="375623")
NEGATIVE_FONT = Font(name="Calibri", size=10, bold=True, color="C00000")
TOTAL_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
PRIMARY_HIGHLIGHT = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
SECONDARY_HIGHLIGHT = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
INCOME_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
OUTCOME_FILL = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
THIN = Side(border_style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# ====================== ВИЗУАЛИЗАЦИЯ 1: Каскад (waterfall) ======================
# Показывает, как из стартового 0 формируется накопленный баланс
fig, ax = plt.subplots(figsize=(14, 7), constrained_layout=True)

# Каскад: каждый месяц показывает изменение (net) и накопленный итог
x_pos = np.arange(len(months))
colors = ['#70AD47' if n > 0 else ('#C00000' if n < 0 else '#A6A6A6') for n in net]

# Строим waterfall: каждый следующий столбец начинается с предыдущего накопленного
bottoms = []
heights = []
running_total = 0
for i, n in enumerate(net):
    if n >= 0:
        bottoms.append(running_total)
        heights.append(n)
    else:
        bottoms.append(running_total + n)  # n отрицательное, поэтому + n = ниже
        heights.append(-n)  # высота положительная
    running_total += n

bars = ax.bar(x_pos, heights, bottom=bottoms, color=colors, edgecolor='black', linewidth=0.5, width=0.7)

# Линия накопленного баланса сверху
ax.plot(x_pos, cumulative, color='#1F4E79', marker='o', linewidth=2.5, markersize=8, label='Накопленный баланс', zorder=5)

# Подписи значений
for i, (h, b, n, c) in enumerate(zip(heights, bottoms, net, colors)):
    if n != 0:
        # Подпись изменения
        y_label = b + h if n > 0 else b
        ax.text(i, y_label + (200000 if n > 0 else -200000), 
                f'{n/1e6:+.1f} млн', 
                ha='center', va='bottom' if n > 0 else 'top',
                fontsize=9, color=c, fontweight='bold')

# Подпись накопленного баланса
for i, c in enumerate(cumulative):
    ax.text(i, c + (400000 if c >= 0 else -400000), 
            f'{c/1e6:+.1f} млн',
            ha='center', va='bottom' if c >= 0 else 'top',
            fontsize=9, color='#1F4E79', fontweight='bold')

# Нулевая линия
ax.axhline(y=0, color='black', linewidth=1.5, linestyle='-', alpha=0.7)

# Зона кассового разрыва (отрицательные значения)
ax.fill_between(x_pos, 0, min(cumulative) - 500000, where=[c < 0 for c in cumulative],
                color='#FFC7CE', alpha=0.3, label='Зона кассового разрыва')

# Выделить месяц с максимальным разрывом
min_idx = cumulative.index(min(cumulative))
ax.annotate(f'Минимум: {cumulative[min_idx]/1e6:+.1f} млн\n({months[min_idx]})',
            xy=(min_idx, cumulative[min_idx]),
            xytext=(min_idx + 1.5, cumulative[min_idx] - 1_500_000),
            arrowprops=dict(arrowstyle='->', color='red', lw=2),
            fontsize=11, color='red', fontweight='bold',
            bbox=dict(boxstyle='round,pad=0.5', facecolor='yellow', alpha=0.8))

ax.set_xticks(x_pos)
ax.set_xticklabels(months, rotation=0, fontsize=10)
ax.set_ylabel('Рубли', fontsize=11)
ax.set_title('Каскадная диаграмма: как формируется накопленный баланс\n(зелёные столбцы — приток, красные — отток)', 
             fontsize=13, fontweight='bold', color='#1F4E79')
ax.grid(True, alpha=0.3, axis='y')
ax.legend(loc='upper left', fontsize=10)

# Формат оси Y
ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1e6:.0f} млн'))

# Сохранить
img1_path = os.path.join(IMG_DIR, 'waterfall.png')
plt.savefig(img1_path, dpi=150, bbox_inches='tight')
plt.close()
print(f"✓ График 1 (waterfall): {img1_path}")

# ====================== ВИЗУАЛИЗАЦИЯ 2: Накопленный баланс с зонами ======================
fig, ax = plt.subplots(figsize=(14, 6), constrained_layout=True)

# Заполняем положительную зону зелёным, отрицательную — красным
ax.fill_between(x_pos, cumulative, 0, 
                where=[c >= 0 for c in cumulative],
                color='#70AD47', alpha=0.4, label='Резерв (положительный баланс)')
ax.fill_between(x_pos, cumulative, 0, 
                where=[c < 0 for c in cumulative],
                color='#C00000', alpha=0.4, label='Кассовый разрыв (отрицательный баланс)')

# Линия накопленного баланса
ax.plot(x_pos, cumulative, color='#1F4E79', marker='o', linewidth=3, markersize=10, 
        label='Накопленный баланс', zorder=5)

# Подписи точек
for i, c in enumerate(cumulative):
    y_offset = 600000 if c >= 0 else -800000
    ax.text(i, c + y_offset, f'{c/1e6:+.1f} млн',
            ha='center', va='bottom' if c >= 0 else 'top',
            fontsize=10, color='#1F4E79', fontweight='bold')

# Нулевая линия
ax.axhline(y=0, color='black', linewidth=2, linestyle='-', alpha=0.8)

# Аннотации ключевых точек
# Точка разрыва
first_neg_idx = next(i for i, c in enumerate(cumulative) if c < 0)
ax.annotate(f'НАЧАЛО РАЗРЫВА\n{months[first_neg_idx]}: {cumulative[first_neg_idx]/1e6:+.1f} млн',
            xy=(first_neg_idx, cumulative[first_neg_idx]),
            xytext=(first_neg_idx - 2, cumulative[first_neg_idx] - 2_500_000),
            arrowprops=dict(arrowstyle='->', color='red', lw=2.5),
            fontsize=11, color='red', fontweight='bold',
            bbox=dict(boxstyle='round,pad=0.5', facecolor='#FFC7CE', edgecolor='red'))

# Точка восстановления
first_pos_after_neg = next((i for i, c in enumerate(cumulative) if c >= 0 and i > first_neg_idx), None)
if first_pos_after_neg:
    ax.annotate(f'ВОССТАНОВЛЕНИЕ\n{months[first_pos_after_neg]}: {cumulative[first_pos_after_neg]/1e6:+.1f} млн',
                xy=(first_pos_after_neg, cumulative[first_pos_after_neg]),
                xytext=(first_pos_after_neg + 0.5, cumulative[first_pos_after_neg] + 2_500_000),
                arrowprops=dict(arrowstyle='->', color='green', lw=2.5),
                fontsize=11, color='green', fontweight='bold',
                bbox=dict(boxstyle='round,pad=0.5', facecolor='#C6EFCE', edgecolor='green'))

ax.set_xticks(x_pos)
ax.set_xticklabels(months, rotation=0, fontsize=10)
ax.set_ylabel('Рубли', fontsize=11)
ax.set_title('Накопленный баланс по месяцам: зоны резерва и кассового разрыва', 
             fontsize=13, fontweight='bold', color='#1F4E79')
ax.grid(True, alpha=0.3, axis='y')
ax.legend(loc='lower right', fontsize=10)
ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1e6:.0f} млн'))

img2_path = os.path.join(IMG_DIR, 'cumulative_areas.png')
plt.savefig(img2_path, dpi=150, bbox_inches='tight')
plt.close()
print(f"✓ График 2 (cumulative areas): {img2_path}")

# ====================== ВИЗУАЛИЗАЦИЯ 3: Притоки vs оттоки (группированные столбцы) ======================
fig, ax = plt.subplots(figsize=(14, 6), constrained_layout=True)

width = 0.35
x = np.arange(len(months))

# Притоки (деньги от РТТ)
bars1 = ax.bar(x - width/2, receive_rtt, width, label='Деньги от РТТ (входящий)', 
               color='#70AD47', edgecolor='black', linewidth=0.5)

# Оттоки (платежи Эмпилс)
bars2 = ax.bar(x + width/2, pay_empils, width, label='Платежи Эмпилс (исходящий)', 
               color='#C00000', edgecolor='black', linewidth=0.5)

# Подписи значений
for i, (rec, pay) in enumerate(zip(receive_rtt, pay_empils)):
    if rec > 0:
        ax.text(i - width/2, rec + 200000, f'{rec/1e6:.1f} млн', 
                ha='center', va='bottom', fontsize=8, color='#375623', fontweight='bold')
    if pay > 0:
        ax.text(i + width/2, pay + 200000, f'{pay/1e6:.1f} млн', 
                ha='center', va='bottom', fontsize=8, color='#C00000', fontweight='bold')

ax.set_xticks(x)
ax.set_xticklabels(months, rotation=0, fontsize=10)
ax.set_ylabel('Рубли', fontsize=11)
ax.set_title('Притоки vs Оттоки по месяцам\n(если зелёный столбец выше красного — месяц в плюсе)', 
             fontsize=13, fontweight='bold', color='#1F4E79')
ax.grid(True, alpha=0.3, axis='y')
ax.legend(loc='upper right', fontsize=10)
ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1e6:.0f} млн'))

img3_path = os.path.join(IMG_DIR, 'inflows_outflows.png')
plt.savefig(img3_path, dpi=150, bbox_inches='tight')
plt.close()
print(f"✓ График 3 (inflows/outflows): {img3_path}")

# ====================== ВИЗУАЛИЗАЦИЯ 4: Схема "бассейна" ======================
fig, axes = plt.subplots(3, 4, figsize=(16, 9), constrained_layout=True)
axes = axes.flatten()

for i, (month_name, cum_val, net_val) in enumerate(zip(months, cumulative, net)):
    ax = axes[i]
    
    # "Бассейн" — прямоугольник, уровень воды = накопленный баланс
    # Если положительный — вода выше нулевой линии (синяя)
    # Если отрицательный — вода ниже (красная, "долг")
    
    max_val = max(cumulative) if max(cumulative) > 0 else 5_000_000
    min_val = min(cumulative) if min(cumulative) < 0 else -5_000_000
    scale = max(abs(max_val), abs(min_val))
    
    # Корзина
    ax.set_xlim(0, 1)
    ax.set_ylim(-1.5, 1.5)
    
    # Рамка бассейна
    ax.add_patch(plt.Rectangle((0.1, -1), 0.8, 2.5, fill=False, edgecolor='black', linewidth=2))
    
    # Нулевая линия
    ax.axhline(y=0, color='gray', linewidth=1, linestyle='--', alpha=0.5)
    ax.text(0.85, 0.05, '0', fontsize=8, color='gray', alpha=0.7)
    
    # Уровень воды
    if cum_val >= 0:
        # Заполнение сверху от нулевой линии
        height = cum_val / scale  # нормализованная высота (0..1)
        ax.add_patch(plt.Rectangle((0.1, 0), 0.8, height, 
                                    facecolor='#70AD47', alpha=0.7, edgecolor='#375623'))
        ax.text(0.5, height/2, f'+{cum_val/1e6:.1f} млн\n(резерв)', 
                ha='center', va='center', fontsize=9, color='white', fontweight='bold')
    else:
        # Заполнение снизу от нулевой линии (красное — "долг")
        height = abs(cum_val) / scale
        ax.add_patch(plt.Rectangle((0.1, -height), 0.8, height, 
                                    facecolor='#C00000', alpha=0.7, edgecolor='#8B0000'))
        ax.text(0.5, -height/2, f'{cum_val/1e6:+.1f} млн\n(разрыв)', 
                ha='center', va='center', fontsize=9, color='white', fontweight='bold')
    
    # Изменение за месяц (net)
    if net_val != 0:
        color = '#375623' if net_val > 0 else '#8B0000'
        sign = '+' if net_val > 0 else ''
        ax.text(0.5, -1.3, f'Δ за мес: {sign}{net_val/1e6:.1f} млн', 
                ha='center', va='center', fontsize=8, color=color, fontweight='bold')
    
    ax.set_title(f'{month_name}', fontsize=11, fontweight='bold', 
                 color='#C00000' if cum_val < 0 else '#1F4E79')
    ax.set_xticks([])
    ax.set_yticks([])

fig.suptitle('Схема "бассейна": зелёный = резерв, красный = кассовый разрыв\n(Δ — изменение за месяц: + приток, − отток)', 
             fontsize=13, fontweight='bold', color='#1F4E79')

img4_path = os.path.join(IMG_DIR, 'pool_diagram.png')
plt.savefig(img4_path, dpi=150, bbox_inches='tight')
plt.close()
print(f"✓ График 4 (pool diagram): {img4_path}")

# ====================== СОЗДАНИЕ EXCEL-ФАЙЛА ======================
wb = Workbook()
wb.properties.creator = "Z.ai"
wb.properties.title = "Модель кассового разрыва — Олимп v2 (с объяснениями и графиками)"

# ==================== ЛИСТ 1: ОБЪЯСНЕНИЕ ТЕРМИНОВ ====================
ws = wb.active
ws.title = "Объяснение простыми словами"

# Заголовок
ws.merge_cells("A1:F1")
ws["A1"] = "Что такое кассовый разрыв, чистый поток и накопленный баланс"
ws["A1"].font = TITLE_FONT
ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
ws.row_dimensions[1].height = 28

ws.merge_cells("A2:F2")
ws["A2"] = "Объяснение на бытовом примере, без бухгалтерских терминов"
ws["A2"].font = SUBTITLE_FONT
ws["A2"].alignment = Alignment(horizontal="left", vertical="center")
ws.row_dimensions[2].height = 18

# === Блок 1: Аналогия с кошельком ===
ws.merge_cells("A4:F4")
ws["A4"] = "АНАЛОГИЯ С КОШЕЛЬКОМ"
ws["A4"].font = Font(name="Calibri", size=12, bold=True, color="1F4E79")
ws["A4"].fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
ws.row_dimensions[4].height = 24

analogy_text = (
    "Представьте, что у вас есть кошелёк. Вы получаете зарплату и тратите деньги. "
    "В течение месяца в кошельке накапливается остаток — это и есть «накопленный баланс».\n\n"
    "Каждый день вы можете либо положить деньги в кошелёк (приток), либо достать (отток). "
    "Разница между тем, что положили и достали за день — это «чистый поток».\n\n"
    "Если в какой-то день нужно достать больше, чем есть в кошельке, — это кассовый разрыв. "
    "Вы вынуждены занимать деньги, иначе не сможете заплатить."
)

ws.merge_cells("A5:F9")
ws["A5"] = analogy_text
ws["A5"].font = Font(name="Calibri", size=11)
ws["A5"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
ws["A5"].fill = PatternFill(start_color="FFF8E1", end_color="FFF8E1", fill_type="solid")
ws.row_dimensions[5].height = 25
ws.row_dimensions[6].height = 25
ws.row_dimensions[7].height = 25
ws.row_dimensions[8].height = 25
ws.row_dimensions[9].height = 25

# === Блок 2: Чистый поток месяца ===
ws.merge_cells("A11:F11")
ws["A11"] = "ЧИСТЫЙ ПОТОК МЕСЯЦА"
ws["A11"].font = Font(name="Calibri", size=12, bold=True, color="1F4E79")
ws["A11"].fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
ws.row_dimensions[11].height = 24

# Пример таблицы
ws["A13"] = "Формула:"
ws["A13"].font = Font(name="Calibri", size=11, bold=True)
ws.merge_cells("B13:F13")
ws["B13"] = "Чистый поток = Деньги, пришедшие за месяц − Деньги, ушедшие за месяц"
ws["B13"].font = Font(name="Calibri", size=11, color="1F4E79", bold=True)
ws["B13"].alignment = BODY_ALIGN

ws["A14"] = "Пример (март):"
ws["A14"].font = Font(name="Calibri", size=11, bold=True)
ws.merge_cells("B14:F14")
ws["B14"] = "Чистый поток = 2,4 млн (пришло от РТТ) − 6,0 млн (платёж Эмпилс) = −3,6 млн руб"
ws["B14"].font = Font(name="Calibri", size=11, color="C00000", bold=True)
ws["B14"].alignment = BODY_ALIGN

# Подробное объяснение
ws.merge_cells("A16:F16")
ws["A16"] = "Что входит в чистый поток:"
ws["A16"].font = Font(name="Calibri", size=11, bold=True, color="1F4E79")

# Таблица компонентов чистого потока
components = [
    ("Компонент", "Что это", "Когда происходит", "В марте, млн руб"),
    ("Деньги от РТТ (входящий поток)", "Поступления от розничных торговых точек, которые заплатили за январскую-февральскую вторичку", "Через 30 дней после отгрузки в РТТ", "+2,4"),
    ("Платёж Эмпилс (исходящий поток)", "Оплата товара, который был закуплен в январе", "Через 45 дней после получения товара", "−6,0"),
    ("Зарплата, аренда, налоги", "Операционные расходы (в модели не учтены)", "Ежемесячно", "0 (упрощение)"),
    ("ЧИСТЫЙ ПОТОК МАРТА", "Пришло минус ушло", "Итог месяца", "−3,6 млн"),
]

for col_idx, (title, width) in enumerate([("Компонент", 32), ("Что это", 35), ("Когда происходит", 25), ("В марте", 15)], 1):
    cell = ws.cell(row=17, column=col_idx, value=title)
    cell.fill = HEAD_FILL
    cell.font = HEAD_FONT
    cell.alignment = HEAD_ALIGN
    cell.border = BORDER
    ws.column_dimensions[get_column_letter(col_idx)].width = width
ws.row_dimensions[17].height = 36

for i, (comp, what, when, val) in enumerate(components[1:], 18):
    ws.cell(row=i, column=1, value=comp).font = Font(name="Calibri", size=10, bold=True)
    ws.cell(row=i, column=1).border = BORDER
    ws.cell(row=i, column=1).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    
    ws.cell(row=i, column=2, value=what).font = BODY_FONT
    ws.cell(row=i, column=2).border = BORDER
    ws.cell(row=i, column=2).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    
    ws.cell(row=i, column=3, value=when).font = BODY_FONT
    ws.cell(row=i, column=3).border = BORDER
    ws.cell(row=i, column=3).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    
    val_cell = ws.cell(row=i, column=4, value=val)
    val_cell.border = BORDER
    val_cell.alignment = CENTER_ALIGN
    if '−' in val:
        val_cell.font = Font(name="Calibri", size=11, bold=True, color="C00000")
        val_cell.fill = GAP_FILL
    elif '+' in val:
        val_cell.font = Font(name="Calibri", size=11, bold=True, color="375623")
        val_cell.fill = POSITIVE_FILL
    else:
        val_cell.font = Font(name="Calibri", size=11, bold=True, color="1F4E79")
        val_cell.fill = TOTAL_FILL
    
    ws.row_dimensions[i].height = 36

# === Блок 3: Накопленный баланс ===
ws.merge_cells("A24:F24")
ws["A24"] = "НАКОПЛЕННЫЙ БАЛАНС"
ws["A24"].font = Font(name="Calibri", size=12, bold=True, color="1F4E79")
ws["A24"].fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
ws.row_dimensions[24].height = 24

ws["A26"] = "Формула:"
ws["A26"].font = Font(name="Calibri", size=11, bold=True)
ws.merge_cells("B26:F26")
ws["B26"] = "Накопленный баланс = Накопленный баланс прошлого месяца + Чистый поток текущего месяца"
ws["B26"].font = Font(name="Calibri", size=11, color="1F4E79", bold=True)
ws["B26"].alignment = BODY_ALIGN

ws["A27"] = "Простыми словами:"
ws["A27"].font = Font(name="Calibri", size=11, bold=True)
ws.merge_cells("B27:F30")
ws["B27"] = (
    "Накопленный баланс — это «сколько денег в кошельке на конец месяца».\n\n"
    "В январе стартуем с 0. В феврале добавили +2 млн → баланс = 2 млн.\n"
    "В марте «достали» 3,6 млн → баланс = 2 − 3,6 = −1,6 млн. ⚠ Кассовый разрыв!\n"
    "В апреле ещё «достали» 1,2 млн → баланс = −1,6 − 1,2 = −2,8 млн. Минимум!\n"
    "С мая начинаем выходить в плюс: +0,8 → −2,0; +0,8 → −1,2; +3,75 → +2,55 ✓"
)
ws["B27"].font = Font(name="Calibri", size=11)
ws["B27"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
ws["B27"].fill = PatternFill(start_color="FFF8E1", end_color="FFF8E1", fill_type="solid")
for rr in range(27, 31):
    ws.row_dimensions[rr].height = 22

# === Блок 4: Кассовый разрыв ===
ws.merge_cells("A32:F32")
ws["A32"] = "КАССОВЫЙ РАЗРЫВ — ЧТО ЭТО ЗНАЧИТ ДЛЯ ДИСТРИБЬЮТОРА"
ws["A32"].font = Font(name="Calibri", size=12, bold=True, color="C00000")
ws["A32"].fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
ws.row_dimensions[32].height = 24

ws.merge_cells("A33:F37")
ws["A33"] = (
    "Кассовый разрыв — это ситуация, когда НАКОПЛЕННЫЙ БАЛАНС становится отрицательным.\n\n"
    "В терминах бизнеса: дистрибьютор должен заплатить поставщику (Эмпилс) и другим кредиторам "
    "бОльшую сумму, чем у него сейчас есть на счетах.\n\n"
    "Что это значит физически:\n"
    "  • Нельзя платить зарплату сотрудникам вовремя\n"
    "  • Просрочка платежей поставщику → штрафы, потеря доверия\n"
    "  • Риск банкротства, если разрыв глубокий и длительный\n\n"
    "В нашем случае: разрыв −2,8 млн руб в апреле. Это не смертельно, но нужно либо иметь "
    "стартовый капитал 3 млн руб, либо договориться с Эмпилс об отсрочке 60 дней."
)
ws["A33"].font = Font(name="Calibri", size=11)
ws["A33"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
ws["A33"].fill = PatternFill(start_color="FFF8E1", end_color="FFF8E1", fill_type="solid")
for rr in range(33, 38):
    ws.row_dimensions[rr].height = 22

# === Блок 5: визуальная аналогия 12 месяцев ===
ws.merge_cells("A40:F40")
ws["A40"] = "ВИЗУАЛЬНАЯ АНАЛОГИЯ: 12 МЕСЯЦЕВ В ОДНОЙ КАРТИНКЕ"
ws["A40"].font = Font(name="Calibri", size=12, bold=True, color="1F4E79")
ws["A40"].fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
ws.row_dimensions[40].height = 24

ws.merge_cells("A41:F41")
ws["A41"] = "Ниже — схема «бассейна»: каждый месяц — отдельный бассейн, заполненный зелёным (резерв) или красным (долг)."
ws["A41"].font = SUBTITLE_FONT
ws["A41"].alignment = Alignment(horizontal="left", vertical="center")
ws.row_dimensions[41].height = 18

# Вставить картинку схемы бассейнов
from openpyxl.drawing.image import Image as XLImage
img = XLImage(img4_path)
img.width = 900
img.height = 500
ws.add_image(img, "A43")

# ==================== ЛИСТ 2: ВИЗУАЛИЗАЦИИ ====================
ws2 = wb.create_sheet("Визуализации")

ws2.merge_cells("A1:N1")
ws2["A1"] = "Три ракурса на кассовый поток: для лучшего понимания"
ws2["A1"].font = TITLE_FONT
ws2.row_dimensions[1].height = 24

# === График 1: Каскад (waterfall) ===
ws2.merge_cells("A3:N3")
ws2["A3"] = "РАКУРС 1: КАСКАД — как из стартового 0 формируется накопленный баланс"
ws2["A3"].font = Font(name="Calibri", size=12, bold=True, color="1F4E79")
ws2["A3"].fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
ws2.row_dimensions[3].height = 24

ws2.merge_cells("A4:N4")
ws2["A4"] = "Каждый столбец = изменение за месяц. Зелёные — приток денег (рост баланса), красные — отток (падение). Линия показывает итог."
ws2["A4"].font = SUBTITLE_FONT
ws2["A4"].alignment = Alignment(horizontal="left", vertical="center")
ws2.row_dimensions[4].height = 18

img1 = XLImage(img1_path)
img1.width = 900
img1.height = 450
ws2.add_image(img1, "A6")

# === График 2: Накопленный баланс с зонами ===
# Place at row 30 (after image 1)
ws2.merge_cells("A30:N30")
ws2["A30"] = "РАКУРС 2: НАКОПЛЕННЫЙ БАЛАНС С ЗОНАМИ — где мы в плюсе, а где в минусе"
ws2["A30"].font = Font(name="Calibri", size=12, bold=True, color="1F4E79")
ws2["A30"].fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
ws2.row_dimensions[30].height = 24

ws2.merge_cells("A31:N31")
ws2["A31"] = "Зелёная зона — деньги есть (резерв). Красная зона — кассовый разрыв (нужно покрывать долг)."
ws2["A31"].font = SUBTITLE_FONT
ws2["A31"].alignment = Alignment(horizontal="left", vertical="center")
ws2.row_dimensions[31].height = 18

img2_xl = XLImage(img2_path)
img2_xl.width = 900
img2_xl.height = 400
ws2.add_image(img2_xl, "A33")

# === График 3: Притоки vs оттоки ===
ws2.merge_cells("A57:N57")
ws2["A57"] = "РАКУРС 3: ПРИТОКИ vs ОТТОКИ — что приходит и что уходит каждый месяц"
ws2["A57"].font = Font(name="Calibri", size=12, bold=True, color="1F4E79")
ws2["A57"].fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
ws2.row_dimensions[57].height = 24

ws2.merge_cells("A58:N58")
ws2["A58"] = "Зелёный столбец = деньги от РТТ (входящий). Красный = платёж Эмпилс (исходящий). Если зелёный выше красного — месяц в плюсе."
ws2["A58"].font = SUBTITLE_FONT
ws2["A58"].alignment = Alignment(horizontal="left", vertical="center")
ws2.row_dimensions[58].height = 18

img3_xl = XLImage(img3_path)
img3_xl.width = 900
img3_xl.height = 400
ws2.add_image(img3_xl, "A60")

# ==================== ЛИСТ 3: КАССОВЫЙ ПОТОК (с подсветкой) ====================
ws3 = wb.create_sheet("Кассовый поток (детально)")

ws3.merge_cells("A1:I1")
ws3["A1"] = "Помесячный кассовый поток с цветовой кодировкой"
ws3["A1"].font = TITLE_FONT
ws3.row_dimensions[1].height = 24

ws3.merge_cells("A2:I2")
ws3["A2"] = ("Чтение таблицы: ЖЁЛТЫЙ — закупка у Эмпилс (первичка), ГОЛУБОЙ — продажа в РТТ (вторичка), "
             "КРАСНЫЙ с минусом — уход денег (платёж Эмпилс), ЗЕЛЁНЫЙ с плюсом — приход денег (от РТТ). "
             "Накопленный баланс — главный показатель: если он в красной зоне — это кассовый разрыв.")
ws3["A2"].font = SUBTITLE_FONT
ws3["A2"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
ws3.row_dimensions[2].height = 50

# Шапка
headers = [
    ("Месяц", 12),
    ("Первичка (закупка)", 16),
    ("Вторичка (продажа в РТТ)", 18),
    ("Платёж Эмпилс (T+45д)", 18),
    ("Деньги от РТТ (T+30д)", 18),
    ("Чистый поток месяца", 18),
    ("Накопленный баланс", 18),
    ("Статус", 14),
    ("Комментарий", 40),
]
for col_idx, (title, width) in enumerate(headers, 1):
    cell = ws3.cell(row=4, column=col_idx, value=title)
    cell.fill = HEAD_FILL
    cell.font = HEAD_FONT
    cell.alignment = HEAD_ALIGN
    cell.border = BORDER
    ws3.column_dimensions[get_column_letter(col_idx)].width = width
ws3.row_dimensions[4].height = 42

# Кастомные комментарии
comments = {
    'Январь': 'Старт. Склада нет, нужно закупить. Денег от РТТ ещё не приходило.',
    'Февраль': 'Первые деньги от РТТ за январскую вторичку. Платить Эмпилс ещё рано.',
    'Март': '⚠ РАЗРЫВ. Платим за январскую первичку (6 млн), но от РТТ пришло только 2,4 млн.',
    'Апрель': '⚠ МИНИМУМ. Платим за февраль (6 млн), а от РТТ пришло 4,8 млн. Нужны 3 млн извне.',
    'Май': 'Улучшение. Пик весенних продаж. От РТТ пришло 7,8 млн, платим 7 млн за март.',
    'Июнь': 'Восстановление. От РТТ 7,8 млн, платим 7 млн за апрель. Разрыв сокращается.',
    'Июль': '✓ ВЫХОД В ПЛЮС. От РТТ 7,75 млн, платим всего 4 млн (за майскую «короткую» первичку).',
    'Август': 'Стабильный плюс. Сезон продолжается, платежи Эмпилс умеренные.',
    'Сентябрь': 'Спад вторички, но платежи Эмпилс тоже спадают.',
    'Октябрь': 'Низкий сезон. Отток превышает приток, но накопленный резерв покрывает.',
    'Ноябрь': 'Зимний спад продолжается. Баланс остаётся положительным.',
    'Декабрь': 'Год закрыт с накопленным резервом +2,36 млн руб — это подушка на следующий год.',
}

for i, cf in enumerate(cash_flow):
    r = 5 + i
    ws3.cell(row=r, column=1, value=cf['month']).font = Font(name="Calibri", size=10, bold=True)
    ws3.cell(row=r, column=1).border = BORDER
    ws3.cell(row=r, column=1).alignment = CENTER_ALIGN
    
    cell = ws3.cell(row=r, column=2, value=cf['primary'])
    cell.font = BODY_FONT
    cell.alignment = NUM_ALIGN
    cell.number_format = '# ##0'
    cell.border = BORDER
    cell.fill = PRIMARY_HIGHLIGHT
    
    cell = ws3.cell(row=r, column=3, value=cf['secondary'])
    cell.font = BODY_FONT
    cell.alignment = NUM_ALIGN
    cell.number_format = '# ##0'
    cell.border = BORDER
    cell.fill = SECONDARY_HIGHLIGHT
    
    cell = ws3.cell(row=r, column=4, value=cf['pay_empils'])
    cell.font = BODY_FONT if cf['pay_empils'] == 0 else Font(name="Calibri", size=10, bold=True, color="C00000")
    cell.alignment = NUM_ALIGN
    cell.number_format = '# ##0;[Red]-# ##0'
    cell.border = BORDER
    cell.fill = OUTCOME_FILL if cf['pay_empils'] > 0 else PatternFill()
    
    cell = ws3.cell(row=r, column=5, value=cf['receive_rtt'])
    cell.font = BODY_FONT if cf['receive_rtt'] == 0 else Font(name="Calibri", size=10, bold=True, color="375623")
    cell.alignment = NUM_ALIGN
    cell.number_format = '# ##0'
    cell.border = BORDER
    cell.fill = INCOME_FILL if cf['receive_rtt'] > 0 else PatternFill()
    
    cell = ws3.cell(row=r, column=6, value=cf['net'])
    if cf['net'] < 0:
        cell.font = NEGATIVE_FONT
        cell.fill = GAP_FILL
    elif cf['net'] > 0:
        cell.font = POSITIVE_FONT
        cell.fill = POSITIVE_FILL
    else:
        cell.font = BODY_FONT
    cell.alignment = NUM_ALIGN
    cell.number_format = '# ##0;[Red]-# ##0'
    cell.border = BORDER
    
    cell = ws3.cell(row=r, column=7, value=cf['cumulative'])
    if cf['cumulative'] < 0:
        cell.font = NEGATIVE_FONT
        cell.fill = GAP_FILL
    elif cf['cumulative'] > 0:
        cell.font = POSITIVE_FONT
        cell.fill = POSITIVE_FILL
    else:
        cell.font = BODY_FONT
    cell.alignment = NUM_ALIGN
    cell.number_format = '# ##0;[Red]-# ##0'
    cell.border = BORDER
    
    if cf['cumulative'] < 0:
        status = "⚠ РАЗРЫВ"
        color = "C00000"
    elif cf['cumulative'] == 0 and i == 0:
        status = "Старт"
        color = "595959"
    else:
        status = "✓ ОК"
        color = "375623"
    cell = ws3.cell(row=r, column=8, value=status)
    cell.font = Font(name="Calibri", size=10, bold=True, color=color)
    cell.alignment = CENTER_ALIGN
    cell.border = BORDER
    
    # Комментарий
    comment = comments.get(cf['month'], '')
    cell = ws3.cell(row=r, column=9, value=comment)
    cell.font = Font(name="Calibri", size=9, italic=True, 
                     color="C00000" if 'РАЗРЫВ' in comment or 'МИНИМУМ' in comment else "595959")
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    cell.border = BORDER
    
    ws3.row_dimensions[r].height = 36

# Итоговая строка
r_total = 5 + len(cash_flow)
ws3.cell(row=r_total, column=1, value="ИТОГО").font = Font(name="Calibri", size=10, bold=True)
ws3.cell(row=r_total, column=1).fill = TOTAL_FILL
ws3.cell(row=r_total, column=1).border = BORDER
ws3.cell(row=r_total, column=1).alignment = CENTER_ALIGN

totals = {
    'primary': sum(cf['primary'] for cf in cash_flow),
    'secondary': sum(cf['secondary'] for cf in cash_flow),
    'pay_empils': sum(cf['pay_empils'] for cf in cash_flow),
    'receive_rtt': sum(cf['receive_rtt'] for cf in cash_flow),
    'net': sum(cf['net'] for cf in cash_flow),
}
for col_idx, key in enumerate(['primary', 'secondary', 'pay_empils', 'receive_rtt', 'net'], 2):
    cell = ws3.cell(row=r_total, column=col_idx, value=totals[key])
    cell.font = Font(name="Calibri", size=10, bold=True)
    cell.fill = TOTAL_FILL
    cell.alignment = NUM_ALIGN
    cell.number_format = '# ##0;[Red]-# ##0'
    cell.border = BORDER

cell = ws3.cell(row=r_total, column=7, value=cash_flow[-1]['cumulative'])
cell.font = Font(name="Calibri", size=10, bold=True, color="375623")
cell.fill = TOTAL_FILL
cell.alignment = NUM_ALIGN
cell.number_format = '# ##0;[Red]-# ##0'
cell.border = BORDER

ws3.cell(row=r_total, column=9, value=f"Год закрыт с резервом {cash_flow[-1]['cumulative']/1e6:+.2f} млн руб").font = Font(name="Calibri", size=10, bold=True, color="375623")
ws3.cell(row=r_total, column=9).fill = TOTAL_FILL
ws3.cell(row=r_total, column=9).alignment = Alignment(horizontal="left", vertical="center")
ws3.cell(row=r_total, column=9).border = BORDER

ws3.freeze_panes = "B5"

# ==================== ЛИСТ 4: ПЛАН ПЕРВИЧКИ И ВТОРИЧКИ ====================
ws4 = wb.create_sheet("План первички и вторички")

ws4.merge_cells("A1:F1")
ws4["A1"] = "Помесячный план первички и вторички с учётом сезонности ЛКМ"
ws4["A1"].font = TITLE_FONT
ws4.row_dimensions[1].height = 24

ws4.merge_cells("A2:F2")
ws4["A2"] = ("Сезонность ЛКМ (по данным открытых источников): пик спроса — июнь (+25% к весне), "
             "строительный сезон — апрель-сентябрь. Зимой (янв-мар) — низкий сезон. "
             "Январь-май: 50% годовой первички (30 млн) — стартовая загрузка склада. "
             "Июнь-дек: 30 млн по убывающей.")
ws4["A2"].font = SUBTITLE_FONT
ws4["A2"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
ws4.row_dimensions[2].height = 50

headers4 = [
    ("Месяц", 12),
    ("Первичка, руб", 16),
    ("Доля, %", 12),
    ("Вторичка, руб", 16),
    ("Доля, %", 12),
    ("Коэфф. опережения (первичка/вторичка)", 22),
]
for col_idx, (title, width) in enumerate(headers4, 1):
    cell = ws4.cell(row=4, column=col_idx, value=title)
    cell.fill = HEAD_FILL
    cell.font = HEAD_FONT
    cell.alignment = HEAD_ALIGN
    cell.border = BORDER
    ws4.column_dimensions[get_column_letter(col_idx)].width = width
ws4.row_dimensions[4].height = 42

for i, cf in enumerate(cash_flow):
    r = 5 + i
    prim = cf['primary']
    sec = cf['secondary']
    prim_pct = prim / 60_000_000
    sec_pct = sec / 60_000_000
    ratio = prim / sec if sec > 0 else 0
    
    ws4.cell(row=r, column=1, value=cf['month']).font = Font(name="Calibri", size=10, bold=True)
    ws4.cell(row=r, column=1).alignment = CENTER_ALIGN
    ws4.cell(row=r, column=1).border = BORDER
    
    cell = ws4.cell(row=r, column=2, value=prim)
    cell.fill = PRIMARY_HIGHLIGHT
    cell.font = BODY_FONT
    cell.alignment = NUM_ALIGN
    cell.number_format = '# ##0'
    cell.border = BORDER
    
    cell = ws4.cell(row=r, column=3, value=prim_pct)
    cell.font = BODY_FONT
    cell.alignment = NUM_ALIGN
    cell.number_format = '0.0%'
    cell.border = BORDER
    
    cell = ws4.cell(row=r, column=4, value=sec)
    cell.fill = SECONDARY_HIGHLIGHT
    cell.font = BODY_FONT
    cell.alignment = NUM_ALIGN
    cell.number_format = '# ##0'
    cell.border = BORDER
    
    cell = ws4.cell(row=r, column=5, value=sec_pct)
    cell.font = BODY_FONT
    cell.alignment = NUM_ALIGN
    cell.number_format = '0.0%'
    cell.border = BORDER
    
    cell = ws4.cell(row=r, column=6, value=ratio)
    if ratio > 1.5:
        cell.font = Font(name="Calibri", size=10, bold=True, color="BF8F00")
        cell.fill = WARN_FILL
    elif ratio > 1.0:
        cell.font = Font(name="Calibri", size=10, bold=True, color="595959")
    else:
        cell.font = Font(name="Calibri", size=10, color="375623")
    cell.alignment = NUM_ALIGN
    cell.number_format = '0.00'
    cell.border = BORDER

r_total = 5 + len(cash_flow)
ws4.cell(row=r_total, column=1, value="ИТОГО").font = Font(name="Calibri", size=10, bold=True)
ws4.cell(row=r_total, column=1).fill = TOTAL_FILL
ws4.cell(row=r_total, column=1).border = BORDER
ws4.cell(row=r_total, column=1).alignment = CENTER_ALIGN

cell = ws4.cell(row=r_total, column=2, value=sum(cf['primary'] for cf in cash_flow))
cell.font = Font(name="Calibri", size=10, bold=True)
cell.fill = TOTAL_FILL
cell.alignment = NUM_ALIGN
cell.number_format = '# ##0'
cell.border = BORDER

cell = ws4.cell(row=r_total, column=3, value=1.0)
cell.font = Font(name="Calibri", size=10, bold=True)
cell.fill = TOTAL_FILL
cell.alignment = NUM_ALIGN
cell.number_format = '0.0%'
cell.border = BORDER

cell = ws4.cell(row=r_total, column=4, value=sum(cf['secondary'] for cf in cash_flow))
cell.font = Font(name="Calibri", size=10, bold=True)
cell.fill = TOTAL_FILL
cell.alignment = NUM_ALIGN
cell.number_format = '# ##0'
cell.border = BORDER

cell = ws4.cell(row=r_total, column=5, value=sum(cf['secondary'] for cf in cash_flow)/60_000_000)
cell.font = Font(name="Calibri", size=10, bold=True)
cell.fill = TOTAL_FILL
cell.alignment = NUM_ALIGN
cell.number_format = '0.0%'
cell.border = BORDER

# Добавить комментарий под таблицей
r = r_total + 3
ws4.merge_cells(f"A{r}:F{r}")
ws4.cell(row=r, column=1, value="ПОЯСНЕНИЕ КОЭФФИЦИЕНТА ОПЕРЕЖЕНИЯ").font = Font(name="Calibri", size=11, bold=True, color="1F4E79")

ws4.merge_cells(f"A{r+1}:F{r+5}")
ws4.cell(row=r+1, column=1, value=(
    "Коэффициент опережения = Первичка / Вторичка.\n\n"
    "Если > 1.0 (жёлтый/оранжевый) — дистрибьютор закупает БОЛЬШЕ, чем продаёт в этом месяце. "
    "Это нормально на старте: склад заполняется товаром под будущие продажи.\n\n"
    "Если < 1.0 (зелёный) — дистрибьютор продаёт больше, чем закупает в текущем месяце. "
    "Это значит, что товар «срабатывается» со склада. Так и должно быть в пик сезона.\n\n"
    "В январе коэффициент = 3,0 — это самая агрессивная фаза запуска (первичка 6 млн, вторичка 2 млн). "
    "В апреле-мае коэффициент ≈ 0,9-0,5 — начинаем «срабатывать» склад."
)).font = BODY_FONT
ws4.cell(row=r+1, column=1).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
ws4.cell(row=r+1, column=1).fill = PatternFill(start_color="FFF8E1", end_color="FFF8E1", fill_type="solid")
for rr in range(r+1, r+6):
    ws4.row_dimensions[rr].height = 22

# ==================== ЛИСТ 5: АНАЛИЗ РАЗРЫВА И РЕШЕНИЯ ====================
ws5 = wb.create_sheet("Анализ разрыва")

ws5.merge_cells("A1:F1")
ws5["A1"] = "Анализ кассового разрыва и варианты решения"
ws5["A1"].font = TITLE_FONT
ws5.row_dimensions[1].height = 28

# Раздел 1: точка разрыва
ws5["A3"] = "1. ТОЧКА КАССОВОГО РАЗРЫВА"
ws5["A3"].font = Font(name="Calibri", size=12, bold=True, color="C00000")
ws5.merge_cells("A3:F3")

gap_data = [
    ("Первый месяц с отрицательным балансом", "МАРТ"),
    ("Накопленный баланс в марте", "-1 600 000 руб"),
    ("Месяц с минимальным балансом", "АПРЕЛЬ"),
    ("Минимальный накопленный баланс", "-2 800 000 руб"),
    ("Глубина разрыва (продолжительность)", "2 месяца (март-апрель)"),
    ("Когда баланс восстанавливается", "ИЮЛЬ (накопл. +2,55 млн руб)"),
    ("Долг Эмпилс на конец апреля", "14 000 000 руб"),
]

for col_idx, (title, width) in enumerate([("Параметр", 45), ("Значение", 35)], 1):
    cell = ws5.cell(row=4, column=col_idx, value=title)
    cell.fill = HEAD_FILL
    cell.font = HEAD_FONT
    cell.alignment = HEAD_ALIGN
    cell.border = BORDER
    ws5.column_dimensions[get_column_letter(col_idx)].width = width

for i, (param, val) in enumerate(gap_data, 5):
    ws5.cell(row=i, column=1, value=param).font = Font(name="Calibri", size=10, bold=True)
    ws5.cell(row=i, column=1).border = BORDER
    ws5.cell(row=i, column=1).alignment = Alignment(horizontal="left", vertical="center")
    
    cell = ws5.cell(row=i, column=2, value=val)
    if "разрыв" in param.lower() or "миним" in param.lower() or "отриц" in param.lower() or "МАРТ" in val or "АПРЕЛЬ" in val:
        cell.font = Font(name="Calibri", size=10, bold=True, color="C00000")
        cell.fill = GAP_FILL
    else:
        cell.font = Font(name="Calibri", size=10, bold=True, color="375623")
        cell.fill = POSITIVE_FILL
    cell.alignment = CENTER_ALIGN
    cell.border = BORDER

# Раздел 2: причины
r = 14
ws5.cell(row=r, column=1, value="2. ПРИЧИНЫ РАЗРЫВА").font = Font(name="Calibri", size=12, bold=True, color="1F4E79")
ws5.merge_cells(f"A{r}:F{r}")

reasons = [
    ("Стартовая загрузка склада в январе", "6 млн руб первички в январе — это 10% годового плана за один месяц. Дистрибьютор должен выкупить у Эмпилс 300 SKU, чтобы заполнить полки РТТ."),
    ("Опережение вторички первичкой", "В январе первичка 6 млн, а вторичка всего 2 млн (коэффициент опережения 3.0). В феврале — 6/2,4 = 2,5. Это нормальная логика дистрибьютора: сначала склад, потом продажи."),
    ("45-дневная отсрочка недостаточна для январской закупки", "Январская первичка (6 млн) должна быть оплачена в марте. К этому моменту деньги от РТТ за январскую вторичку (2 млн) и февральскую (2,4 млн) уже пришли (4,4 млн), но это покрывает только 73% долга Эмпилс."),
    ("Мартовская «контрактная вилка»", "В марте нужно одновременно платить за январь (6 млн) И продолжать закупать (первичка марта = 7 млн). Суммарная потребность в марте = 13 млн, при поступлениях от РТТ только 2,4 млн."),
    ("Минимум в апреле", "В апреле приходит платеж за февральскую первичку (6 млн), но деньги от РТТ за мартовскую вторичку (4,8 млн) ещё не покрывают его. Дополнительно нужна апрельская первичка 7 млн."),
]

for i, (cause, detail) in enumerate(reasons):
    rr = 15 + i
    ws5.cell(row=rr, column=1, value=cause).font = Font(name="Calibri", size=10, bold=True, color="C00000")
    ws5.cell(row=rr, column=1).border = BORDER
    ws5.cell(row=rr, column=1).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws5.merge_cells(f"B{rr}:F{rr}")
    ws5.cell(row=rr, column=2, value=detail).font = BODY_FONT
    ws5.cell(row=rr, column=2).border = BORDER
    ws5.cell(row=rr, column=2).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws5.row_dimensions[rr].height = 50

# Раздел 3: варианты решения
r = 22
ws5.cell(row=r, column=1, value="3. ВАРИАНТЫ РЕШЕНИЯ").font = Font(name="Calibri", size=12, bold=True, color="375623")
ws5.merge_cells(f"A{r}:F{r}")

solutions = [
    ("Вариант", "Описание", "Эффект на разрыв", "Сложность"),
    ("А. Стартовый капитал 3 млн руб", "Дистрибьютор вносит собственные средства на старте (январь), чтобы покрыть разрыв марта-апреля", "Полное закрытие разрыва (баланс всегда ≥ 0)", "Низкая — нужно найти 3 млн"),
    ("Б. Увеличить отсрочку до 60 дней", "Просить Эмпилс увеличить отсрочку с 45 до 60 дней (доп. 15 дней)", "Сдвиг платежей на 1 месяц вперёд: разрыв начнётся в апреле (-1,2 млн), но будет мягче", "Средняя — нужны переговоры"),
    ("В. Снизить январскую первичку до 4 млн", "Уменьшить стартовую загрузку: 4 млн в январе вместо 6 млн", "Разрыв начнётся в марте с -0,8 млн, минимум -2,0 млн в апреле. Меньше глубина.", "Низкая — но риск дефицита на полках"),
    ("Г. Предоплата от РТТ первые 3 мес", "Договориться с РТТ о предоплате в январе-марте (вместо постоплаты 30 дн)", "Деньги приходят сразу: январь +2 млн, февраль +2,4 млн, март +4,8 млн. Разрыв полностью закрывается.", "Высокая — РТТ редко соглашаются"),
    ("Д. Комбинация: кап. 1,5 млн + отсрочка 60 дн", "Минимальный капитал 1,5 млн и отсрочка 60 дней", "Полное закрытие разрыва с запасом", "Средняя"),
]

for col_idx, (title, width) in enumerate([("Вариант", 30), ("Описание", 50), ("Эффект на разрыв", 40), ("Сложность", 25)], 1):
    cell = ws5.cell(row=23, column=col_idx, value=title)
    cell.fill = HEAD_FILL
    cell.font = HEAD_FONT
    cell.alignment = HEAD_ALIGN
    cell.border = BORDER
    ws5.column_dimensions[get_column_letter(col_idx)].width = width

for i, (var, desc, eff, diff) in enumerate(solutions[1:], 24):
    ws5.cell(row=i, column=1, value=var).font = Font(name="Calibri", size=10, bold=True)
    ws5.cell(row=i, column=1).border = BORDER
    ws5.cell(row=i, column=1).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    
    ws5.cell(row=i, column=2, value=desc).font = BODY_FONT
    ws5.cell(row=i, column=2).border = BORDER
    ws5.cell(row=i, column=2).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    
    ws5.cell(row=i, column=3, value=eff).font = BODY_FONT
    ws5.cell(row=i, column=3).border = BORDER
    ws5.cell(row=i, column=3).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    
    ws5.cell(row=i, column=4, value=diff).font = BODY_FONT
    ws5.cell(row=i, column=4).border = BORDER
    ws5.cell(row=i, column=4).alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
    
    ws5.row_dimensions[i].height = 50

# Раздел 4: вывод
r = 30
ws5.cell(row=r, column=1, value="4. ВЫВОД ПО ЛОГИКЕ ДИСТРИБЬЮТОРСКОГО БИЗНЕСА В ЛКМ").font = Font(name="Calibri", size=12, bold=True, color="1F4E79")
ws5.merge_cells(f"A{r}:F{r}")

conclusion = (
    "Логика подтверждена. Дистрибьюторский бизнес в ЛКМ имеет ярко выраженный кассовый разрыв на старте по следующей причине:\n\n"
    "1. СЕЗОННОСТЬ: пик спроса в июне (+25% к весне, по данным открытых источников). Чтобы встретить пик, нужно в январе-мае сделать 50% годовой первички (30 млн руб).\n"
    "2. ОТСРОЧКА 45 ДНЕЙ НЕДОСТАТОЧНА: при постоплате 45 дней январская первичка (6 млн) оплачивается в марте — раньше, чем дистрибьютор успевает распродать её через РТТ и собрать деньги (с учётом 30-дневной отсрочки от РТТ).\n"
    "3. ОПЕРЕЖЕНИЕ ПЕРВИЧКИ ВТОРИЧКОЙ: в январе коэффициент опережения = 3,0 (6/2), в феврале — 2,5 (6/2,4). Это нормально для фазы запуска, но требует оборотного капитала.\n"
    "4. ВОССТАНОВЛЕНИЕ: с июня график выходит в плюс (накопленный баланс +2,55 млн в июле). К концу года накопленный излишек +2,36 млн — это резерв на следующий сезон.\n\n"
    "РЕКОМЕНДАЦИЯ ЭМПИЛС: при переговорах с Олимпом предложить гибридный вариант (Д) — стартовый капитал 1,5 млн руб со стороны Олимпа + увеличенная отсрочка 60 дней от Эмпилс на первые 6 месяцев сотрудничества. "
    "Это снимет кассовый разрыв и создаст доверие к партнёрству."
)
ws5.merge_cells(f"A{r+1}:F{r+8}")
ws5.cell(row=r+1, column=1, value=conclusion).font = BODY_FONT
ws5.cell(row=r+1, column=1).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
ws5.cell(row=r+1, column=1).fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
for rr in range(r+1, r+9):
    ws5.row_dimensions[rr].height = 22

# ==================== ЛИСТ 6: ПАРАМЕТРЫ И АССОРТИМЕНТ ====================
ws6 = wb.create_sheet("Параметры модели")

ws6.merge_cells("A1:C1")
ws6["A1"] = "Параметры модели и ассортимент"
ws6["A1"].font = TITLE_FONT
ws6.row_dimensions[1].height = 24

# Параметры
ws6["A3"] = "КЛЮЧЕВЫЕ ПАРАМЕТРЫ"
ws6["A3"].font = Font(name="Calibri", size=12, bold=True, color="1F4E79")
ws6.merge_cells("A3:C3")

params = [
    ("Годовой товарооборот (план)", "60 000 000 руб", "Из КП Эмпилс от 12.10.2026, Год 1"),
    ("Наценка дистрибьютора (Дистр→РТТ)", "35%", "Из КП: заводская цена × 1.35 = цена для РТТ"),
    ("Наценка РТТ (РТТ→Потребитель)", "50%", "Из КП: цена для РТТ × 1.50 = розничная цена"),
    ("Условия оплаты Эмпилс→Дистр", "Постоплата 45 дней", "Из КП: «диапазон 30-45 дней» — берём максимум 45"),
    ("Условия оплаты РТТ→Дистр", "Постоплата 30 дней", "Типовая рыночная практика по РТТ"),
    ("Стартовый капитал дистрибьютора", "0 руб", "Моделируем «без вливаний», чтобы увидеть разрыв"),
    ("Количество SKU в матрице", "300", "Ассортимент Олимп v4"),
]

for col_idx, (title, width) in enumerate([("Параметр", 38), ("Значение", 24), ("Источник / комментарий", 50)], 1):
    cell = ws6.cell(row=4, column=col_idx, value=title)
    cell.fill = HEAD_FILL
    cell.font = HEAD_FONT
    cell.alignment = HEAD_ALIGN
    cell.border = BORDER
    ws6.column_dimensions[get_column_letter(col_idx)].width = width
ws6.row_dimensions[4].height = 32

for i, (param, val, src) in enumerate(params):
    r = 5 + i
    ws6.cell(row=r, column=1, value=param).font = Font(name="Calibri", size=10, bold=True)
    ws6.cell(row=r, column=1).border = BORDER
    ws6.cell(row=r, column=1).alignment = Alignment(horizontal="left", vertical="center")
    
    ws6.cell(row=r, column=2, value=val).font = Font(name="Calibri", size=10, bold=True, color="1F4E79")
    ws6.cell(row=r, column=2).border = BORDER
    ws6.cell(row=r, column=2).alignment = CENTER_ALIGN
    
    ws6.cell(row=r, column=3, value=src).font = BODY_FONT
    ws6.cell(row=r, column=3).border = BORDER
    ws6.cell(row=r, column=3).alignment = BODY_ALIGN

# Сегментация ассортимента
r = 5 + len(params) + 2
ws6.cell(row=r, column=1, value="СЕГМЕНТАЦИЯ АССОРТИМЕНТА ПО ЦЕНАМ").font = Font(name="Calibri", size=12, bold=True, color="1F4E79")
ws6.merge_cells(f"A{r}:C{r}")
r += 1

# Загрузить ассортимент v4 для сегментации
wb_src = openpyxl.load_workbook("/home/z/my-project/upload/Ассортимент_Олимп_300_SKU_финал_v4.xlsx", data_only=True)
ws_src = wb_src["300 SKU с ценами"]
all_skus = []
for rr in range(5, ws_src.max_row + 1):
    price = ws_src.cell(row=rr, column=10).value
    if price and isinstance(price, (int, float)):
        all_skus.append(price)

total_count = len(all_skus)
total_sum = sum(all_skus)

segments = [
    ("До 100 руб", 0, 100),
    ("100-300 руб", 100, 300),
    ("300-500 руб", 300, 500),
    ("500-1000 руб", 500, 1000),
    ("1000-2000 руб", 1000, 2000),
    ("2000-5000 руб", 2000, 5000),
    ("Более 5000 руб", 5000, 999999),
]

for col_idx, (title, width) in enumerate([("Ценовой сегмент", 22), ("Кол-во SKU", 14), ("Сумма, руб", 18)], 1):
    cell = ws6.cell(row=r, column=col_idx, value=title)
    cell.fill = HEAD_FILL
    cell.font = HEAD_FONT
    cell.alignment = HEAD_ALIGN
    cell.border = BORDER
ws6.row_dimensions[r].height = 32
r += 1

for seg_name, lo, hi in segments:
    in_seg = [p for p in all_skus if lo <= p < hi]
    count = len(in_seg)
    seg_sum = sum(in_seg)
    
    ws6.cell(row=r, column=1, value=seg_name).font = Font(name="Calibri", size=10, bold=True)
    ws6.cell(row=r, column=1).border = BORDER
    ws6.cell(row=r, column=1).alignment = Alignment(horizontal="left", vertical="center")
    
    ws6.cell(row=r, column=2, value=count).font = BODY_FONT
    ws6.cell(row=r, column=2).border = BORDER
    ws6.cell(row=r, column=2).alignment = NUM_ALIGN
    
    ws6.cell(row=r, column=3, value=seg_sum).font = BODY_FONT
    ws6.cell(row=r, column=3).border = BORDER
    ws6.cell(row=r, column=3).alignment = NUM_ALIGN
    ws6.cell(row=r, column=3).number_format = '# ##0'
    r += 1

# Итого
ws6.cell(row=r, column=1, value="ИТОГО 300 SKU").font = Font(name="Calibri", size=10, bold=True)
ws6.cell(row=r, column=1).fill = TOTAL_FILL
ws6.cell(row=r, column=1).border = BORDER

ws6.cell(row=r, column=2, value=total_count).font = Font(name="Calibri", size=10, bold=True)
ws6.cell(row=r, column=2).fill = TOTAL_FILL
ws6.cell(row=r, column=2).border = BORDER
ws6.cell(row=r, column=2).alignment = NUM_ALIGN

ws6.cell(row=r, column=3, value=total_sum).font = Font(name="Calibri", size=10, bold=True)
ws6.cell(row=r, column=3).fill = TOTAL_FILL
ws6.cell(row=r, column=3).border = BORDER
ws6.cell(row=r, column=3).alignment = NUM_ALIGN
ws6.cell(row=r, column=3).number_format = '# ##0'

# Сохранить
wb.save(OUT_PATH)
print(f"\n✓ Файл сохранён: {OUT_PATH}")
print(f"  Размер: {os.path.getsize(OUT_PATH):,} байт")
print(f"  Листы: {wb.sheetnames}")
