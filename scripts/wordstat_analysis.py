"""
Wordstat API — анализ собранных данных.
Создаёт Excel + визуализации + Word отчёт с рекомендациями.
"""
import json
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import matplotlib
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import numpy as np

# Шрифт
fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf')
fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf')
plt.rcParams.update({
    'font.sans-serif': ['DejaVu Sans'],
    'axes.unicode_minus': False,
    'figure.facecolor': '#FFFFFF',
    'axes.facecolor': '#FFFFFF',
    'axes.spines.top': False,
    'axes.spines.right': False,
    'axes.grid': False,
    'xtick.major.size': 0,
    'ytick.major.size': 0,
})

# Цвета
ORANGE = '#F47629'
DARK = '#313131'
GREY = '#D1D5DB'

# Загрузка данных
with open('/home/z/my-project/download/wordstat_titan_raw.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

CHART_DIR = '/home/z/my-project/scripts/charts_wordstat'
os.makedirs(CHART_DIR, exist_ok=True)

# ============================================================
# 1. EXCEL-ТАБЛИЦА
# ============================================================
wb = Workbook()

# Стиль
ORANGE_FILL = PatternFill('solid', fgColor='F47629')
DARK_FILL = PatternFill('solid', fgColor='313131')
WHITE_FONT = Font(name='Arial', size=11, bold=True, color='FFFFFF')
ORANGE_FONT = Font(name='Arial', size=11, bold=True, color='F47629')
NORMAL_FONT = Font(name='Arial', size=10)
THIN_BORDER = Border(
    left=Side(style='thin', color='D0D0D0'),
    right=Side(style='thin', color='D0D0D0'),
    top=Side(style='thin', color='D0D0D0'),
    bottom=Side(style='thin', color='D0D0D0'),
)

# --- Лист 1: Сводка по фразам ---
ws1 = wb.active
ws1.title = 'Сводка по фразам'

# Заголовок
ws1.merge_cells('A1:E1')
ws1['A1'] = 'СВОДКА ПО КЛЮЧЕВЫМ ФРАЗАМ — Wordstat API'
ws1['A1'].font = WHITE_FONT
ws1['A1'].fill = ORANGE_FILL
ws1['A1'].alignment = Alignment(horizontal='center', vertical='center')
ws1.row_dimensions[1].height = 30

# Заголовки
headers = ['№', 'Ключевая фраза', 'Показов в месяц', 'Регионов с показами', 'Приоритет для Директ']
for i, h in enumerate(headers, 1):
    c = ws1.cell(row=3, column=i, value=h)
    c.font = WHITE_FONT
    c.fill = DARK_FILL
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    c.border = THIN_BORDER
ws1.row_dimensions[3].height = 30

# Данные
sorted_phrases = sorted(data.items(), key=lambda x: int(x[1].get('total_count', '0')), reverse=True)
for ri, (phrase, pdata) in enumerate(sorted_phrases, 1):
    total = int(pdata.get('total_count', '0'))
    n_reg = len(pdata.get('regions', []))
    # Приоритет
    if total >= 10000:
        priority = '🔴 Высокий'
    elif total >= 1000:
        priority = '🟡 Средний'
    elif total >= 100:
        priority = '🟢 Низкий'
    else:
        priority = '⚪ Не использовать'

    row_data = [ri, phrase, total, n_reg, priority]
    for ci, val in enumerate(row_data, 1):
        c = ws1.cell(row=3 + ri, column=ci, value=val)
        c.font = NORMAL_FONT
        c.border = THIN_BORDER
        if ci == 1:
            c.alignment = Alignment(horizontal='center')
        elif ci == 3:
            c.alignment = Alignment(horizontal='center')
            c.font = Font(name='Arial', size=10, bold=True, color='F47629')
        elif ci == 4:
            c.alignment = Alignment(horizontal='center')
        elif ci == 5:
            c.alignment = Alignment(horizontal='center')

# Ширины колонок
for col, w in {'A': 5, 'B': 35, 'C': 18, 'D': 18, 'E': 20}.items():
    ws1.column_dimensions[col].width = w

# --- Лист 2: Топ запросов по каждой фразе ---
ws2 = wb.create_sheet('Топ запросов')
ws2.merge_cells('A1:D1')
ws2['A1'] = 'ТОП ПОИСКОВЫХ ЗАПРОСОВ ПО КАЖДОЙ ФРАЗЕ'
ws2['A1'].font = WHITE_FONT
ws2['A1'].fill = ORANGE_FILL
ws2['A1'].alignment = Alignment(horizontal='center')
ws2.row_dimensions[1].height = 25

headers2 = ['Ключевая фраза', 'Связанный запрос', 'Показов', 'Группа']
for i, h in enumerate(headers2, 1):
    c = ws2.cell(row=3, column=i, value=h)
    c.font = WHITE_FONT
    c.fill = DARK_FILL
    c.alignment = Alignment(horizontal='center')
    c.border = THIN_BORDER

# Группы
groups = {}
group_names = [
    'Брендовые', 'Брендовые', 'Брендовые', 'Брендовые', 'Брендовые',
    'Продуктовые', 'Продуктовые', 'Продуктовые', 'Продуктовые',
    'Продуктовые', 'Продуктовые', 'Продуктовые', 'Продуктовые',
    'Продуктовые', 'Продуктовые',
    'Применение', 'Применение', 'Применение', 'Применение', 'Применение', 'Применение',
    'Конкурентные', 'Конкурентные', 'Конкурентные', 'Конкурентные',
]
for i, kw in enumerate(data.keys()):
    groups[kw] = group_names[i] if i < len(group_names) else 'Другое'

row = 4
for phrase, pdata in data.items():
    results = pdata.get('top_results', [])
    for item in results[:20]:  # Топ-20 на фразу
        related = item.get('phrase', '')
        count = int(item.get('count', '0'))
        if count < 10:
            continue
        ws2.cell(row=row, column=1, value=phrase).font = NORMAL_FONT
        ws2.cell(row=row, column=2, value=related).font = NORMAL_FONT
        c = ws2.cell(row=row, column=3, value=count)
        c.font = Font(name='Arial', size=10, bold=True, color='F47629')
        c.alignment = Alignment(horizontal='center')
        ws2.cell(row=row, column=4, value=groups.get(phrase, '')).font = NORMAL_FONT
        for ci in range(1, 5):
            ws2.cell(row=row, column=ci).border = THIN_BORDER
        row += 1

for col, w in {'A': 30, 'B': 45, 'C': 15, 'D': 15}.items():
    ws2.column_dimensions[col].width = w

# --- Лист 3: Топ-20 регионов по топ-5 фразам ---
ws3 = wb.create_sheet('Топ регионы')
ws3.merge_cells('A1:E1')
ws3['A1'] = 'ТОП-20 РЕГИОНОВ ПО ТОП-5 ФРАЗАМ'
ws3['A1'].font = WHITE_FONT
ws3['A1'].fill = ORANGE_FILL
ws3['A1'].alignment = Alignment(horizontal='center')
ws3.row_dimensions[1].height = 25

# Топ-5 фраз по объёму
top5_phrases = [p for p, d in sorted_phrases[:5]]

headers3 = ['Регион ID', 'Фраза', 'Показов', 'Доля, %', 'Affinity']
for i, h in enumerate(headers3, 1):
    c = ws3.cell(row=3, column=i, value=h)
    c.font = WHITE_FONT
    c.fill = DARK_FILL
    c.alignment = Alignment(horizontal='center')
    c.border = THIN_BORDER

row = 4
for phrase in top5_phrases:
    pdata = data[phrase]
    regions = pdata.get('regions', [])
    # Сортируем по count
    regions_sorted = sorted(regions, key=lambda x: int(x.get('count', '0')), reverse=True)
    for item in regions_sorted[:20]:
        region_id = item.get('region', '')
        count = int(item.get('count', '0'))
        share = item.get('share', 0) * 100
        affinity = item.get('affinityIndex', 0)

        ws3.cell(row=row, column=1, value=region_id).font = NORMAL_FONT
        ws3.cell(row=row, column=2, value=phrase).font = NORMAL_FONT
        c = ws3.cell(row=row, column=3, value=count)
        c.font = Font(name='Arial', size=10, bold=True, color='F47629')
        c.alignment = Alignment(horizontal='center')
        c2 = ws3.cell(row=row, column=4, value=f'{share:.2f}%')
        c2.font = NORMAL_FONT
        c2.alignment = Alignment(horizontal='center')
        c3 = ws3.cell(row=row, column=5, value=f'{affinity:.1f}')
        c3.font = NORMAL_FONT
        c3.alignment = Alignment(horizontal='center')
        for ci in range(1, 6):
            ws3.cell(row=row, column=ci).border = THIN_BORDER
        row += 1

for col, w in {'A': 12, 'B': 30, 'C': 15, 'D': 12, 'E': 12}.items():
    ws3.column_dimensions[col].width = w

# --- Лист 4: Динамика по месяцам ---
ws4 = wb.create_sheet('Динамика по месяцам')
ws4.merge_cells('A1:N1')
ws4['A1'] = 'ДИНАМИКА ПОИСКОВЫХ ЗАПРОСОВ ПО МЕСЯЦАМ (сент 2025 — авг 2026)'
ws4['A1'].font = WHITE_FONT
ws4['A1'].fill = ORANGE_FILL
ws4['A1'].alignment = Alignment(horizontal='center')
ws4.row_dimensions[1].height = 25

# Месяцы
months = ['Сен 25', 'Окт 25', 'Ноя 25', 'Дек 25', 'Янв 26', 'Фев 26',
          'Мар 26', 'Апр 26', 'Май 26', 'Июн 26', 'Июл 26', 'Авг 26']
headers4 = ['Фраза'] + months
for i, h in enumerate(headers4, 1):
    c = ws4.cell(row=3, column=i, value=h)
    c.font = WHITE_FONT
    c.fill = DARK_FILL
    c.alignment = Alignment(horizontal='center')
    c.border = THIN_BORDER

row = 4
for phrase, pdata in sorted_phrases:
    dynamics = pdata.get('dynamics', [])
    ws4.cell(row=row, column=1, value=phrase).font = Font(name='Arial', size=10, bold=True, color='F47629')
    ws4.cell(row=row, column=1).border = THIN_BORDER
    for mi, item in enumerate(dynamics):
        count = int(item.get('count', '0'))
        c = ws4.cell(row=row, column=2 + mi, value=count)
        c.font = NORMAL_FONT
        c.alignment = Alignment(horizontal='center')
        c.border = THIN_BORDER
    row += 1

ws4.column_dimensions['A'].width = 35
for col in 'BCDEFGHIJKLMN':
    ws4.column_dimensions[col].width = 12

# Сохранение Excel
excel_path = '/home/z/my-project/download/Wordstat_TITAN_анализ.xlsx'
wb.save(excel_path)
print(f'Excel saved: {excel_path} ({os.path.getsize(excel_path):,} bytes)')

# ============================================================
# 2. ВИЗУАЛИЗАЦИИ
# ============================================================

# --- График 1: Топ-15 фраз по объёму показов ---
top15 = sorted_phrases[:15]
labels = [p[0] for p in top15]
counts = [int(p[1].get('total_count', '0')) for p in top15]

fig, ax = plt.subplots(figsize=(12, 7), constrained_layout=True)
y_pos = np.arange(len(labels))
colors = [ORANGE if i < 5 else GREY for i in range(len(labels))]
bars = ax.barh(y_pos, counts, color=colors, height=0.6, edgecolor='white')
ax.set_yticks(y_pos)
ax.set_yticklabels(labels, fontsize=10)
ax.invert_yaxis()
ax.set_xlabel('Показов в месяц', fontsize=11, color=DARK)
ax.set_xlim(0, max(counts) * 1.2)
for bar, count in zip(bars, counts):
    w = bar.get_width()
    ax.text(w + max(counts)*0.01, bar.get_y() + bar.get_height()/2,
            f'{count:,}'.replace(',', ' '),
            va='center', ha='left', fontsize=9, color=DARK)
ax.set_title('Топ-15 ключевых фраз по объёму показов\n(Wordstat, сентябрь 2026)',
             fontsize=14, fontweight='bold', color=DARK, pad=16)
fig.savefig(os.path.join(CHART_DIR, 'top15_phrases.png'), dpi=200, bbox_inches='tight', facecolor='white')
plt.close(fig)
print('Chart 1 saved')

# --- График 2: Динамика по месяцам для топ-5 фраз ---
fig2, ax2 = plt.subplots(figsize=(12, 6), constrained_layout=True)
top5_labels = [p[0] for p in sorted_phrases[:5]]
colors5 = [ORANGE, '#FBA668', '#F5C896', '#6B7280', '#D1D5DB']
for i, phrase in enumerate(top5_labels[:5]):
    dynamics = data[phrase].get('dynamics', [])
    months_labels = [item['date'][:7] for item in dynamics]
    counts = [int(item.get('count', '0')) for item in dynamics]
    ax2.plot(months_labels, counts, marker='o', linewidth=2, markersize=5,
             label=phrase[:30], color=colors5[i])
ax2.set_xlabel('Месяц', fontsize=11, color=DARK)
ax2.set_ylabel('Показов', fontsize=11, color=DARK)
ax2.set_title('Динамика поисковых запросов по месяцам\n(Топ-5 фраз, сент 2025 — авг 2026)',
              fontsize=14, fontweight='bold', color=DARK, pad=16)
ax2.legend(loc='upper left', bbox_to_anchor=(1.02, 1), fontsize=9)
ax2.tick_params(axis='x', rotation=45)
fig2.savefig(os.path.join(CHART_DIR, 'dynamics_top5.png'), dpi=200, bbox_inches='tight', facecolor='white')
plt.close(fig2)
print('Chart 2 saved')

# --- График 3: Топ-10 регионов по фразе "краска по металлу" (самая популярная) ---
top_phrase = sorted_phrases[0][0]
regions = data[top_phrase].get('regions', [])
regions_sorted = sorted(regions, key=lambda x: int(x.get('count', '0')), reverse=True)[:10]

# Маппинг region_id -> название
region_names = {
    '225': 'Вся Россия', '213': 'Москва', '3': 'Санкт-Петербург',
    '1': 'Московская обл.', '40': 'Краснодарский край', '17': 'Ростовская обл.',
    '26': 'Свердловская обл.', '59': 'Республика Татарстан', '52': 'Новосибирская обл.',
    '66': 'Свердловская обл.', '11144': 'Пермский край',
    '11086': 'Челябинская обл.', '10174': 'Самарская обл.',
    '54': 'Нижегородская обл.', '37': 'Волгоградская обл.',
    '38': 'Воронежская обл.', '4': 'Краснодарский край',
    '5': 'Республика Адыгея', '21': 'Республика Дагестан',
    '39': 'Республика Башкортостан', '75': 'Челябинская обл.',
    '76': 'Ярославская обл.', '77': 'Ставропольский край',
    '97': 'Иркутская обл.', '102': 'Республика Саха (Якутия)',
}
reg_labels = [region_names.get(item['region'], f"Регион {item['region']}") for item in regions_sorted]
reg_counts = [int(item.get('count', '0')) for item in regions_sorted]

fig3, ax3 = plt.subplots(figsize=(10, 6), constrained_layout=True)
y_pos = np.arange(len(reg_labels))
bars = ax3.barh(y_pos, reg_counts, color=[ORANGE if i < 3 else GREY for i in range(len(reg_labels))],
                height=0.6, edgecolor='white')
ax3.set_yticks(y_pos)
ax3.set_yticklabels(reg_labels, fontsize=10)
ax3.invert_yaxis()
ax3.set_xlabel('Показов в месяц', fontsize=11, color=DARK)
for bar, count in zip(bars, reg_counts):
    w = bar.get_width()
    ax3.text(w + max(reg_counts)*0.01, bar.get_y() + bar.get_height()/2,
             f'{count:,}'.replace(',', ' '),
             va='center', ha='left', fontsize=9, color=DARK)
ax3.set_title(f'Топ-10 регионов по фразе «{top_phrase}»\n(Wordstat, сентябрь 2026)',
              fontsize=14, fontweight='bold', color=DARK, pad=16)
fig3.savefig(os.path.join(CHART_DIR, 'top10_regions.png'), dpi=200, bbox_inches='tight', facecolor='white')
plt.close(fig3)
print('Chart 3 saved')

# --- График 4: Сезонность — усреднённая динамика по топ-5 ---
fig4, ax4 = plt.subplots(figsize=(10, 5), constrained_layout=True)
# Усредняем долю (share) по топ-5 фразам
avg_shares = []
months_labels = []
for phrase in top5_labels[:5]:
    dynamics = data[phrase].get('dynamics', [])
    if not months_labels:
        months_labels = [item['date'][:7] for item in dynamics]
    shares = [float(item.get('share', 0)) for item in dynamics]
    if not avg_shares:
        avg_shares = shares
    else:
        for i in range(min(len(avg_shares), len(shares))):
            avg_shares[i] += shares[i]
# Нормализуем
avg_shares = [s / len(top5_labels[:5]) for s in avg_shares]
# Конвертируем в индекс (100 = средний месяц)
avg_val = sum(avg_shares) / len(avg_shares)
indices = [s / avg_val * 100 for s in avg_shares]

ax4.fill_between(range(len(indices)), indices, 100, where=[i > 100 for i in indices],
                 color=ORANGE, alpha=0.3, label='Выше среднего (сезон)')
ax4.fill_between(range(len(indices)), indices, 100, where=[i <= 100 for i in indices],
                 color='#6B7280', alpha=0.2, label='Ниже среднего (не сезон)')
ax4.plot(range(len(indices)), indices, color=ORANGE, linewidth=2, marker='o', markersize=5)
ax4.axhline(y=100, color=DARK, linestyle='--', linewidth=0.5, alpha=0.5)
ax4.set_xticks(range(len(months_labels)))
ax4.set_xticklabels(months_labels, rotation=45, fontsize=9)
ax4.set_ylabel('Индекс сезонности (100 = средний)', fontsize=11, color=DARK)
ax4.set_title('Сезонность спроса — усреднённо по топ-5 фразам\n(Wordstat, сент 2025 — авг 2026)',
              fontsize=14, fontweight='bold', color=DARK, pad=16)
ax4.legend(loc='upper left', fontsize=9)
fig4.savefig(os.path.join(CHART_DIR, 'seasonality.png'), dpi=200, bbox_inches='tight', facecolor='white')
plt.close(fig4)
print('Chart 4 saved')

# --- График 5: Брендовые vs продуктовые запросы ---
brand_phrases = ['TITAN краска', 'Ореол TITAN', 'Эмпилс', 'краска TITAN по металлу', 'грунт-эмаль TITAN']
brand_total = sum(int(data[p].get('total_count', '0')) for p in brand_phrases)
product_phrases = ['грунт-эмаль', 'грунт-эмаль по ржавчине', 'грунт-эмаль 3 в 1',
                   'краска для металла', 'краска по металлу', 'эмаль по металлу',
                   'быстросохнущая краска по металлу', 'молотковая краска',
                   'краска по ржавчине', 'грунт по металлу']
product_total = sum(int(data[p].get('total_count', '0')) for p in product_phrases)
app_phrases = ['краска для ворот', 'краска для забора', 'краска для гаража',
               'краска для металлической двери', 'покрасить металлический забор', 'чем покрасить металл']
app_total = sum(int(data[p].get('total_count', '0')) for p in app_phrases)
comp_phrases = ['Dali краска по металлу', 'Салван краска', 'Husky краска', 'Dufa краска по металлу']
comp_total = sum(int(data[p].get('total_count', '0')) for p in comp_phrases)

fig5, ax5 = plt.subplots(figsize=(8, 8), constrained_layout=True)
sizes = [brand_total, product_total, app_total, comp_total]
labels5 = [f'Брендовые\n({brand_total:,})'.replace(',', ' '),
           f'Продуктовые\n({product_total:,})'.replace(',', ' '),
           f'Применение\n({app_total:,})'.replace(',', ' '),
           f'Конкурентные\n({comp_total:,})'.replace(',', ' ')]
colors5 = [GREY, ORANGE, '#FBA668', '#F5C896']
wedges, texts, autotexts = ax5.pie(sizes, labels=labels5, autopct='%1.1f%%',
                                    colors=colors5, startangle=90,
                                    pctdistance=0.75, wedgeprops={'edgecolor': 'white', 'linewidth': 2},
                                    textprops={'fontsize': 10, 'color': DARK, 'fontweight': 'bold'})
ax5.set_title('Структура спроса по группам запросов\n(Wordstat, сентябрь 2026)',
              fontsize=14, fontweight='bold', color=DARK, pad=16)
fig5.savefig(os.path.join(CHART_DIR, 'demand_structure.png'), dpi=200, bbox_inches='tight', facecolor='white')
plt.close(fig5)
print('Chart 5 saved')

print(f'\nВсе визуализации сохранены в {CHART_DIR}')
print(f'Excel: {excel_path}')
