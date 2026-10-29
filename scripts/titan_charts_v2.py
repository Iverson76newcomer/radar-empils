"""
Generate visualizations for TITAN sales analysis.
Charts:
1. Monthly sales trend (line chart)
2. Top-10 SKUs horizontal bar chart
3. Color distribution pie chart
4. Color × Volume heatmap
5. Shop performance horizontal bar chart
"""
import csv
import re
import os
from collections import defaultdict
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import numpy as np

# Font
fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf')
plt.rcParams['font.sans-serif'] = ['DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# Palette (pastel)
BG       = '#F5F2EC'
INK      = '#2F3A48'
MUTE     = '#8A95A3'
RED      = '#C00000'
SAND     = '#C9A77C'
GREEN    = '#8BAA8A'
BLUE     = '#6B8FA8'
ORANGE   = '#D4A574'
GRAY     = '#D9D2C2'
DIV      = '#EFE9DD'

# Color palette for charts
CHART_COLORS = ['#C00000', '#C9A77C', '#6B8FA8', '#8BAA8A', '#D4A574',
                '#B5614A', '#5A6776', '#A7C0D0', '#8B7355', '#8A95A3',
                '#D9D2C2', '#C7333E']

SRC = '/home/z/my-project/upload/lemena_sales_clean_20261002.csv'
OUT_DIR = '/home/z/my-project/download/titan_charts_02_10_2026'
os.makedirs(OUT_DIR, exist_ok=True)

# ── Load + parse ───────────────────────────────────────────────────────
COLOR_TOKENS = {
    'БЕЛ': 'БЕЛЫЙ', 'БЕЛЫЙ': 'БЕЛЫЙ', 'БЕЛАЯ': 'БЕЛЫЙ',
    'СЕР': 'СЕРЫЙ', 'СЕРЫЙ': 'СЕРЫЙ', 'СЕРАЯ': 'СЕРЫЙ',
    'ЧЁРН': 'ЧЁРНЫЙ', 'ЧЕРН': 'ЧЁРНЫЙ', 'ЧЁРНЫЙ': 'ЧЁРНЫЙ', 'ЧЕРНЫЙ': 'ЧЁРНЫЙ',
    'СИН': 'СИНИЙ', 'СИНИЙ': 'СИНИЙ', 'СИНЯЯ': 'СИНИЙ',
    'ЗЕЛЁН': 'ЗЕЛЁНЫЙ', 'ЗЕЛЕН': 'ЗЕЛЁНЫЙ', 'ЗЕЛЁНЫЙ': 'ЗЕЛЁНЫЙ', 'ЗЕЛЕНЫЙ': 'ЗЕЛЁНЫЙ',
    'ЖЁЛТ': 'ЖЁЛТЫЙ', 'ЖЕЛТ': 'ЖЁЛТЫЙ', 'ЖЁЛТЫЙ': 'ЖЁЛТЫЙ', 'ЖЕЛТЫЙ': 'ЖЁЛТЫЙ',
    'КРАСН': 'КРАСНЫЙ', 'КРАСНЫЙ': 'КРАСНЫЙ', 'КРАСНАЯ': 'КРАСНЫЙ',
    'ГРАФИТ': 'ГРАФИТ',
    'КР-КОР': 'КРАСНО-КОРИЧНЕВЫЙ',
    'ШОК-КОР': 'ШОКОЛАДНО-КОРИЧНЕВЫЙ',
    'СЛОН.К': 'СЛОНОВАЯ КОСТЬ', 'СЛОН К': 'СЛОНОВАЯ КОСТЬ', 'СЛОНОВАЯ': 'СЛОНОВАЯ КОСТЬ',
    'СЕРЕБР': 'СЕРЕБРИСТЫЙ',
    'RAL9005': 'RAL 9005', 'RAL9010': 'RAL 9010',
    'RAL 9005': 'RAL 9005', 'RAL 9010': 'RAL 9010',
}
SORTED_TOKENS = sorted(COLOR_TOKENS.keys(), key=len, reverse=True)
COLOR_RE = re.compile(r'\b(' + '|'.join(re.escape(t) for t in SORTED_TOKENS) + r')(?:\b|\d)', re.IGNORECASE)
COLOR_NOSEP_RE = re.compile(r'(БЕЛ|СЕР|ЧЁРН|ЧЕРН)(RAL\s*\d{4})', re.IGNORECASE)
VOL_L_RE = re.compile(r'(\d+(?:,\d+)?)\s*Л\b', re.IGNORECASE)
VOL_KG_RE = re.compile(r'(\d+(?:,\d+)?)\s*КГ\b', re.IGNORECASE)

def parse_color_volume(name: str):
    name = name.upper().strip()
    color = 'НЕ ОПРЕДЕЛЕНА'
    volume_str = '—'
    volume_l = None
    m_nosep = COLOR_NOSEP_RE.search(name)
    if m_nosep:
        prefix = m_nosep.group(1).upper()
        ral = m_nosep.group(2).upper().replace(' ', '')
        color = COLOR_TOKENS.get(ral, f'RAL {ral[3:7]}')
    else:
        m = COLOR_RE.search(name)
        if m:
            color = COLOR_TOKENS.get(m.group(1).upper(), m.group(1).upper())
        ral_m = re.search(r'RAL\s*(\d{4})', name, re.IGNORECASE)
        if ral_m and color == 'НЕ ОПРЕДЕЛЕНА':
            color = f'RAL {ral_m.group(1)}'
    m_l = VOL_L_RE.search(name)
    m_kg = VOL_KG_RE.search(name)
    if m_l:
        v = m_l.group(1).replace(',', '.')
        volume_l = float(v)
        volume_str = f'{v} Л'
    elif m_kg:
        v = m_kg.group(1).replace(',', '.')
        volume_l = float(v)
        volume_str = f'{v} КГ'
    return color, volume_str, volume_l

# Load TITAN rows
titan_rows = []
with open(SRC, 'r', encoding='utf-8') as f:
    reader = csv.reader(f, delimiter=';')
    next(reader)
    for row in reader:
        if len(row) >= 5 and row[4] and 'TITAN' in row[4].upper():
            titan_rows.append(row)

# Parse all rows
records = []
for row in titan_rows:
    sku_name = row[4]
    month = row[0]
    shop = row[8]
    try: stock = int(row[12])
    except: stock = 0
    try: in_transit = int(row[13])
    except: in_transit = 0
    try: sales = int(row[14])
    except: sales = 0
    try: price = float(row[16]) if row[16] else 0
    except: price = 0
    color, vol_str, vol_l = parse_color_volume(sku_name)
    records.append({'month': month, 'shop': shop, 'sku': sku_name,
                    'color': color, 'volume_str': vol_str, 'volume_l': vol_l,
                    'stock': stock, 'in_transit': in_transit, 'sales': sales,
                    'price': price})

# Sort months chronologically
MONTHS_ORDER = {
    '10.2025': 1, '11.2025': 2, '12.2025': 3,
    '01.2026': 4, '02.2026': 5, '03.2026': 6,
    '04.2026': 7, '05.2026': 8, '06.2026': 9,
    '07.2026': 10, '08.2026': 11, '09.2026': 12,
    '10.2026': 13,
}
unique_months = sorted(set(r['month'] for r in records),
                       key=lambda m: (MONTHS_ORDER.get(m, 99), m))
print(f"Months: {unique_months}")

# ─── 1. Monthly sales trend ─────────────────────────────────────────
by_month = defaultdict(lambda: {'sales': 0, 'stock': 0, 'revenue': 0.0})
for r in records:
    by_month[r['month']]['sales'] += r['sales']
    by_month[r['month']]['stock'] += r['stock']
    by_month[r['month']]['revenue'] += r['sales'] * r['price']

fig, ax1 = plt.subplots(figsize=(12, 5.5), dpi=160, facecolor=BG)
ax1.set_facecolor(BG)
months = unique_months
sales = [by_month[m]['sales'] for m in months]
revenue = [by_month[m]['revenue']/1000 for m in months]  # in thousands

ax1.fill_between(range(len(months)), sales, alpha=0.2, color=RED, zorder=2)
ax1.plot(range(len(months)), sales, marker='o', markersize=8,
         color=RED, linewidth=2.5, markerfacecolor='white',
         markeredgewidth=2.5, markeredgecolor=RED, zorder=4, label='Продажи (шт.)')
for i, v in enumerate(sales):
    ax1.annotate(f'{v:,}', (i, v), textcoords="offset points",
                 xytext=(0, 12), ha='center', fontsize=9,
                 color=INK, fontweight='bold')
ax1.set_ylabel('Продажи (шт.)', color=INK, fontsize=11, fontweight='bold')
ax1.tick_params(axis='y', labelcolor=INK)
ax1.set_xticks(range(len(months)))
ax1.set_xticklabels(months, rotation=30, ha='right', color=INK, fontsize=10)
ax1.set_xlim(-0.5, len(months) - 0.5)
ax1.set_ylim(0, max(sales) * 1.2)

# Right axis: revenue
ax2 = ax1.twinx()
ax2.bar(range(len(months)), revenue, alpha=0.15, color=SAND,
        width=0.55, zorder=1, label='Выручка (тыс. ₽)')
ax2.set_ylabel('Выручка (тыс. ₽)', color='#8B7355', fontsize=11, fontweight='bold')
ax2.tick_params(axis='y', labelcolor='#8B7355')
ax2.set_ylim(0, max(revenue) * 1.3 if revenue else 100)

# Title
plt.title('Динамика продаж ОРЕОЛ TITAN по месяцам',
          fontsize=14, fontweight='bold', color=INK, pad=15)

# Combined legend
lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper left',
           framealpha=0.95, facecolor='white', edgecolor=GRAY)

# Grid
ax1.grid(axis='y', alpha=0.25, color=MUTE, linestyle='--', linewidth=0.5)
ax1.spines['top'].set_visible(False)
ax2.spines['top'].set_visible(False)
ax1.spines['right'].set_visible(False)
ax1.spines['left'].set_color(MUTE)
ax1.spines['bottom'].set_color(MUTE)

plt.tight_layout()
out = f'{OUT_DIR}/01_monthly_trend.png'
plt.savefig(out, dpi=160, facecolor=BG, bbox_inches='tight')
plt.close()
print(f"Saved: {out}")

# ─── 2. Top 10 SKUs ─────────────────────────────────────────────────
by_sku = defaultdict(lambda: {'sales': 0, 'color': '', 'volume_str': ''})
for r in records:
    by_sku[r['sku']]['sales'] += r['sales']
    by_sku[r['sku']]['color'] = r['color']
    by_sku[r['sku']]['volume_str'] = r['volume_str']

top10 = sorted(by_sku.items(), key=lambda x: -x[1]['sales'])[:10]
top10.reverse()  # for horizontal bar (bottom-up)

fig, ax = plt.subplots(figsize=(12, 6.5), dpi=160, facecolor=BG)
ax.set_facecolor(BG)

skus = [s[:50] for s, _ in top10]
sales10 = [d['sales'] for _, d in top10]
colors10 = [d['color'] for _, d in top10]

# Color the bars by their actual color
color_map = {
    'ГРАФИТ': '#5A6776', 'ШОКОЛАДНО-КОРИЧНЕВЫЙ': '#6B4423',
    'СЕРЫЙ': '#A8A8A8', 'БЕЛЫЙ': '#F0F0F0', 'ЧЁРНЫЙ': '#2F3A48',
    'СИНИЙ': '#4A6FA5', 'ЗЕЛЁНЫЙ': '#5E8B5E', 'ЖЁЛТЫЙ': '#D4A82C',
    'КРАСНО-КОРИЧНЕВЫЙ': '#8B4513', 'СЛОНОВАЯ КОСТЬ': '#F5DEB3',
    'СЕРЕБРИСТЫЙ': '#C0C0C0', 'RAL 9005': '#1A1A1A', 'RAL 9010': '#F8F8F0',
}
bar_colors = [color_map.get(c, RED) for c in colors10]
# For visibility — if a bar color is too pale, use border
bars = ax.barh(range(len(top10)), sales10, color=bar_colors,
               edgecolor=INK, linewidth=0.8, height=0.7)

# Value labels
for i, (bar, val) in enumerate(zip(bars, sales10)):
    ax.text(val + max(sales10) * 0.01, i, f'{val:,}',
            va='center', ha='left', fontsize=10, color=INK, fontweight='bold')

ax.set_yticks(range(len(top10)))
ax.set_yticklabels(skus, fontsize=9, color=INK)
ax.set_xlabel('Продажи (шт.)', fontsize=11, color=INK, fontweight='bold')
ax.set_title('Топ-10 SKU по продажам (октябрь 2025 — сентябрь 2026)',
             fontsize=14, fontweight='bold', color=INK, pad=15)
ax.set_xlim(0, max(sales10) * 1.15)
ax.grid(axis='x', alpha=0.25, color=MUTE, linestyle='--', linewidth=0.5)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['left'].set_color(MUTE)
ax.spines['bottom'].set_color(MUTE)

plt.tight_layout()
out = f'{OUT_DIR}/02_top10_sku.png'
plt.savefig(out, dpi=160, facecolor=BG, bbox_inches='tight')
plt.close()
print(f"Saved: {out}")

# ─── 3. Color distribution (pie + donut) ────────────────────────────
by_color = defaultdict(int)
for r in records:
    by_color[r['color']] += r['sales']

sorted_colors = sorted(by_color.items(), key=lambda x: -x[1])
top_colors = sorted_colors[:6]
other_sales = sum(v for _, v in sorted_colors[6:])
if other_sales > 0:
    top_colors.append(('Прочие', other_sales))

labels = [c for c, _ in top_colors]
sizes = [v for _, v in top_colors]
colors_pie = [color_map.get(c, SAND) for c in labels]

fig, ax = plt.subplots(figsize=(9, 7), dpi=160, facecolor=BG)
ax.set_facecolor(BG)

# Donut chart
wedges, texts, autotexts = ax.pie(sizes, labels=labels, colors=colors_pie,
                                    autopct='%1.1f%%', startangle=90,
                                    wedgeprops=dict(width=0.45, edgecolor='white', linewidth=2),
                                    textprops={'fontsize': 10, 'color': INK},
                                    pctdistance=0.78)
for autotext in autotexts:
    autotext.set_color('white')
    autotext.set_fontweight('bold')
    autotext.set_fontsize(10)

# Center text
total_sales = sum(sizes)
ax.text(0, 0.05, f'{total_sales:,}', ha='center', va='center',
        fontsize=22, fontweight='bold', color=INK)
ax.text(0, -0.12, 'шт. всего', ha='center', va='center',
        fontsize=10, color=MUTE)

plt.title('Структура продаж TITAN по цветам',
          fontsize=14, fontweight='bold', color=INK, pad=20)
plt.tight_layout()
out = f'{OUT_DIR}/03_color_donut.png'
plt.savefig(out, dpi=160, facecolor=BG, bbox_inches='tight')
plt.close()
print(f"Saved: {out}")

# ─── 4. Color × Volume heatmap ─────────────────────────────────────
by_volume = defaultdict(int)
for r in records:
    by_volume[r['volume_str']] += r['sales']

# Sort volumes
VOL_ORDER = ['0.4 Л', '0.75 Л', '2 Л', '5 Л', '8 Л', '10 КГ']
all_volumes = [v for v in VOL_ORDER if v in by_volume]
all_colors_sorted = [c for c, _ in sorted_colors if c != 'Прочие']

matrix = np.zeros((len(all_colors_sorted), len(all_volumes)))
for i, color in enumerate(all_colors_sorted):
    for j, vol in enumerate(all_volumes):
        s = sum(r['sales'] for r in records if r['color'] == color and r['volume_str'] == vol)
        matrix[i, j] = s

fig, ax = plt.subplots(figsize=(11, 7), dpi=160, facecolor=BG)
ax.set_facecolor(BG)

# Use sequential colormap
im = ax.imshow(matrix, cmap='YlOrBr', aspect='auto')
# Add text annotations
for i in range(len(all_colors_sorted)):
    for j in range(len(all_volumes)):
        val = int(matrix[i, j])
        if val > 0:
            color = 'white' if val > matrix.max() * 0.5 else INK
            ax.text(j, i, f'{val:,}', ha='center', va='center',
                    fontsize=10, color=color, fontweight='bold')

ax.set_xticks(range(len(all_volumes)))
ax.set_xticklabels(all_volumes, fontsize=11, color=INK)
ax.set_yticks(range(len(all_colors_sorted)))
ax.set_yticklabels(all_colors_sorted, fontsize=10, color=INK)

# Color bar
cbar = plt.colorbar(im, ax=ax, fraction=0.025, pad=0.02)
cbar.set_label('Продажи (шт.)', fontsize=10, color=INK, rotation=270, labelpad=15)
cbar.ax.tick_params(labelcolor=INK)

plt.title('Матрица продаж: Цвет × Объём',
          fontsize=14, fontweight='bold', color=INK, pad=15)
plt.xlabel('Объём фасовки', fontsize=11, color=INK, fontweight='bold', labelpad=10)
plt.ylabel('Цвет', fontsize=11, color=INK, fontweight='bold', labelpad=10)
plt.tight_layout()
out = f'{OUT_DIR}/04_color_volume_heatmap.png'
plt.savefig(out, dpi=160, facecolor=BG, bbox_inches='tight')
plt.close()
print(f"Saved: {out}")

# ─── 5. Shop performance ───────────────────────────────────────────
by_shop = defaultdict(lambda: {'sales': 0, 'revenue': 0.0, 'sku_count': set()})
for r in records:
    by_shop[r['shop']]['sales'] += r['sales']
    by_shop[r['shop']]['revenue'] += r['sales'] * r['price']
    by_shop[r['shop']]['sku_count'].add(r['sku'])

sorted_shops = sorted(by_shop.items(), key=lambda x: -x[1]['sales'])
shops = [s for s, _ in sorted_shops]
shop_sales = [d['sales'] for _, d in sorted_shops]
shop_rev = [d['revenue']/1000 for _, d in sorted_shops]
shop_sku = [len(d['sku_count']) for _, d in sorted_shops]

fig, ax1 = plt.subplots(figsize=(12, 6), dpi=160, facecolor=BG)
ax1.set_facecolor(BG)

x = np.arange(len(shops))
width = 0.35
bars1 = ax1.bar(x - width/2, shop_sales, width, label='Продажи (шт.)',
                color=RED, edgecolor='white', linewidth=1.2)
ax1.set_ylabel('Продажи (шт.)', color=RED, fontsize=11, fontweight='bold')
ax1.tick_params(axis='y', labelcolor=RED)

ax2 = ax1.twinx()
bars2 = ax2.bar(x + width/2, shop_rev, width, label='Выручка (тыс. ₽)',
                color=SAND, edgecolor='white', linewidth=1.2)
ax2.set_ylabel('Выручка (тыс. ₽)', color='#8B7355', fontsize=11, fontweight='bold')
ax2.tick_params(axis='y', labelcolor='#8B7355')

ax1.set_xticks(x)
ax1.set_xticklabels(shops, rotation=30, ha='right', fontsize=10, color=INK)
ax1.set_title('Продажи и выручка TITAN по магазинам',
              fontsize=14, fontweight='bold', color=INK, pad=15)

# Value labels on bars
for bar, val in zip(bars1, shop_sales):
    ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + max(shop_sales)*0.02,
             f'{val:,}', ha='center', va='bottom', fontsize=8.5, color=RED, fontweight='bold')

# SKU count as text annotations
for i, n in enumerate(shop_sku):
    ax1.text(i, -max(shop_sales)*0.08, f'SKU: {n}', ha='center', va='top',
             fontsize=8, color=MUTE, style='italic')

# Combined legend
lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper right',
           framealpha=0.95, facecolor='white', edgecolor=GRAY)

ax1.spines['top'].set_visible(False)
ax2.spines['top'].set_visible(False)
ax1.grid(axis='y', alpha=0.2, color=MUTE, linestyle='--', linewidth=0.5)
ax1.set_ylim(0, max(shop_sales) * 1.15)
ax2.set_ylim(0, max(shop_rev) * 1.15 if shop_rev else 100)

plt.tight_layout()
out = f'{OUT_DIR}/05_shop_performance.png'
plt.savefig(out, dpi=160, facecolor=BG, bbox_inches='tight')
plt.close()
print(f"Saved: {out}")

# ─── 6. Volume distribution (mini-chart) ───────────────────────────
sorted_volumes = sorted(by_volume.items(), key=lambda x: -x[1])
v_labels = [v for v, _ in sorted_volumes]
v_sizes = [v for _, v in sorted_volumes]

fig, ax = plt.subplots(figsize=(10, 5), dpi=160, facecolor=BG)
ax.set_facecolor(BG)

bars = ax.bar(range(len(v_labels)), v_sizes,
              color=[SAND, RED, BLUE, GREEN, ORANGE, MUTE][:len(v_labels)],
              edgecolor='white', linewidth=1.5)

for bar, val in zip(bars, v_sizes):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + max(v_sizes)*0.02,
            f'{val:,}\n({val/sum(v_sizes)*100:.1f}%)',
            ha='center', va='bottom', fontsize=10, color=INK, fontweight='bold')

ax.set_xticks(range(len(v_labels)))
ax.set_xticklabels(v_labels, fontsize=11, color=INK)
ax.set_ylabel('Продажи (шт.)', fontsize=11, color=INK, fontweight='bold')
ax.set_title('Распределение продаж TITAN по объёмам фасовки',
             fontsize=14, fontweight='bold', color=INK, pad=15)
ax.set_ylim(0, max(v_sizes) * 1.2)
ax.grid(axis='y', alpha=0.25, color=MUTE, linestyle='--', linewidth=0.5)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['left'].set_color(MUTE)
ax.spines['bottom'].set_color(MUTE)

plt.tight_layout()
out = f'{OUT_DIR}/06_volume_distribution.png'
plt.savefig(out, dpi=160, facecolor=BG, bbox_inches='tight')
plt.close()
print(f"Saved: {out}")

print(f"\nAll charts in: {OUT_DIR}")
print(os.listdir(OUT_DIR))
