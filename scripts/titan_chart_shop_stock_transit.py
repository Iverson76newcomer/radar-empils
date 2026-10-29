"""
Generate a shop performance chart for TITAN sales 2026-10-02 data.

Style: replicate the user's reference image (cream bg, dark-red sales bars,
beige second series, grouped columns, SKU sub-labels, comma-formatted data labels).

Spec change: second series is "Текущий остаток (сток + в пути)" in шт.
Single Y-axis (both series in units).
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

# Palette (matches reference image)
BG       = '#F5F2EC'
INK      = '#2F3A48'
MUTE     = '#8A95A3'
RED      = '#C00000'
SAND     = '#C9A77C'
GRAY     = '#D9D2C2'
DIV      = '#EFE9DD'
GRID     = '#FDFCFB'  # near-white gridlines over cream bg

SRC = '/home/z/my-project/upload/lemena_sales_clean_20261002.csv'
OUT = '/home/z/my-project/download/titan_charts_02_10_2026/05_shop_performance_stock_transit.png'

# ── Load TITAN rows ────────────────────────────────────────────────────
titan_rows = []
with open(SRC, 'r', encoding='utf-8') as f:
    reader = csv.reader(f, delimiter=';')
    next(reader)
    for row in reader:
        if len(row) >= 5 and row[4] and 'TITAN' in row[4].upper():
            titan_rows.append(row)

# Aggregate per shop: total sales + current stock + in_transit + SKU count
by_shop = defaultdict(lambda: {
    'sales': 0,
    'stock': 0,
    'in_transit': 0,
    'sku_count': set(),
})

for row in titan_rows:
    shop = row[8]
    try: s = int(row[14]) if row[14].lstrip('-').isdigit() else 0
    except: s = 0
    try: stk = int(row[12]) if row[12].lstrip('-').isdigit() else 0
    except: stk = 0
    try: it = int(row[13]) if row[13].lstrip('-').isdigit() else 0
    except: it = 0
    by_shop[shop]['sales'] += s
    by_shop[shop]['stock'] += stk
    by_shop[shop]['in_transit'] += it
    by_shop[shop]['sku_count'].add(row[4])

# Sort shops by sales descending
sorted_shops = sorted(by_shop.items(), key=lambda x: -x[1]['sales'])

shops = [s for s, _ in sorted_shops]
sales = [d['sales'] for _, d in sorted_shops]
stock_with_transit = [d['stock'] + d['in_transit'] for _, d in sorted_shops]
sku_counts = [len(d['sku_count']) for _, d in sorted_shops]

print("Shop data (sorted by sales):")
print(f"{'Shop':25s} {'Sales':>7s} {'Stock':>7s} {'InTr':>6s} {'Stock+Tr':>9s} {'SKU':>4s}")
for s, d in sorted_shops:
    swt = d['stock'] + d['in_transit']
    print(f"{s:25s} {d['sales']:7,} {d['stock']:7,} {d['in_transit']:6,} {swt:9,} {len(d['sku_count']):4d}")

# ── Build chart ───────────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(13, 6.5), dpi=160, facecolor=BG)
ax.set_facecolor(BG)

x = np.arange(len(shops))
width = 0.38

# Bars: Sales (red) + Stock-with-transit (beige)
bars1 = ax.bar(x - width/2, sales, width,
               color=RED, edgecolor='white', linewidth=1.2,
               label='Продажи (шт.)')
bars2 = ax.bar(x + width/2, stock_with_transit, width,
               color=SAND, edgecolor='white', linewidth=1.2,
               label='Текущий остаток с учётом товара в пути (шт.)')

# Data labels on red bars (sales)
max_val = max(max(sales), max(stock_with_transit))
for bar, val in zip(bars1, sales):
    ax.text(bar.get_x() + bar.get_width()/2,
            bar.get_height() + max_val*0.015,
            f'{val:,}',
            ha='center', va='bottom',
            fontsize=9, color=RED, fontweight='bold')

# Data labels on beige bars (stock + transit)
for bar, val in zip(bars2, stock_with_transit):
    ax.text(bar.get_x() + bar.get_width()/2,
            bar.get_height() + max_val*0.015,
            f'{val:,}',
            ha='center', va='bottom',
            fontsize=9, color='#8B7355', fontweight='bold')

# X-axis: shop names + SKU sub-labels
ax.set_xticks(x)
# Shop name on first line, SKU count on second
ax.set_xticklabels(shops, rotation=30, ha='right', fontsize=10, color=INK)

# Add SKU sub-labels below x-tick labels
# We need to manually position them — use fig.text annotations
for i, n in enumerate(sku_counts):
    ax.text(i, -max_val*0.10, f'SKU: {n}',
            ha='center', va='top',
            fontsize=8, color=MUTE, style='italic',
            transform=ax.transData)

# Y-axis (single, both series in шт.)
ax.set_ylabel('Штуки (шт.)', fontsize=11, color=INK, fontweight='bold')
ax.set_ylim(0, max_val * 1.18)
ax.tick_params(axis='y', labelcolor=INK)
ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{int(x):,}'))

# Title
ax.set_title('Продажи и текущий остаток TITAN по магазинам\n(остаток = сток + товар в пути)',
             fontsize=14, fontweight='bold', color=INK, pad=15)

# Legend (top-right, white box with thin border, like the reference)
legend = ax.legend(loc='upper right',
                   framealpha=0.95, facecolor='white',
                   edgecolor='#8B7355')
legend.get_frame().set_linewidth(0.8)
legend.get_frame().set_edgecolor('#8B7355')

# Gridlines (horizontal, near-white)
ax.grid(axis='y', alpha=0.6, color=GRID, linestyle='-', linewidth=0.8, zorder=0)
ax.set_axisbelow(True)

# Spines
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['left'].set_color(MUTE)
ax.spines['bottom'].set_color(MUTE)

# Adjust layout to leave room for SKU sub-labels
plt.tight_layout()
plt.subplots_adjust(bottom=0.18)

plt.savefig(OUT, dpi=160, facecolor=BG, bbox_inches='tight')
plt.close()
print(f"\nSaved: {OUT} ({os.path.getsize(OUT)} bytes)")
