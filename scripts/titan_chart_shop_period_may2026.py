"""
Shop performance chart for TITAN — period restricted to 05.2026 - 10.2026.
Sales = sum over period
Stock + in_transit = latest snapshot per SKU×shop (deduped)

Same style as 05_shop_performance_stock_transit.png but with period-filtered numbers.
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

fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf')
plt.rcParams['font.sans-serif'] = ['DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# Palette
BG       = '#F5F2EC'
INK      = '#2F3A48'
MUTE     = '#8A95A3'
RED      = '#C00000'
SAND     = '#C9A77C'
GRID     = '#FDFCFB'

SRC = '/home/z/my-project/upload/lemena_sales_clean_20261002.csv'
OUT = '/home/z/my-project/download/titan_charts_02_10_2026/05_shop_perf_period_may_oct_2026.png'

# Period filter
def in_period(month_str: str) -> bool:
    try:
        m, y = month_str.split('.')
        m, y = int(m), int(y)
        if y > 2026: return True
        if y == 2026 and m >= 5: return True
        return False
    except:
        return False

# ── Load TITAN rows in period ─────────────────────────────────────────
titan_rows = []
with open(SRC, 'r', encoding='utf-8') as f:
    reader = csv.reader(f, delimiter=';')
    next(reader)
    for row in reader:
        if len(row) >= 5 and row[4] and 'TITAN' in row[4].upper():
            if in_period(row[0]):
                titan_rows.append(row)

# Aggregate per shop: sales (sum), stock+transit (deduped max per SKU×shop), SKU count
by_shop = defaultdict(lambda: {
    'sales': 0,
    'stock': 0,
    'in_transit': 0,
    'sku_count': set(),
    'stock_dedup': {},  # (sku, shop) -> stock
    'in_transit_dedup': {},
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
    by_shop[shop]['sku_count'].add(row[4])
    key = row[4]  # SKU name (shop is implicit in dict key)
    if key not in by_shop[shop]['stock_dedup'] or stk > by_shop[shop]['stock_dedup'][key]:
        by_shop[shop]['stock_dedup'][key] = stk
    if key not in by_shop[shop]['in_transit_dedup'] or it > by_shop[shop]['in_transit_dedup'][key]:
        by_shop[shop]['in_transit_dedup'][key] = it

# Finalize: sum deduped values
for shop in by_shop:
    by_shop[shop]['stock'] = sum(by_shop[shop]['stock_dedup'].values())
    by_shop[shop]['in_transit'] = sum(by_shop[shop]['in_transit_dedup'].values())

# Sort by sales desc
sorted_shops = sorted(by_shop.items(), key=lambda x: -x[1]['sales'])
shops = [s for s, _ in sorted_shops]
sales = [d['sales'] for _, d in sorted_shops]
stock_with_transit = [d['stock'] + d['in_transit'] for _, d in sorted_shops]
sku_counts = [len(d['sku_count']) for _, d in sorted_shops]

print("Period: 05.2026 — 10.2026 (5 months)")
print(f"\n{'Shop':25s} {'Sales':>7s} {'Stock':>7s} {'InTr':>6s} {'Stock+Tr':>9s} {'SKU':>4s}")
for s, d in sorted_shops:
    swt = d['stock'] + d['in_transit']
    print(f"{s:25s} {d['sales']:7,} {d['stock']:7,} {d['in_transit']:6,} {swt:9,} {len(d['sku_count']):4d}")

# ── Build chart ──────────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(13, 6.5), dpi=160, facecolor=BG)
ax.set_facecolor(BG)

x = np.arange(len(shops))
width = 0.38

bars1 = ax.bar(x - width/2, sales, width,
               color=RED, edgecolor='white', linewidth=1.2,
               label='Продажи за 5 мес. (шт.)')
bars2 = ax.bar(x + width/2, stock_with_transit, width,
               color=SAND, edgecolor='white', linewidth=1.2,
               label='Текущий остаток + в пути (шт.)')

max_val = max(max(sales), max(stock_with_transit))
for bar, val in zip(bars1, sales):
    ax.text(bar.get_x() + bar.get_width()/2,
            bar.get_height() + max_val*0.015,
            f'{val:,}',
            ha='center', va='bottom',
            fontsize=9, color=RED, fontweight='bold')
for bar, val in zip(bars2, stock_with_transit):
    ax.text(bar.get_x() + bar.get_width()/2,
            bar.get_height() + max_val*0.015,
            f'{val:,}',
            ha='center', va='bottom',
            fontsize=9, color='#8B7355', fontweight='bold')

ax.set_xticks(x)
ax.set_xticklabels(shops, rotation=30, ha='right', fontsize=10, color=INK)

# SKU sub-labels below x-axis
for i, n in enumerate(sku_counts):
    ax.text(i, -max_val*0.10, f'SKU: {n}',
            ha='center', va='top',
            fontsize=8, color=MUTE, style='italic',
            transform=ax.transData)

# Y-axis (single, both in шт.)
ax.set_ylabel('Штуки (шт.)', fontsize=11, color=INK, fontweight='bold')
ax.set_ylim(0, max_val * 1.18)
ax.tick_params(axis='y', labelcolor=INK)
ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{int(x):,}'))

# Title
ax.set_title('Продажи (5 мес.) и текущий остаток TITAN по магазинам\n'
             'Период: май — октябрь 2026 · Остаток = сток + товар в пути',
             fontsize=14, fontweight='bold', color=INK, pad=15)

# Legend
legend = ax.legend(loc='upper right',
                   framealpha=0.95, facecolor='white',
                   edgecolor='#8B7355')
legend.get_frame().set_linewidth(0.8)
legend.get_frame().set_edgecolor('#8B7355')

# Gridlines
ax.grid(axis='y', alpha=0.6, color=GRID, linestyle='-', linewidth=0.8, zorder=0)
ax.set_axisbelow(True)

ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['left'].set_color(MUTE)
ax.spines['bottom'].set_color(MUTE)

plt.tight_layout()
plt.subplots_adjust(bottom=0.18)

plt.savefig(OUT, dpi=160, facecolor=BG, bbox_inches='tight')
plt.close()
print(f"\nSaved: {OUT} ({os.path.getsize(OUT)} bytes)")
