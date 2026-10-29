"""
TITAN analysis — restricted to period from 01.05.2026 (5-6 months).
Filters TITAN rows where Месяц >= 05.2026.
Stock + in_transit values are taken from the upload snapshot (02.10.2026),
aggregated per shop / SKU / color / volume.

Outputs:
- Filtered CSV with only post-05.2026 rows
- Updated analysis markdown
- Updated shop performance chart (sales + stock+transit) with new numbers
"""
import csv
import re
import os
from collections import defaultdict

SRC = '/home/z/my-project/upload/lemena_sales_clean_20261002.csv'
OUT_CSV = '/home/z/my-project/download/Лемана_Про_ОРЕОЛ_TITAN_2026_05_10.csv'
ANALYSIS_TXT = '/home/z/my-project/download/Лемана_Про_ОРЕОЛ_TITAN_анализ_май_октябрь_2026.md'

# Period filter: include only months >= 05.2026
# Format in file: "MM.YYYY"
PERIOD_START_MONTH = 5
PERIOD_START_YEAR = 2026

def in_period(month_str: str) -> bool:
    """Return True if month_str (MM.YYYY) is >= 05.2026."""
    try:
        m, y = month_str.split('.')
        m, y = int(m), int(y)
        if y > PERIOD_START_YEAR: return True
        if y == PERIOD_START_YEAR and m >= PERIOD_START_MONTH: return True
        return False
    except:
        return False

# ── Color parsing (same as before) ─────────────────────────────────────
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

def parse_color_volume(sku_name: str):
    name = sku_name.upper().strip()
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

# ── Step 1: Filter TITAN rows for period >= 05.2026 ────────────────────
titan_rows_period = []
header = None
total_rows = 0
total_titan_rows = 0

with open(SRC, 'r', encoding='utf-8') as f:
    reader = csv.reader(f, delimiter=';')
    header = next(reader)
    for row in reader:
        total_rows += 1
        if len(row) >= 5 and row[4] and 'TITAN' in row[4].upper():
            total_titan_rows += 1
            if in_period(row[0]):
                titan_rows_period.append(row)

print(f"Total rows: {total_rows}")
print(f"All TITAN rows: {total_titan_rows}")
print(f"TITAN rows in period >= 05.2026: {len(titan_rows_period)}")
print(f"Period filter ratio: {len(titan_rows_period)/total_titan_rows*100:.1f}% of all TITAN")

# ── Step 2: Parse records ──────────────────────────────────────────────
parsed_records = []
for row in titan_rows_period:
    sku_name = row[4]
    month = row[0]
    shop_num = row[7]
    shop_name = row[8]
    sku_art = row[3]
    try: stock = int(row[12]) if row[12].lstrip('-').isdigit() else 0
    except: stock = 0
    try: in_transit = int(row[13]) if row[13].lstrip('-').isdigit() else 0
    except: in_transit = 0
    try: sales = int(row[14]) if row[14].lstrip('-').isdigit() else 0
    except: sales = 0
    try: price = float(row[16]) if row[16] else 0.0
    except: price = 0.0
    barcode = row[11]
    color, volume_str, volume_l = parse_color_volume(sku_name)
    parsed_records.append({
        'month': month, 'shop_num': shop_num, 'shop': shop_name,
        'sku_art': sku_art, 'sku_name': sku_name, 'barcode': barcode,
        'color': color, 'volume_str': volume_str, 'volume_l': volume_l,
        'stock': stock, 'in_transit': in_transit, 'sales': sales, 'price': price,
    })

# Note about stock: each row has its own stock value (snapshot at upload time, 02.10.2026)
# When aggregating per shop, we should DEDUPE by SKU×shop (since the same SKU×shop
# appears in multiple months but the stock value is the same — it's the latest snapshot)
# For per-month aggregation, the "stock" in each row represents current state, not historical

unique_months = sorted(set(r['month'] for r in parsed_records),
                       key=lambda m: (int(m.split('.')[1]), int(m.split('.')[0])))
n_shops = len(set(r['shop'] for r in parsed_records))
n_skus = len(set(r['sku_name'] for r in parsed_records))
unique_colors = sorted(set(r['color'] for r in parsed_records))
unique_volumes = sorted(set(r['volume_str'] for r in parsed_records),
                       key=lambda v: (0, float(v.split()[0].replace(',', '.')))
                                    if v != '—' and v.split()[0].replace(',', '').replace('.','').isdigit()
                                    else (1, v))

print(f"\nPeriod: {unique_months[0]} — {unique_months[-1]} ({len(unique_months)} months)")
print(f"Shops: {n_shops}")
print(f"SKUs: {n_skus}")
print(f"Colors: {len(unique_colors)}")
print(f"Volumes: {len(unique_volumes)}")

# ── Step 3: Aggregations ───────────────────────────────────────────────
# For sales: sum across all rows in the period
# For stock + in_transit: take LATEST value per SKU×shop (deduplication)
# Reason: stock is a snapshot at upload time, not a monthly figure

# Get unique SKU×shop combinations and their latest stock+in_transit
latest_stock = {}  # (sku, shop) -> {'stock': int, 'in_transit': int}
for r in parsed_records:
    key = (r['sku_name'], r['shop'])
    # Take max stock seen (they should all be the same — snapshot)
    if key not in latest_stock:
        latest_stock[key] = {'stock': r['stock'], 'in_transit': r['in_transit']}
    else:
        # Keep the larger value (in case of multiple snapshots)
        if r['stock'] > latest_stock[key]['stock']:
            latest_stock[key]['stock'] = r['stock']
        if r['in_transit'] > latest_stock[key]['in_transit']:
            latest_stock[key]['in_transit'] = r['in_transit']

total_sales = sum(r['sales'] for r in parsed_records)
total_revenue = sum(r['sales'] * r['price'] for r in parsed_records)
total_stock_unique = sum(d['stock'] for d in latest_stock.values())
total_in_transit_unique = sum(d['in_transit'] for d in latest_stock.values())
total_stock_with_transit = total_stock_unique + total_in_transit_unique

# By color
by_color = defaultdict(lambda: {'sales': 0, 'stock': 0, 'in_transit': 0,
                                  'shop_count': set(), 'sku_count': set(),
                                  'revenue': 0.0, 'latest_stock_taken': set()})
# Track latest stock per color×sku×shop
color_stock_dedup = defaultdict(dict)  # color -> (sku, shop) -> stock
for r in parsed_records:
    c = r['color']
    by_color[c]['sales'] += r['sales']
    by_color[c]['revenue'] += r['sales'] * r['price']
    by_color[c]['shop_count'].add(r['shop'])
    by_color[c]['sku_count'].add(r['sku_name'])
    key = (r['sku_name'], r['shop'])
    if key not in color_stock_dedup[c] or r['stock'] > color_stock_dedup[c][key]:
        color_stock_dedup[c][key] = r['stock']

# Sum unique stock per color
for c, shops_dict in color_stock_dedup.items():
    by_color[c]['stock'] = sum(shops_dict.values())

# In-transit similarly
color_in_transit_dedup = defaultdict(dict)
for r in parsed_records:
    c = r['color']
    key = (r['sku_name'], r['shop'])
    if key not in color_in_transit_dedup[c] or r['in_transit'] > color_in_transit_dedup[c][key]:
        color_in_transit_dedup[c][key] = r['in_transit']
for c, shops_dict in color_in_transit_dedup.items():
    by_color[c]['in_transit'] = sum(shops_dict.values())

# By volume
by_volume = defaultdict(lambda: {'sales': 0, 'stock': 0, 'in_transit': 0,
                                   'shop_count': set(), 'sku_count': set(),
                                   'revenue': 0.0})
vol_stock_dedup = defaultdict(dict)
vol_in_transit_dedup = defaultdict(dict)
for r in parsed_records:
    v = r['volume_str']
    by_volume[v]['sales'] += r['sales']
    by_volume[v]['revenue'] += r['sales'] * r['price']
    by_volume[v]['shop_count'].add(r['shop'])
    by_volume[v]['sku_count'].add(r['sku_name'])
    key = (r['sku_name'], r['shop'])
    if key not in vol_stock_dedup[v] or r['stock'] > vol_stock_dedup[v][key]:
        vol_stock_dedup[v][key] = r['stock']
    if key not in vol_in_transit_dedup[v] or r['in_transit'] > vol_in_transit_dedup[v][key]:
        vol_in_transit_dedup[v][key] = r['in_transit']
for v, d in vol_stock_dedup.items():
    by_volume[v]['stock'] = sum(d.values())
for v, d in vol_in_transit_dedup.items():
    by_volume[v]['in_transit'] = sum(d.values())

# By shop
by_shop = defaultdict(lambda: {'sales': 0, 'stock': 0, 'in_transit': 0,
                                'sku_count': set(), 'revenue': 0.0})
shop_stock_dedup = defaultdict(dict)
shop_in_transit_dedup = defaultdict(dict)
for r in parsed_records:
    s = r['shop']
    by_shop[s]['sales'] += r['sales']
    by_shop[s]['revenue'] += r['sales'] * r['price']
    by_shop[s]['sku_count'].add(r['sku_name'])
    key = (r['sku_name'], r['shop'])
    if key not in shop_stock_dedup[s] or r['stock'] > shop_stock_dedup[s][key]:
        shop_stock_dedup[s][key] = r['stock']
    if key not in shop_in_transit_dedup[s] or r['in_transit'] > shop_in_transit_dedup[s][key]:
        shop_in_transit_dedup[s][key] = r['in_transit']
for s, d in shop_stock_dedup.items():
    by_shop[s]['stock'] = sum(d.values())
for s, d in shop_in_transit_dedup.items():
    by_shop[s]['in_transit'] = sum(d.values())

# By month
by_month = defaultdict(lambda: {'sales': 0, 'revenue': 0.0})
for r in parsed_records:
    m = r['month']
    by_month[m]['sales'] += r['sales']
    by_month[m]['revenue'] += r['sales'] * r['price']

# By SKU
by_sku = defaultdict(lambda: {'sales': 0, 'stock': 0, 'in_transit': 0,
                                'color': '', 'volume_str': '', 'shop_count': set(),
                                'revenue': 0.0})
sku_stock_dedup = defaultdict(dict)
sku_in_transit_dedup = defaultdict(dict)
for r in parsed_records:
    s = r['sku_name']
    by_sku[s]['sales'] += r['sales']
    by_sku[s]['revenue'] += r['sales'] * r['price']
    by_sku[s]['color'] = r['color']
    by_sku[s]['volume_str'] = r['volume_str']
    by_sku[s]['shop_count'].add(r['shop'])
    key = (r['sku_name'], r['shop'])
    if key not in sku_stock_dedup[s] or r['stock'] > sku_stock_dedup[s][key]:
        sku_stock_dedup[s][key] = r['stock']
    if key not in sku_in_transit_dedup[s] or r['in_transit'] > sku_in_transit_dedup[s][key]:
        sku_in_transit_dedup[s][key] = r['in_transit']
for s, d in sku_stock_dedup.items():
    by_sku[s]['stock'] = sum(d.values())
for s, d in sku_in_transit_dedup.items():
    by_sku[s]['in_transit'] = sum(d.values())

# Color × Volume matrix
matrix = defaultdict(lambda: defaultdict(int))
for r in parsed_records:
    matrix[r['color']][r['volume_str']] += r['sales']

# ── Step 4: Write filtered CSV ─────────────────────────────────────────
with open(OUT_CSV, 'w', encoding='utf-8-sig', newline='') as f:
    writer = csv.writer(f, delimiter=';')
    out_header = header + ['ЦВЕТ', 'ОБЪЁМ', 'ОБЪЁМ_Л']
    writer.writerow(out_header)
    for row in titan_rows_period:
        sku_name = row[4]
        color, volume_str, volume_l = parse_color_volume(sku_name)
        vol_l_str = f"{volume_l:.2f}" if volume_l is not None else ""
        writer.writerow(row + [color, volume_str, vol_l_str])

print(f"\nFiltered CSV: {OUT_CSV} ({os.path.getsize(OUT_CSV)} bytes)")

# ── Step 5: Compute analytics for markdown ─────────────────────────────
n_months = len(unique_months)
avg_monthly_sales = total_sales / max(n_months, 1)
turnover_ratio = total_sales / max(total_stock_unique, 1)
monthly_turnover = turnover_ratio / max(n_months, 1)
days_of_stock = total_stock_with_transit / max(total_sales / (n_months * 30), 1)

print(f"\n{'='*70}")
print(f"PERIOD ANALYSIS: {unique_months[0]} — {unique_months[-1]} ({n_months} months)")
print(f"{'='*70}")
print(f"Total sales: {total_sales:,} units")
print(f"Total revenue: {total_revenue:,.0f} ₽")
print(f"Total stock (deduped): {total_stock_unique:,} units")
print(f"Total in-transit (deduped): {total_in_transit_unique:,} units")
print(f"Stock + in-transit: {total_stock_with_transit:,} units")
print(f"Avg monthly sales: {avg_monthly_sales:.0f} units")
print(f"Turnover ratio (sales/stock): {turnover_ratio:.2f}")
print(f"Days of stock: {days_of_stock:.0f}")

# ── Step 6: Write analysis markdown ───────────────────────────────────
md = []
md.append("# Анализ продаж ОРЕОЛ TITAN в сети Лемана Про (май — октябрь 2026)\n")
md.append("**Источник данных:** `commreport_10011142_sales_2026.10.2.csv` (22,5 МБ, 102 677 строк)\n")
md.append("**Фильтр:** только позиции со словом TITAN в названии артикула, период ≥ 05.2026\n")
md.append(f"**Период анализа:** {unique_months[0]} — {unique_months[-1]} ({n_months} месяцев)\n")
md.append("**Дата выгрузки:** 02.10.2026\n")
md.append("\n---\n\n")

md.append("## 1. Общие итоги\n\n")
md.append("| Метрика | Значение |\n|---|---|")
md.append(f"| Всего строк (TITAN, период) | {len(titan_rows_period):,} |")
md.append(f"| Уникальных SKU | {n_skus} |")
md.append(f"| Уникальных цветов | {len(by_color)} |")
md.append(f"| Уникальных фасовок | {len(by_volume)} |")
md.append(f"| Магазинов с TITAN | {n_shops} |")
md.append(f"| Месяцев в выборке | {n_months} |")
md.append(f"| **Продажи (шт.)** | **{total_sales:,}** |")
md.append(f"| Выручка по закуп. цене (₽) | {total_revenue:,.0f} |")
md.append(f"| Средняя закуп. цена | {total_revenue/max(total_sales,1):.0f} ₽/шт. |")
md.append(f"| **Остаток на 02.10.2026 (шт.)** | **{total_stock_unique:,}** |")
md.append(f"| В пути (шт.) | {total_in_transit_unique:,} |")
md.append(f"| **Остаток + в пути (шт.)** | **{total_stock_with_transit:,}** |")
md.append("\n---\n\n")

md.append("## 2. Динамика продаж по месяцам\n\n")
md.append("| Месяц | Продажи (шт.) | Доля | Выручка (₽) | Средний чек |\n")
md.append("|---|---:|---:|---:|---:|")
for m in unique_months:
    d = by_month[m]
    share = d['sales'] / max(total_sales, 1) * 100
    avg_check = d['revenue'] / max(d['sales'], 1) if d['sales'] > 0 else 0
    md.append(f"| {m} | {d['sales']:,} | {share:.1f}% | {d['revenue']:,.0f} | {avg_check:.0f} |")
md.append("\n---\n\n")

md.append("## 3. Продажи по цветам\n\n")
md.append("| Цвет | Продажи (шт.) | Доля | Остаток | В пути | Ост.+в пути | Выручка (₽) | Магаз. | SKU |\n")
md.append("|---|---:|---:|---:|---:|---:|---:|---:|---:|")
sorted_colors = sorted(by_color.items(), key=lambda x: -x[1]['sales'])
for color, d in sorted_colors:
    share = d['sales'] / max(total_sales, 1) * 100
    swt = d['stock'] + d['in_transit']
    md.append(f"| {color} | {d['sales']:,} | {share:.1f}% | {d['stock']:,} | {d['in_transit']:,} | {swt:,} | {d['revenue']:,.0f} | {len(d['shop_count'])} | {len(d['sku_count'])} |")
md.append("\n---\n\n")

md.append("## 4. Продажи по объёмам фасовки\n\n")
md.append("| Объём | Продажи (шт.) | Доля | Остаток | В пути | Ост.+в пути | Выручка (₽) | Магаз. | SKU |\n")
md.append("|---|---:|---:|---:|---:|---:|---:|---:|---:|")
sorted_volumes = sorted(by_volume.items(), key=lambda x: -x[1]['sales'])
for vol, d in sorted_volumes:
    share = d['sales'] / max(total_sales, 1) * 100
    swt = d['stock'] + d['in_transit']
    md.append(f"| {vol} | {d['sales']:,} | {share:.1f}% | {d['stock']:,} | {d['in_transit']:,} | {swt:,} | {d['revenue']:,.0f} | {len(d['shop_count'])} | {len(d['sku_count'])} |")
md.append("\n---\n\n")

md.append("## 5. Матрица: Цвет × Объём (продажи, шт.)\n\n")
all_volumes_sorted = sorted({r['volume_str'] for r in parsed_records},
                             key=lambda v: (0, float(v.split()[0].replace(',', '.')))
                             if v != '—' and v.split()[0].replace(',', '').replace('.', '').isdigit()
                             else (1, v))
md.append("| Цвет | " + " | ".join(all_volumes_sorted) + " | ИТОГО |\n")
md.append("|" + "---|" * (len(all_volumes_sorted) + 2))
for color, d in sorted_colors:
    row = f"| {color} "
    for v in all_volumes_sorted:
        row += f"| {matrix[color].get(v, 0):,} "
    row += f"| {d['sales']:,} |"
    md.append(row)
md.append("\n---\n\n")

md.append("## 6. Топ-20 SKU по продажам\n\n")
md.append("| # | SKU | Цвет | Объём | Продажи (шт.) | Доля | Остаток | В пути | Выручка (₽) | Магаз. |\n")
md.append("|---:|---|---|---|---:|---:|---:|---:|---:|---:|")
top_skus = sorted(by_sku.items(), key=lambda x: -x[1]['sales'])[:20]
for i, (sku, d) in enumerate(top_skus, 1):
    share = d['sales'] / max(total_sales, 1) * 100
    md.append(f"| {i} | {sku[:55]} | {d['color']} | {d['volume_str']} | {d['sales']:,} | {share:.1f}% | {d['stock']:,} | {d['in_transit']:,} | {d['revenue']:,.0f} | {len(d['shop_count'])} |")
md.append("\n---\n\n")

md.append("## 7. Продажи и остатки по магазинам\n\n")
md.append("| Магазин | Продажи (шт.) | Доля | Остаток | В пути | **Ост.+в пути** | Выручка (₽) | SKU |\n")
md.append("|---|---:|---:|---:|---:|---:|---:|---:|")
sorted_shops = sorted(by_shop.items(), key=lambda x: -x[1]['sales'])
for shop, d in sorted_shops:
    share = d['sales'] / max(total_sales, 1) * 100
    swt = d['stock'] + d['in_transit']
    md.append(f"| {shop} | {d['sales']:,} | {share:.1f}% | {d['stock']:,} | {d['in_transit']:,} | **{swt:,}** | {d['revenue']:,.0f} | {len(d['sku_count'])} |")
md.append("\n---\n\n")

md.append("## 8. Анализ закономерностей и тенденций\n\n")

best_color = sorted_colors[0] if sorted_colors else None
worst_color = sorted_colors[-1] if sorted_colors else None
md.append("### 8.1. Цветовая структура спроса\n\n")
if best_color:
    md.append(f"- **Лидер по цвету:** {best_color[0]} — {best_color[1]['sales']:,} шт. ({best_color[1]['sales']/total_sales*100:.1f}% от всех продаж TITAN)\n")
if len(sorted_colors) > 1 and worst_color != best_color:
    md.append(f"- **Аутсайдер:** {worst_color[0]} — {worst_color[1]['sales']:,} шт. ({worst_color[1]['sales']/total_sales*100:.1f}%)\n")
top3_colors = sorted_colors[:3]
top3_share = sum(c[1]['sales'] for c in top3_colors) / max(total_sales, 1) * 100
md.append(f"- **Топ-3 цвета** ({', '.join(c[0] for c in top3_colors)}) формируют {top3_share:.1f}% продаж\n")
md.append("\n")

md.append("### 8.2. Структура спроса по фасовкам\n\n")
if sorted_volumes:
    best_vol = sorted_volumes[0]
    md.append(f"- **Лидер по фасовке:** {best_vol[0]} — {best_vol[1]['sales']:,} шт. ({best_vol[1]['sales']/total_sales*100:.1f}%)\n")
    top2_vols = sorted_volumes[:2]
    top2_share = sum(v[1]['sales'] for v in top2_vols) / max(total_sales, 1) * 100
    md.append(f"- **Топ-2 фасовки** ({', '.join(v[0] for v in top2_vols)}) дают {top2_share:.1f}% продаж\n")
md.append("\n")

md.append("### 8.3. География спроса\n\n")
if sorted_shops:
    md.append(f"- **Лидер по магазину:** {sorted_shops[0][0]} — {sorted_shops[0][1]['sales']:,} шт. ({sorted_shops[0][1]['sales']/total_sales*100:.1f}%)\n")
    top5_shops_share = sum(s[1]['sales'] for s in sorted_shops[:5]) / max(total_sales, 1) * 100
    md.append(f"- **Топ-5 магазинов** дают {top5_shops_share:.1f}% продаж из {n_shops} торговых точек\n")
md.append("\n")

md.append("### 8.4. Оборачиваемость и остатки\n\n")
md.append(f"- Период анализа: {n_months} мес. ({unique_months[0]} — {unique_months[-1]})\n")
md.append(f"- Суммарные продажи за период: {total_sales:,} шт.\n")
md.append(f"- Средние продажи/мес.: {avg_monthly_sales:.0f} шт.\n")
md.append(f"- Остаток на дату выгрузки (02.10.2026): {total_stock_unique:,} шт.\n")
md.append(f"- В пути: {total_in_transit_unique:,} шт.\n")
md.append(f"- **Остаток + в пути:** {total_stock_with_transit:,} шт.\n")
md.append(f"- Коэффициент оборачиваемости (продажи/остаток): {turnover_ratio:.2f}\n")
md.append(f"- В среднем за месяц: {monthly_turnover:.2f} об/мес.\n")
md.append(f"- **Запас на складе (при среднегодовом темпе):** ~{days_of_stock:.0f} дней\n")
# Also calculate peak month turnover
peak_month_sales = max((by_month[m]['sales'] for m in unique_months), default=0)
peak_days_of_stock = total_stock_with_transit / max(peak_month_sales / 30, 1)
md.append(f"- Пиковые продажи/мес. (лучший месяц): {peak_month_sales:,} шт.\n")
md.append(f"- **Запас при пиковом темпе:** ~{peak_days_of_stock:.0f} дней\n")
md.append("\n")

md.append("### 8.5. Товарный портфель (BCG-анализ)\n\n")
avg_sku_sales = total_sales / max(n_skus, 1)
high_share_skus = [(s, d) for s, d in by_sku.items() if d['sales'] > avg_sku_sales * 2]
mid_share_skus = [(s, d) for s, d in by_sku.items() if avg_sku_sales * 0.5 < d['sales'] <= avg_sku_sales * 2]
low_share_skus = [(s, d) for s, d in by_sku.items() if d['sales'] <= avg_sku_sales * 0.5]
md.append("| Категория | Кол-во SKU | Доля SKU | Доля продаж | Что делать |\n|---|---:|---:|---:|---|")
high_sales = sum(d['sales'] for _, d in high_share_skus)
mid_sales = sum(d['sales'] for _, d in mid_share_skus)
low_sales = sum(d['sales'] for _, d in low_share_skus)
md.append(f"| ⭐ Звёзды (>2× среднего) | {len(high_share_skus)} | {len(high_share_skus)/n_skus*100:.0f}% | {high_sales/total_sales*100:.0f}% | 100% наличие во всех ТТ, расширить выкладку |")
md.append(f"| 🐄 Дойные коровы (0,5-2× среднего) | {len(mid_share_skus)} | {len(mid_share_skus)/n_skus*100:.0f}% | {mid_sales/total_sales*100:.0f}% | Поддерживать MML, не расширять |")
md.append(f"| ❓ Собаки (<0,5× среднего) | {len(low_share_skus)} | {len(low_share_skus)/n_skus*100:.0f}% | {low_sales/total_sales*100:.0f}% | Рассмотреть вывод из ассортимента |")
md.append("\n---\n\n")

md.append("### 8.6. Рекомендации\n\n")
md.append("**На основе анализа периода май — октябрь 2026:**\n\n")
if best_color:
    md.append(f"1. **Цветовая стратегия.** Лидер — {best_color[0]} ({best_color[1]['sales']/total_sales*100:.1f}%). Топ-3 цвета дают {top3_share:.0f}% продаж. Рекомендуется 100% наличие топ-3 во всех ТТ.\n")
if sorted_volumes:
    md.append(f"2. **Фасовочная стратегия.** Лидер — {best_vol[0]} ({best_vol[1]['sales']/total_sales*100:.1f}%). Поддерживать 100% наличие топ-3 фасовок.\n")
md.append(f"3. **Оборачиваемость.** Запас {total_stock_with_transit:,} шт. при средних продажах {avg_monthly_sales:.0f} шт./мес. = ~{days_of_stock:.0f} дней. ")
if days_of_stock > 180:
    md.append(f"Сезон TITAN (май-сентябрь) — пик продаж пройден, нужно сократить остатки до 30% к ноябрю.\n")
else:
    md.append(f"Оборачиваемость в норме.\n")
md.append(f"4. **Ассортимент.** Из {n_skus} SKU: {len(high_share_skus)} «звёзд» ({high_sales/total_sales*100:.0f}% продаж), {len(low_share_skus)} «собак» ({low_sales/total_sales*100:.0f}% продаж). Оптимизация MML до 25-30 SKU поднимет оборачиваемость.\n")
md.append("\n---\n\n*Анализ подготовлен на основе данных Лемана Про от 02.10.2026. Период: май — октябрь 2026.*\n")

with open(ANALYSIS_TXT, 'w', encoding='utf-8') as f:
    f.write('\n'.join(md))
print(f"\nAnalysis: {ANALYSIS_TXT} ({os.path.getsize(ANALYSIS_TXT)} bytes)")

# ── Console summary ───────────────────────────────────────────────────
print(f"\n{'='*70}")
print(f"TOP 5 SHOPS:")
for s, d in sorted_shops[:5]:
    swt = d['stock'] + d['in_transit']
    print(f"  {s:25s} sales={d['sales']:6,}  stock={d['stock']:6,}  in_tr={d['in_transit']:5,}  st+tr={swt:6,}  SKU={len(d['sku_count'])}")
print(f"\nTOP 5 COLORS:")
for c, d in sorted_colors[:5]:
    print(f"  {c:25s} {d['sales']:6,} шт. ({d['sales']/total_sales*100:5.1f}%)  stock={d['stock']:,}")
print(f"\nTOP 5 VOLUMES:")
for v, d in sorted_volumes[:5]:
    print(f"  {v:15s} {d['sales']:6,} шт. ({d['sales']/total_sales*100:5.1f}%)  stock={d['stock']:,}")
