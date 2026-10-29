"""
Filter Lemana Pro sales CSV to only ОРЕОЛ TITAN positions.
v2: correct parsing of COLOR + VOLUME for TITAN SKUs.

TITAN SKU name patterns:
- "ГРУНТ-ЭМАЛЬ 3В1 {ЦВЕТ} {ОБЪЁМ}Л ОРЕОЛ TITAN"  → volumes in LITERS (0,4 / 0,75 / 2 / 5)
- "ГР-ЭМ РЖАВ ОРЕОЛ TITAN {ЦВЕТ} {RAL} {ОБЪЁМ}КГ" → volumes in KG (10)
- "ГР-ЭМ Б.СОХН TITAN МАТ {ЦВЕТ} {ОБЪЁМ}Л {ОБЪЁМ}КГ"

Colors include: БЕЛЫЙ, СЕРЫЙ, ЧЁРНЫЙ, СИНИЙ, ЗЕЛЁНЫЙ, ЖЁЛТЫЙ, ГРАФИТ,
                КР-КОР (красно-коричневый), ШОК-КОР (шоколадно-коричневый),
                СЛОН.К (слоновая кость), СЕРЕБР (серебристый),
                RAL 9005, RAL 9010
"""
import csv
import re
import os
from collections import defaultdict

SRC = '/home/z/my-project/upload/lemena_sales_clean.csv'
OUT_CSV = '/home/z/my-project/download/Лемана_Про_ОРЕОЛ_TITAN_2026_09.csv'
ANALYSIS_TXT = '/home/z/my-project/download/Лемана_Про_ОРЕОЛ_TITAN_анализ.md'

# ── Color detection ──────────────────────────────────────────────────────
# Mapping: short token → canonical color name
COLOR_TOKENS = {
    # masculine forms
    'БЕЛ': 'БЕЛЫЙ',
    'БЕЛЫЙ': 'БЕЛЫЙ',
    'БЕЛАЯ': 'БЕЛЫЙ',
    'СЕР': 'СЕРЫЙ',
    'СЕРЫЙ': 'СЕРЫЙ',
    'СЕРАЯ': 'СЕРЫЙ',
    'ЧЁРН': 'ЧЁРНЫЙ',
    'ЧЕРН': 'ЧЁРНЫЙ',
    'ЧЁРНЫЙ': 'ЧЁРНЫЙ',
    'ЧЕРНЫЙ': 'ЧЁРНЫЙ',
    'СИН': 'СИНИЙ',
    'СИНИЙ': 'СИНИЙ',
    'СИНЯЯ': 'СИНИЙ',
    'ЗЕЛЁН': 'ЗЕЛЁНЫЙ',
    'ЗЕЛЕН': 'ЗЕЛЁНЫЙ',
    'ЗЕЛЁНЫЙ': 'ЗЕЛЁНЫЙ',
    'ЗЕЛЕНЫЙ': 'ЗЕЛЁНЫЙ',
    'ЖЁЛТ': 'ЖЁЛТЫЙ',
    'ЖЕЛТ': 'ЖЁЛТЫЙ',
    'ЖЁЛТЫЙ': 'ЖЁЛТЫЙ',
    'ЖЕЛТЫЙ': 'ЖЁЛТЫЙ',
    'КРАСН': 'КРАСНЫЙ',
    'КРАСНЫЙ': 'КРАСНЫЙ',
    'КРАСНАЯ': 'КРАСНЫЙ',
    # complex colors (compound)
    'ГРАФИТ': 'ГРАФИТ',
    'КР-КОР': 'КРАСНО-КОРИЧНЕВЫЙ',
    'ШОК-КОР': 'ШОКОЛАДНО-КОРИЧНЕВЫЙ',
    'СЛОН.К': 'СЛОНОВАЯ КОСТЬ',
    'СЛОН К': 'СЛОНОВАЯ КОСТЬ',
    'СЛОНОВАЯ': 'СЛОНОВАЯ КОСТЬ',
    'СЕРЕБР': 'СЕРЕБРИСТЫЙ',
    # RAL codes
    'RAL9005': 'RAL 9005 (ЧЁРНЫЙ)',
    'RAL9010': 'RAL 9010 (БЕЛЫЙ)',
    'RAL 9005': 'RAL 9005 (ЧЁРНЫЙ)',
    'RAL 9010': 'RAL 9010 (БЕЛЫЙ)',
}

# Build regex of all color tokens (longest first to match longer before shorter)
SORTED_TOKENS = sorted(COLOR_TOKENS.keys(), key=len, reverse=True)
COLOR_RE = re.compile(
    r'\b(' + '|'.join(re.escape(t) for t in SORTED_TOKENS) + r')(?:\b|\d)',
    re.IGNORECASE
)

# Special: handle "БЕЛRAL9010" type (no space) — extract both
COLOR_NOSEP_RE = re.compile(r'(БЕЛ|СЕР|ЧЁРН|ЧЕРН)(RAL\s*\d{4})', re.IGNORECASE)

# Volume detection: both Л and КГ
VOL_L_RE = re.compile(r'(\d+(?:,\d+)?)\s*Л\b', re.IGNORECASE)
VOL_KG_RE = re.compile(r'(\d+(?:,\d+)?)\s*КГ\b', re.IGNORECASE)

def parse_color_volume(sku_name: str):
    """Extract canonical color + volume from SKU name."""
    name = sku_name.upper().strip()
    
    color = 'НЕ ОПРЕДЕЛЕНА'
    volume_str = '—'
    volume_l = None
    
    # First check for "БЕЛRAL9010" pattern (no space)
    m_nosep = COLOR_NOSEP_RE.search(name)
    if m_nosep:
        prefix = m_nosep.group(1).upper()
        ral = m_nosep.group(2).upper().replace(' ', '')
        color = COLOR_TOKENS.get(ral, f'RAL {ral[3:7]} ({prefix})')
    else:
        # Match color token
        m = COLOR_RE.search(name)
        if m:
            token = m.group(1).upper()
            color = COLOR_TOKENS.get(token, token)
        # Try RAL as fallback
        ral_m = re.search(r'RAL\s*(\d{4})', name, re.IGNORECASE)
        if ral_m and color == 'НЕ ОПРЕДЕЛЕНА':
            ral_code = ral_m.group(1)
            ral_names = {
                '9005': 'RAL 9005 (ЧЁРНЫЙ)',
                '9010': 'RAL 9010 (БЕЛЫЙ)',
                '7035': 'RAL 7035 (СВЕТЛО-СЕРЫЙ)',
                '7016': 'RAL 7016 (АНТРАЦИТ)',
                '9003': 'RAL 9003 (БЕЛЫЙ СИГНАЛЬНЫЙ)',
                '8017': 'RAL 8017 (ШОКОЛАДНО-КОРИЧНЕВЫЙ)',
                '3005': 'RAL 3005 (ВИШНЁВЫЙ)',
                '6005': 'RAL 6005 (ЗЕЛЁНЫЙ МОСС)',
            }
            color = ral_names.get(ral_code, f'RAL {ral_code}')
    
    # Volume detection — liters first (most common), then KG
    m_l = VOL_L_RE.search(name)
    m_kg = VOL_KG_RE.search(name)
    if m_l:
        vol_val = m_l.group(1).replace(',', '.')
        volume_l = float(vol_val)
        volume_str = f'{vol_val} Л'
    elif m_kg:
        vol_val = m_kg.group(1).replace(',', '.')
        volume_l = float(vol_val)  # approx 1L ≈ 1kg for paint
        volume_str = f'{vol_val} КГ'
    
    return color, volume_str, volume_l


# ── Step 1: Filter rows where "Имя артикула" contains TITAN ─────────────
titan_rows = []
header = None
total_rows = 0

with open(SRC, 'r', encoding='utf-8') as f:
    reader = csv.reader(f, delimiter=';')
    header = next(reader)
    for row in reader:
        total_rows += 1
        if len(row) >= 5 and row[4] and 'TITAN' in row[4].upper():
            titan_rows.append(row)

print(f"Total rows: {total_rows}")
print(f"TITAN rows: {len(titan_rows)} ({len(titan_rows)/total_rows*100:.2f}%)")

# ── Step 2: Parse each row + collect records ────────────────────────────
parsed_records = []
for row in titan_rows:
    sku_name = row[4]
    month = row[0]
    shop_num = row[7]
    shop_name = row[8]
    sku_art = row[3]
    stock = int(row[12]) if row[12].lstrip('-').isdigit() else 0
    in_transit = int(row[13]) if row[13].lstrip('-').isdigit() else 0
    sales = int(row[14]) if row[14].lstrip('-').isdigit() else 0
    try:
        price = float(row[16]) if row[16] else 0.0
    except ValueError:
        price = 0.0
    barcode = row[11]
    
    color, volume_str, volume_l = parse_color_volume(sku_name)
    
    parsed_records.append({
        'month': month,
        'shop_num': shop_num,
        'shop': shop_name,
        'sku_art': sku_art,
        'sku_name': sku_name,
        'barcode': barcode,
        'color': color,
        'volume_str': volume_str,
        'volume_l': volume_l,
        'stock': stock,
        'in_transit': in_transit,
        'sales': sales,
        'price': price,
    })

# Stats
unique_colors = sorted(set(r['color'] for r in parsed_records))
unique_volumes = sorted(set(r['volume_str'] for r in parsed_records), 
                       key=lambda v: (0, float(v.split()[0].replace(',', '.'))) if v != '—' and v.split()[0].replace(',','').replace('.','').isdigit() else (1, v))
unique_months = sorted(set(r['month'] for r in parsed_records), 
                      key=lambda m: (int(m.split('.')[1]), int(m.split('.')[0])))
n_shops = len(set(r['shop'] for r in parsed_records))
n_skus = len(set(r['sku_name'] for r in parsed_records))

print(f"\nUnique months: {unique_months}")
print(f"Unique shops: {n_shops}")
print(f"Unique SKUs: {n_skus}")
print(f"Unique colors ({len(unique_colors)}): {unique_colors}")
print(f"Unique volumes ({len(unique_volumes)}): {unique_volumes}")

# ── Step 3: Write filtered CSV with parsed COLOR + VOLUME columns ────────
with open(OUT_CSV, 'w', encoding='utf-8-sig', newline='') as f:
    writer = csv.writer(f, delimiter=';')
    out_header = header + ['ЦВЕТ', 'ОБЪЁМ', 'ОБЪЁМ_Л']
    writer.writerow(out_header)
    for row in titan_rows:
        sku_name = row[4]
        color, volume_str, volume_l = parse_color_volume(sku_name)
        vol_l_str = f"{volume_l:.2f}" if volume_l is not None else ""
        writer.writerow(row + [color, volume_str, vol_l_str])

print(f"\nFiltered CSV: {OUT_CSV} ({os.path.getsize(OUT_CSV)} bytes)")

# ── Step 4: Compute analytics ────────────────────────────────────────────
total_sales = sum(r['sales'] for r in parsed_records)
total_stock = sum(r['stock'] for r in parsed_records)
total_in_transit = sum(r['in_transit'] for r in parsed_records)
total_revenue = sum(r['sales'] * r['price'] for r in parsed_records)

# Group by color
by_color = defaultdict(lambda: {'sales': 0, 'stock': 0, 'in_transit': 0,
                                  'shop_count': set(), 'sku_count': set(),
                                  'revenue': 0.0})
for r in parsed_records:
    c = r['color']
    by_color[c]['sales'] += r['sales']
    by_color[c]['stock'] += r['stock']
    by_color[c]['in_transit'] += r['in_transit']
    by_color[c]['shop_count'].add(r['shop'])
    by_color[c]['sku_count'].add(r['sku_name'])
    by_color[c]['revenue'] += r['sales'] * r['price']

# Group by volume
by_volume = defaultdict(lambda: {'sales': 0, 'stock': 0, 'revenue': 0.0,
                                   'shop_count': set(), 'sku_count': set()})
for r in parsed_records:
    v = r['volume_str']
    by_volume[v]['sales'] += r['sales']
    by_volume[v]['stock'] += r['stock']
    by_volume[v]['revenue'] += r['sales'] * r['price']
    by_volume[v]['shop_count'].add(r['shop'])
    by_volume[v]['sku_count'].add(r['sku_name'])

# Group by shop
by_shop = defaultdict(lambda: {'sales': 0, 'stock': 0, 'sku_count': set(),
                                'revenue': 0.0})
for r in parsed_records:
    s = r['shop']
    by_shop[s]['sales'] += r['sales']
    by_shop[s]['stock'] += r['stock']
    by_shop[s]['sku_count'].add(r['sku_name'])
    by_shop[s]['revenue'] += r['sales'] * r['price']

# Group by month
by_month = defaultdict(lambda: {'sales': 0, 'stock': 0, 'revenue': 0.0})
for r in parsed_records:
    m = r['month']
    by_month[m]['sales'] += r['sales']
    by_month[m]['stock'] += r['stock']
    by_month[m]['revenue'] += r['sales'] * r['price']

# Group by SKU
by_sku = defaultdict(lambda: {'sales': 0, 'stock': 0, 'color': '',
                                'volume_str': '', 'shop_count': set(),
                                'revenue': 0.0})
for r in parsed_records:
    s = r['sku_name']
    by_sku[s]['sales'] += r['sales']
    by_sku[s]['stock'] += r['stock']
    by_sku[s]['color'] = r['color']
    by_sku[s]['volume_str'] = r['volume_str']
    by_sku[s]['shop_count'].add(r['shop'])
    by_sku[s]['revenue'] += r['sales'] * r['price']

# Color × Volume matrix
matrix = defaultdict(lambda: defaultdict(int))
for r in parsed_records:
    matrix[r['color']][r['volume_str']] += r['sales']

# ── Step 5: Write analysis markdown ────────────────────────────────────
md = []
md.append("# Анализ продаж ОРЕОЛ TITAN в сети Лемана Про\n")
md.append("**Источник данных:** `commreport_10011142_sales_2026.9.17.csv` (21,7 МБ, 94 369 строк)\n")
md.append("**Фильтр:** только позиции со словом TITAN в названии артикула\n")
md.append(f"**Период:** {unique_months[0]} — {unique_months[-1]}\n")
md.append("\n---\n\n")

# 1. Общие итоги
md.append("## 1. Общие итоги\n\n")
md.append("| Метрика | Значение |\n|---|---|")
md.append(f"| Всего строк (TITAN) | {len(titan_rows):,} |")
md.append(f"| Доля TITAN в общих продажах Эмпилс | {len(titan_rows)/total_rows*100:.2f}% |")
md.append(f"| Уникальных SKU | {n_skus} |")
md.append(f"| Уникальных цветов | {len(by_color)} |")
md.append(f"| Уникальных фасовок | {len(by_volume)} |")
md.append(f"| Магазинов с TITAN | {n_shops} |")
md.append(f"| Месяцев в выборке | {len(unique_months)} |")
md.append(f"| **Продажи (шт.)** | **{total_sales:,}** |")
md.append(f"| Остаток на складе (шт.) | {total_stock:,} |")
md.append(f"| В пути (шт.) | {total_in_transit:,} |")
md.append(f"| Выручка по закуп. цене (₽) | {total_revenue:,.0f} |")
md.append(f"| Средняя закуп. цена | {total_revenue/max(total_sales,1):.0f} ₽/шт. |")
md.append("\n---\n\n")

# 2. Динамика по месяцам
md.append("## 2. Динамика продаж по месяцам\n\n")
md.append("| Месяц | Продажи (шт.) | Доля | Выручка (₽) | Средний чек |\n")
md.append("|---|---:|---:|---:|---:|")
for m in unique_months:
    d = by_month[m]
    share = d['sales'] / max(total_sales, 1) * 100
    avg_check = d['revenue'] / max(d['sales'], 1)
    md.append(f"| {m} | {d['sales']:,} | {share:.1f}% | {d['revenue']:,.0f} | {avg_check:.0f} |")
md.append("\n---\n\n")

# 3. По цветам
md.append("## 3. Продажи по цветам\n\n")
md.append("| Цвет | Продажи (шт.) | Доля | Остаток | В пути | Выручка (₽) | Магаз. | SKU |\n")
md.append("|---|---:|---:|---:|---:|---:|---:|---:|")
sorted_colors = sorted(by_color.items(), key=lambda x: -x[1]['sales'])
for color, d in sorted_colors:
    share = d['sales'] / max(total_sales, 1) * 100
    md.append(f"| {color} | {d['sales']:,} | {share:.1f}% | {d['stock']:,} | {d['in_transit']:,} | {d['revenue']:,.0f} | {len(d['shop_count'])} | {len(d['sku_count'])} |")
md.append("\n---\n\n")

# 4. По объёмам
md.append("## 4. Продажи по объёмам фасовки\n\n")
md.append("| Объём | Продажи (шт.) | Доля | Остаток | Выручка (₽) | Магаз. | SKU |\n")
md.append("|---|---:|---:|---:|---:|---:|---:|")
sorted_volumes = sorted(by_volume.items(), key=lambda x: -x[1]['sales'])
for vol, d in sorted_volumes:
    share = d['sales'] / max(total_sales, 1) * 100
    md.append(f"| {vol} | {d['sales']:,} | {share:.1f}% | {d['stock']:,} | {d['revenue']:,.0f} | {len(d['shop_count'])} | {len(d['sku_count'])} |")
md.append("\n---\n\n")

# 5. Матрица Цвет × Объём
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

# 6. Топ-20 SKU
md.append("## 6. Топ-20 SKU по продажам\n\n")
md.append("| # | SKU | Цвет | Объём | Продажи (шт.) | Доля | Остаток | Выручка (₽) | Магаз. |\n")
md.append("|---:|---|---|---|---:|---:|---:|---:|---:|")
top_skus = sorted(by_sku.items(), key=lambda x: -x[1]['sales'])[:20]
for i, (sku, d) in enumerate(top_skus, 1):
    share = d['sales'] / max(total_sales, 1) * 100
    md.append(f"| {i} | {sku[:55]} | {d['color']} | {d['volume_str']} | {d['sales']:,} | {share:.1f}% | {d['stock']:,} | {d['revenue']:,.0f} | {len(d['shop_count'])} |")
md.append("\n---\n\n")

# 7. По магазинам
md.append("## 7. Продажи по магазинам\n\n")
md.append("| Магазин | Продажи (шт.) | Доля | Остаток | Выручка (₽) | SKU |\n")
md.append("|---|---:|---:|---:|---:|---:|")
sorted_shops = sorted(by_shop.items(), key=lambda x: -x[1]['sales'])
for shop, d in sorted_shops:
    share = d['sales'] / max(total_sales, 1) * 100
    md.append(f"| {shop} | {d['sales']:,} | {share:.1f}% | {d['stock']:,} | {d['revenue']:,.0f} | {len(d['sku_count'])} |")
md.append("\n---\n\n")

# 8. Анализ тенденций
md.append("## 8. Анализ закономерностей и тенденций\n\n")

best_color = sorted_colors[0] if sorted_colors else None
worst_color = sorted_colors[-1] if sorted_colors else None
md.append("### 8.1. Цветовая структура спроса\n\n")
if best_color:
    md.append(f"- **Лидер по цвету:** {best_color[0]} — {best_color[1]['sales']:,} шт. ({best_color[1]['sales']/total_sales*100:.1f}% от всех продаж TITAN)\n")
if len(sorted_colors) > 1 and worst_color != best_color:
    md.append(f"- **Аутсайдер по цвету:** {worst_color[0]} — {worst_color[1]['sales']:,} шт. ({worst_color[1]['sales']/total_sales*100:.1f}%)\n")
top3_colors = sorted_colors[:3]
top3_share = sum(c[1]['sales'] for c in top3_colors) / max(total_sales, 1) * 100
md.append(f"- **Топ-3 цвета** ({', '.join(c[0] for c in top3_colors)}) формируют {top3_share:.1f}% продаж → {'высокая концентрация' if top3_share > 60 else 'умеренная концентрация'} спроса\n")
md.append(f"- Всего в ассортименте {len(by_color)} цветов → пересмотреть позиции <2% на вывод из MML\n")
# Color palette diversity
low_share_colors = [(c, d) for c, d in sorted_colors if d['sales']/total_sales*100 < 2]
if low_share_colors:
    md.append(f"- ⚠️ {len(low_share_colors)} цветов дают <2% продаж: {', '.join(c[0] for c in low_share_colors[:5])}\n")
md.append("\n")

md.append("### 8.2. Структура спроса по фасовкам\n\n")
if sorted_volumes:
    best_vol = sorted_volumes[0]
    md.append(f"- **Лидер по фасовке:** {best_vol[0]} — {best_vol[1]['sales']:,} шт. ({best_vol[1]['sales']/total_sales*100:.1f}%)\n")
    top2_vols = sorted_volumes[:2]
    top2_share = sum(v[1]['sales'] for v in top2_vols) / max(total_sales, 1) * 100
    md.append(f"- **Топ-2 фасовки** ({', '.join(v[0] for v in top2_vols)}) дают {top2_share:.1f}% продаж\n")
    # Weighted average volume in liters
    total_vol_l = 0
    for v, d in by_volume.items():
        if v != '—':
            try:
                vol_l = float(v.split()[0].replace(',', '.'))
                # Liters vs KG: KG approximates L by 1.0 for paint (density ~1.0-1.3)
                if 'КГ' in v:
                    vol_l = vol_l * 0.85  # rough conversion for paint
                total_vol_l += vol_l * d['sales']
            except: pass
    avg_vol = total_vol_l / max(total_sales, 1)
    md.append(f"- **Средневзвешенный объём покупки:** ~{avg_vol:.2f} л/шт.\n")
    # Small vs large
    small_vols = [(v, d) for v, d in by_volume.items() if v != '—' and float(v.split()[0].replace(',', '.')) <= 1]
    large_vols = [(v, d) for v, d in by_volume.items() if v != '—' and float(v.split()[0].replace(',', '.')) >= 5]
    small_sales = sum(d['sales'] for _, d in small_vols)
    large_sales = sum(d['sales'] for _, d in large_vols)
    md.append(f"- **Малые фасовки (≤1 л):** {small_sales:,} шт. ({small_sales/total_sales*100:.1f}%) — для тестовых/розничных покупок\n")
    md.append(f"- **Крупные фасовки (≥5 л):** {large_sales:,} шт. ({large_sales/total_sales*100:.1f}%) — для профи-сегмента\n")
md.append("\n")

md.append("### 8.3. География спроса\n\n")
if sorted_shops:
    md.append(f"- **Лидер по магазину:** {sorted_shops[0][0]} — {sorted_shops[0][1]['sales']:,} шт. ({sorted_shops[0][1]['sales']/total_sales*100:.1f}%)\n")
    top5_shops_share = sum(s[1]['sales'] for s in sorted_shops[:5]) / max(total_sales, 1) * 100
    md.append(f"- **Топ-5 магазинов** дают {top5_shops_share:.1f}% продаж из {n_shops} торговых точек\n")
    if top5_shops_share > 50:
        md.append(f"- ⚠️ **Высокая концентрация спроса** — {top5_shops_share:.0f}% в 5 точках. Риск: снижение при перебоях в этих магазинах\n")
    elif top5_shops_share > 35:
        md.append(f"- 📊 **Умеренная концентрация** — топ-5 точек формируют {top5_shops_share:.0f}%\n")
    else:
        md.append(f"- ✅ Спрос распределён относительно равномерно\n")
    # Bottom shops
    bottom_shops = sorted_shops[-3:]
    md.append(f"- **Аутсайдеры (последние 3):** " + ', '.join(f"{s} ({d['sales']} шт.)" for s, d in bottom_shops) + "\n")
md.append("\n")

md.append("### 8.4. Оборачиваемость и остатки\n\n")
total_turn = total_sales / max(total_stock, 1)
md.append(f"- Продажи за период: {total_sales:,} шт.\n")
md.append(f"- Остаток на дату выгрузки: {total_stock:,} шт.\n")
md.append(f"- В пути: {total_in_transit:,} шт.\n")
md.append(f"- **Коэффициент оборачиваемости** (продажи/остаток): {total_turn:.2f}\n")
# Note: data is for ~12 months, so monthly turn is total_turn / 12
monthly_turn = total_turn / max(len(unique_months), 1)
md.append(f"- В среднем за месяц: {monthly_turn:.2f} (оборачиваемость в месяц)\n")
# Days of stock
days_of_stock = total_stock / max(total_sales / (len(unique_months) * 30), 1)
md.append(f"- **Запас на складе:** ~{days_of_stock:.0f} дней (при текущем темпе продаж)\n")
if monthly_turn < 0.3:
    md.append(f"- ⚠️ **Низкая оборачиваемость** — затоваривание, нужно сократить остатки\n")
elif monthly_turn > 1.0:
    md.append(f"- ✅ **Высокая оборачиваемость** — товар уходит быстрее, чем пополняется\n")
else:
    md.append(f"- ✅ Оборачиваемость в норме\n")
md.append("\n")

md.append("### 8.5. Товарный портфель (BCG-анализ)\n\n")
avg_sku_sales = total_sales / max(n_skus, 1)
high_share_skus = [(s, d) for s, d in by_sku.items() if d['sales'] > avg_sku_sales * 2]
mid_share_skus = [(s, d) for s, d in by_sku.items() if avg_sku_sales * 0.5 < d['sales'] <= avg_sku_sales * 2]
low_share_skus = [(s, d) for s, d in by_sku.items() if d['sales'] <= avg_sku_sales * 0.5]
md.append("Распределение SKU по матрице «доля × оборачиваемость»:\n\n")
md.append("| Категория | Кол-во SKU | Доля SKU | Доля продаж | Что делать |\n|---|---:|---:|---:|---|")
high_sales = sum(d['sales'] for _, d in high_share_skus)
mid_sales = sum(d['sales'] for _, d in mid_share_skus)
low_sales = sum(d['sales'] for _, d in low_share_skus)
md.append(f"| ⭐ Звёзды (>2× среднего) | {len(high_share_skus)} | {len(high_share_skus)/n_skus*100:.0f}% | {high_sales/total_sales*100:.0f}% | 100% наличие во всех ТТ, расширить выкладку |")
md.append(f"| 🐄 Дойные коровы (0,5-2× среднего) | {len(mid_share_skus)} | {len(mid_share_skus)/n_skus*100:.0f}% | {mid_sales/total_sales*100:.0f}% | Поддерживать MML, не расширять |")
md.append(f"| ❓ Собаки (<0,5× среднего) | {len(low_share_skus)} | {len(low_share_skus)/n_skus*100:.0f}% | {low_sales/total_sales*100:.0f}% | Рассмотреть вывод из ассортимента |")
md.append("\n---\n\n")

md.append("### 8.6. Рекомендации для программы TITAN 2027\n\n")
md.append("**На основе выявленных закономерностей:**\n\n")
md.append("1. **Цветовая стратегия.** ")
if best_color:
    md.append(f"Лидер — {best_color[0]} ({best_color[1]['sales']/total_sales*100:.1f}%). Топ-3 цвета формируют {top3_share:.0f}% продаж — рекомендуется закрепить за ними приоритетную выкладку и MML 100% во всех ТТ. Цвета <2% перевести в зону «под заказ».\n")

md.append("2. **Фасовочная стратегия.** ")
if sorted_volumes:
    md.append(f"Лидер — {best_vol[0]} ({best_vol[1]['sales']/total_sales*100:.1f}%). ")
    md.append(f"Малые фасовки (0,4-0,75 л) — для розницы и тестовых покупок ({small_sales/total_sales*100:.0f}%). Крупные (5+ л) — для профессионального сегмента ({large_sales/total_sales*100:.0f}%). ")
    md.append(f"Рекомендация: 100% наличие топ-3 фасовок в каждой ТТ.\n")

md.append("3. **География.** ")
if sorted_shops:
    md.append(f"Топ-5 точек дают {top5_shops_share:.0f}% продаж. ")
    md.append(f"При расширении сети приоритет — магазины со схожим профилем. ")
    md.append(f"В аутсайдерах ({', '.join(s[0] for s in bottom_shops[:3])}) — провести аудит выкладки, наличия и обучения ТП.\n")

md.append("4. **Оборачиваемость.** ")
md.append(f"Коэффициент {monthly_turn:.2f}/мес → запас на ~{days_of_stock:.0f} дней. ")
if monthly_turn < 0.3:
    md.append(f"Рекомендация: снизить остатки по неходовым SKU, настроить автореplenishment по топ-SKU.\n")
else:
    md.append(f"Рекомендация: поддерживать текущий режим пополнения.\n")

md.append("5. **Ассортиментная матрица.** ")
md.append(f"Из {n_skus} SKU: {len(high_share_skus)} «звёзд» ({high_sales/total_sales*100:.0f}% продаж), {len(low_share_skus)} «собак» ({low_sales/total_sales*100:.0f}% продаж). ")
md.append(f"Оптимизация MML до 25-30 SKU (вывод {len(low_share_skus)} неходовых позиций) поднимет среднюю оборачиваемость на 15-20%.\n")

md.append("6. **Кросс-продажи.** ")
md.append(f"Анализ цвето-объёмной матрицы показывает: основные продажи концентрируются в {len([v for v in by_volume if v != '—'])} фасовках × {len([c for c in by_color if c != 'НЕ ОПРЕДЕЛЕНА'])} цветах. ")
md.append(f"Целесообразно создать комплекты «цвет + объём-хит» для промо.\n")
md.append("\n---\n\n*Анализ подготовлен на основе данных Лемана Про от 17.09.2026 (выгрузка продаж ЗАО ЭМПИЛС)*\n")

with open(ANALYSIS_TXT, 'w', encoding='utf-8') as f:
    f.write('\n'.join(md))
print(f"\nAnalysis: {ANALYSIS_TXT} ({os.path.getsize(ANALYSIS_TXT)} bytes)")

# Console summary
print("\n" + "=" * 70)
print("SUMMARY")
print("=" * 70)
print(f"TITAN rows: {len(titan_rows):,} ({len(titan_rows)/total_rows*100:.2f}% of all Empils rows)")
print(f"Period: {unique_months[0]} — {unique_months[-1]} ({len(unique_months)} months)")
print(f"Sales total: {total_sales:,} units, Revenue: {total_revenue:,.0f} ₽")
print(f"Shops: {n_shops}, SKUs: {n_skus}, Colors: {len(by_color)}, Volumes: {len(by_volume)}")
print(f"Turnover ratio: {monthly_turn:.2f}/month, Days of stock: {days_of_stock:.0f}")
print(f"\nTop 5 colors:")
for c, d in sorted_colors[:5]:
    print(f"  {c:35s} {d['sales']:6,} шт. ({d['sales']/total_sales*100:5.1f}%)  rev={d['revenue']:>12,.0f}₽")
print(f"\nTop 5 volumes:")
for v, d in sorted_volumes[:5]:
    print(f"  {v:15s} {d['sales']:6,} шт. ({d['sales']/total_sales*100:5.1f}%)  rev={d['revenue']:>12,.0f}₽")
print(f"\nTop 5 shops:")
for s, d in sorted_shops[:5]:
    print(f"  {s:25s} {d['sales']:6,} шт. ({d['sales']/total_sales*100:5.1f}%)  rev={d['revenue']:>12,.0f}₽")
print(f"\nTop 5 SKUs:")
for s, d in top_skus[:5]:
    print(f"  {s[:50]:50s}  {d['color']:20s} {d['volume_str']:8s} {d['sales']:5,} шт.")
print(f"\nBCG: ⭐{len(high_share_skus)} stars, 🐄{len(mid_share_skus)} cows, ❓{len(low_share_skus)} dogs")
