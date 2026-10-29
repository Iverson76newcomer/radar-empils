"""
Parse Excel TITAN sales data and create:
1. Normalized JSON data with parsed SKU attributes
2. Interactive HTML dashboard with charts
3. Markdown analysis report

Excel structure: 3 sheets
- "БРОЗЭКС, Рыгалов" (combined, 235 rows)
- "Рыгалов С.В." (Rygalov only, 135 rows)
- "Общий по точкам" (general totals, 526 rows)

Columns: Дистрибьютор | Класс ТТ | Наименование | АКБ 2025 | Продажи 2025 | АКБ 2026 | Продажи 2026
"""
import openpyxl
import json
import re
import os
from collections import defaultdict

SRC = "/home/z/my-project/upload/Титан запрос Заколодний Роман.xlsx"

# ── SKU parsing ─────────────────────────────────────────────────────────
# Pattern examples:
# "Грунт-эмаль быстросохнущая ОРЕОЛ TITAN матовая белая RAL 9010 0.75 л 0.9 кг"
# "Грунт-эмаль по ржавчине ОРЕОЛ TITAN белый RAL 9010 0.4 л 0.45 кг"
# "Грунт-эмаль по ржавчине ОРЕОЛ TITAN чёрный RAL 9005 8.5 л 8.2 кг"
# "Эмаль с молотковым эффектом ОРЕОЛ TITAN золотистая 0.8 кг"

# Type: "Грунт-эмаль быстросохнущая" / "Грунт-эмаль по ржавчине" / "Эмаль с молотковым эффектом"
# Color tokens in SKU
COLOR_PATTERNS = {
    'БЕЛЫЙ': ['белый', 'белая'],
    'ЧЁРНЫЙ': ['чёрный', 'чёрная', 'черный', 'черная'],
    'СЕРЫЙ': ['серый', 'серая'],
    'ГРАФИТОВЫЙ СЕРЫЙ': ['графитовый серый'],
    'ЖЁЛТЫЙ': ['жёлтый', 'жёлтая', 'желтый', 'желтая'],
    'ЗЕЛЁНЫЙ': ['зеленый', 'зеленая', 'зелёный', 'зелёная'],
    'КРАСНО-КОРИЧНЕВЫЙ': ['красно-коричневый'],
    'ШОКОЛАДНО-КОРИЧНЕВЫЙ': ['шоколадно-коричневый', 'шоколадно-коричневая'],
    'КОРИЧНЕВЫЙ': ['коричневый', 'коричневая'],
    'ТЁМНО-ЗЕЛЁНЫЙ': ['тёмно-зелёная', 'темно-зеленая'],
    'СЕРЕБРИСТЫЙ': ['серебристая', 'серебро'],
    'ЗОЛОТИСТЫЙ': ['золотистая'],
    'МЕДНЫЙ': ['медная'],
}

# RAL: extract RAL XXXX
RAL_RE = re.compile(r'RAL\s*(\d{4})', re.IGNORECASE)
# Volume: extract X л or X.X л
VOL_L_RE = re.compile(r'(\d+(?:\.\d+)?)\s*л\b', re.IGNORECASE)
VOL_KG_RE = re.compile(r'(\d+(?:\.\d+)?)\s*кг\b', re.IGNORECASE)
# Type: detect by prefix
TYPE_PATTERNS = [
    ('Грунт-эмаль быстросохнущая', re.compile(r'^грунт-эмаль\s+быстросохнущая', re.IGNORECASE)),
    ('Грунт-эмаль по ржавчине', re.compile(r'^грунт-эмаль\s+по\s+ржавчине', re.IGNORECASE)),
    ('Эмаль с молотковым эффектом', re.compile(r'^эмаль\s+с\s+молотковым\s+эффектом', re.IGNORECASE)),
    # Fallback: Эмаль общая
    ('Эмаль', re.compile(r'^эмаль', re.IGNORECASE)),
    # Грунт-эмаль без указания типа
    ('Грунт-эмаль (проч.)', re.compile(r'^грунт-эмаль', re.IGNORECASE)),
]
# Surface finish: матовая / матовый / глянцевая
FINISH_PATTERNS = {
    'МАТОВАЯ': ['матовая', 'матовый', 'матовой'],
    'ГЛЯНЦЕВАЯ': ['глянцевая', 'глянцевый'],
    'ПОЛУМАТОВАЯ': ['полуматовая'],
}

def parse_sku(name: str) -> dict:
    """Parse a SKU name into structured attributes."""
    if not name:
        return {'type': '', 'finish': '', 'color': '', 'ral': '', 'volume_l': '', 'volume_kg': ''}
    n = name.strip()
    
    # 1. Type
    sku_type = ''
    for t, pat in TYPE_PATTERNS:
        if pat.search(n):
            sku_type = t
            break
    
    # 2. Finish (матовая/матовый/глянцевая)
    finish = ''
    n_lower = n.lower()
    for fname, patterns in FINISH_PATTERNS.items():
        if any(p in n_lower for p in patterns):
            finish = fname
            break
    
    # 3. Color
    color = ''
    for cname, patterns in COLOR_PATTERNS.items():
        if any(p in n_lower for p in patterns):
            color = cname
            break
    
    # 4. RAL
    ral_match = RAL_RE.search(n)
    ral = f'RAL {ral_match.group(1)}' if ral_match else ''
    # Combine RAL with color if both present
    if ral and color:
        color_display = f'{color} ({ral})'
    elif ral and not color:
        color_display = ral
    else:
        color_display = color
    
    # 5. Volume: prefer liters, fallback to kg
    vol_l_match = VOL_L_RE.search(n)
    vol_kg_match = VOL_KG_RE.search(n)
    vol_l = vol_l_match.group(1) if vol_l_match else ''
    vol_kg = vol_kg_match.group(1) if vol_kg_match else ''
    
    if vol_l:
        volume_display = f'{vol_l} л'
    elif vol_kg:
        volume_display = f'{vol_kg} кг'
    else:
        volume_display = ''
    
    return {
        'type': sku_type,
        'finish': finish,
        'color': color,
        'ral': ral,
        'color_display': color_display,
        'volume_l': vol_l,
        'volume_kg': vol_kg,
        'volume_display': volume_display,
    }


# ── Read all 3 sheets ───────────────────────────────────────────────────
wb = openpyxl.load_workbook(SRC, data_only=True)

def read_sheet(sheet, has_distributor_col=True):
    """Read a sheet and return list of records.
    Column structure:
    A (0): Дистрибьютор (or Класс ТТ if no distributor column)
    B (1): Класс ТТ (or Наименование)
    C (2): Наименование
    D (3): АКБ 2025
    E (4): Продажи 2025 (руб, с НДС)
    F (5): АКБ 2026
    G (6): Продажи 2026 (руб, с НДС)
    """
    records = []
    for row in sheet.iter_rows(min_row=3, values_only=True):
        if not row or not row[0] and not row[1] and not row[2]:
            continue
        if has_distributor_col:
            distributor = row[0] or ''
            class_tt = row[1] or ''
            sku_name = row[2] or ''
            akb_2025 = row[3] if len(row) > 3 else None
            sales_2025 = row[4] if len(row) > 4 else None
            akb_2026 = row[5] if len(row) > 5 else None
            sales_2026 = row[6] if len(row) > 6 else None
        else:
            distributor = ''
            class_tt = row[0] or ''
            sku_name = row[1] or ''
            akb_2025 = row[2] if len(row) > 2 else None
            sales_2025 = row[3] if len(row) > 3 else None
            akb_2026 = row[4] if len(row) > 4 else None
            sales_2026 = row[5] if len(row) > 5 else None
        
        # Normalize numbers
        def num(v):
            if v is None or v == '' or v == 0:
                return 0
            try:
                return float(v)
            except:
                return 0
        
        def akb_int(v):
            if v is None or v == '':
                return 0
            try:
                return int(v)
            except:
                return 0
        
        if not sku_name:
            continue
        
        # Parse SKU attributes
        attrs = parse_sku(sku_name)
        
        records.append({
            'distributor': str(distributor).strip() if distributor else '',
            'class': str(class_tt).strip() if class_tt else '',
            'sku_name': sku_name.strip(),
            'akb_2025': akb_int(akb_2025),
            'sales_2025': num(sales_2025),
            'akb_2026': akb_int(akb_2026),
            'sales_2026': num(sales_2026),
            'type': attrs['type'],
            'finish': attrs['finish'],
            'color': attrs['color'],
            'color_display': attrs['color_display'],
            'ral': attrs['ral'],
            'volume_l': attrs['volume_l'],
            'volume_kg': attrs['volume_kg'],
            'volume_display': attrs['volume_display'],
        })
    return records


# Read all 3 sheets
sheet1 = wb['БРОЗЭКС, Рыгалов']  # has distributor column
sheet2 = wb['Рыгалов С.В.']      # has distributor column
sheet3 = wb['Общий по точкам']   # NO distributor column (just class)

records_brozeks_rygalov = read_sheet(sheet1, has_distributor_col=True)
records_rygalov = read_sheet(sheet2, has_distributor_col=True)
records_total = read_sheet(sheet3, has_distributor_col=False)

print(f"Sheet 1 'БРОЗЭКС, Рыгалов': {len(records_brozeks_rygalov)} rows")
print(f"Sheet 2 'Рыгалов С.В.':     {len(records_rygalov)} rows")
print(f"Sheet 3 'Общий по точкам':  {len(records_total)} rows")

# Aggregate per sheet
def aggregate(records):
    """Aggregate totals per sheet."""
    total_2025 = sum(r['sales_2025'] for r in records)
    total_2026 = sum(r['sales_2026'] for r in records)
    total_akb_2025 = sum(r['akb_2025'] for r in records)
    total_akb_2026 = sum(r['akb_2026'] for r in records)
    return {
        'rows': len(records),
        'total_2025': total_2025,
        'total_2026': total_2026,
        'growth_pct': ((total_2026 - total_2025) / total_2025 * 100) if total_2025 > 0 else 0,
        'akb_2025': total_akb_2025,
        'akb_2026': total_akb_2026,
        'unique_skus': len(set(r['sku_name'] for r in records)),
        'unique_colors': len(set(r['color_display'] for r in records if r['color_display'])),
        'unique_volumes': len(set(r['volume_display'] for r in records if r['volume_display'])),
    }

agg1 = aggregate(records_brozeks_rygalov)
agg2 = aggregate(records_rygalov)
agg3 = aggregate(records_total)

print(f"\nAggregates:")
print(f"  БРОЗЭКС+Рыгалов: 2025={agg1['total_2025']:,.0f} ₽, 2026={agg1['total_2026']:,.0f} ₽, рост={agg1['growth_pct']:.1f}%")
print(f"  Рыгалов:          2025={agg2['total_2025']:,.0f} ₽, 2026={agg2['total_2026']:,.0f} ₽, рост={agg2['growth_pct']:.1f}%")
print(f"  Общий:            2025={agg3['total_2025']:,.0f} ₽, 2026={agg3['total_2026']:,.0f} ₽, рост={agg3['growth_pct']:.1f}%")

# Save parsed data to JSON for use by HTML dashboard
data_payload = {
    'brozeks_rygalov': records_brozeks_rygalov,
    'rygalov': records_rygalov,
    'total': records_total,
    'aggregates': {
        'brozeks_rygalov': agg1,
        'rygalov': agg2,
        'total': agg3,
    },
}

OUT_JSON = '/home/z/my-project/download/titan_distrib_data.json'
with open(OUT_JSON, 'w', encoding='utf-8') as f:
    json.dump(data_payload, f, ensure_ascii=False, indent=2)
print(f"\nJSON saved: {OUT_JSON} ({os.path.getsize(OUT_JSON)} bytes)")

# ── Quick top colors/volumes/types summary ──────────────────────────────
def agg_by(records, key):
    d_2025 = defaultdict(float)
    d_2026 = defaultdict(float)
    for r in records:
        k = r[key] or '—'
        d_2025[k] += r['sales_2025']
        d_2026[k] += r['sales_2026']
    return d_2025, d_2026

# Top colors for total
print("\n=== TOP COLORS (total) ===")
c25, c26 = agg_by(records_total, 'color_display')
for c, v in sorted(c26.items(), key=lambda x: -x[1])[:8]:
    print(f"  {c:30s} 2025={c25[c]:>12,.0f}  2026={v:>12,.0f}  Δ={(v-c25[c])/max(c25[c],1)*100:+.0f}%")

print("\n=== TOP VOLUMES (total) ===")
v25, v26 = agg_by(records_total, 'volume_display')
for v, val in sorted(v26.items(), key=lambda x: -x[1])[:6]:
    print(f"  {v:10s} 2025={v25[v]:>12,.0f}  2026={val:>12,.0f}  Δ={(val-v25[v])/max(v25[v],1)*100:+.0f}%")

print("\n=== TYPES (total) ===")
t25, t26 = agg_by(records_total, 'type')
for t, val in sorted(t26.items(), key=lambda x: -x[1]):
    print(f"  {t:35s} 2025={t25[t]:>12,.0f}  2026={val:>12,.0f}  Δ={(val-t25[t])/max(t25[t],1)*100:+.0f}%")

print("\n=== CLASS TT (total) ===")
cl25, cl26 = agg_by(records_total, 'class')
for cl, val in sorted(cl26.items(), key=lambda x: -x[1]):
    print(f"  Класс {cl:5s} 2025={cl25[cl]:>12,.0f}  2026={val:>12,.0f}  Δ={(val-cl25[cl])/max(cl25[cl],1)*100:+.0f}%")

print("\n=== TOP SKUS 2026 (total) ===")
sk25, sk26 = agg_by(records_total, 'sku_name')
for sk, val in sorted(sk26.items(), key=lambda x: -x[1])[:10]:
    print(f"  {sk[:60]:60s} 2025={sk25[sk]:>10,.0f}  2026={val:>10,.0f}")
