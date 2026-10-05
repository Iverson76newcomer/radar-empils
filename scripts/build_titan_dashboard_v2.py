"""
TITAN distributors dashboard v2 — focused on 4 key comparisons:
1. БРОЗЭКС — АКБ 2025 vs 2026
2. Рыгалов — АКБ 2025 vs 2026
3. БРОЗЭКС — товарооборот 2025 vs 2026
4. Рыгалов — товарооборот 2025 vs 2026

Plus full breakdown by SKU/colors/volumes for each distributor.
"""
import openpyxl
import json
import re
import os
from collections import defaultdict

SRC = "/home/z/my-project/upload/Титан запрос Заколодний Роман.xlsx"
OUT_HTML = '/home/z/my-project/download/TITAN_дистрибьюторы_дашборд.html'

# ── SKU parsing ────────────────────────────────────────────────────────
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
RAL_RE = re.compile(r'RAL\s*(\d{4})', re.IGNORECASE)
VOL_L_RE = re.compile(r'(\d+(?:\.\d+)?)\s*л\b', re.IGNORECASE)
VOL_KG_RE = re.compile(r'(\d+(?:\.\d+)?)\s*кг\b', re.IGNORECASE)
TYPE_PATTERNS = [
    ('Грунт-эмаль быстросохнущая', re.compile(r'^грунт-эмаль\s+быстросохнущая', re.IGNORECASE)),
    ('Грунт-эмаль по ржавчине', re.compile(r'^грунт-эмаль\s+по\s+ржавчине', re.IGNORECASE)),
    ('Эмаль с молотковым эффектом', re.compile(r'^эмаль\s+с\s+молотковым\s+эффектом', re.IGNORECASE)),
    ('Эмаль', re.compile(r'^эмаль', re.IGNORECASE)),
    ('Грунт-эмаль (проч.)', re.compile(r'^грунт-эмаль', re.IGNORECASE)),
]
FINISH_PATTERNS = {
    'МАТОВАЯ': ['матовая', 'матовый', 'матовой'],
    'ГЛЯНЦЕВАЯ': ['глянцевая', 'глянцевый'],
    'ПОЛУМАТОВАЯ': ['полуматовая'],
}

def parse_sku(name):
    if not name:
        return {'type': '', 'finish': '', 'color': '', 'ral': '',
                'color_display': '', 'volume_l': '', 'volume_kg': '',
                'volume_display': ''}
    n = name.strip()
    n_lower = n.lower()
    sku_type = ''
    for t, pat in TYPE_PATTERNS:
        if pat.search(n):
            sku_type = t
            break
    finish = ''
    for fname, patterns in FINISH_PATTERNS.items():
        if any(p in n_lower for p in patterns):
            finish = fname
            break
    color = ''
    for cname, patterns in COLOR_PATTERNS.items():
        if any(p in n_lower for p in patterns):
            color = cname
            break
    ral_match = RAL_RE.search(n)
    ral = f'RAL {ral_match.group(1)}' if ral_match else ''
    if ral and color:
        color_display = f'{color} ({ral})'
    elif ral and not color:
        color_display = ral
    else:
        color_display = color
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
        'type': sku_type, 'finish': finish, 'color': color, 'ral': ral,
        'color_display': color_display, 'volume_l': vol_l, 'volume_kg': vol_kg,
        'volume_display': volume_display,
    }


def num(v):
    if v is None or v == '' or v == 0: return 0
    try: return float(v)
    except: return 0

def akb_int(v):
    if v is None or v == '': return 0
    try: return int(v)
    except: return 0

def read_sheet(sheet, has_distributor_col=True):
    """Read sheet — skip 'Итог'/'Общий итог' rows."""
    records = []
    for row in sheet.iter_rows(min_row=3, values_only=True):
        if not row or (not row[0] and not row[1] and not row[2]):
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

        # Skip "Итог" rows (where sku_name is None and one of distributor/class is "Итог"/"Общий итог")
        if not sku_name or 'Итог' in str(distributor) or 'Итог' in str(class_tt):
            continue
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


wb = openpyxl.load_workbook(SRC, data_only=True)
records_brozeks = read_sheet(wb['БРОЗЭКС'], has_distributor_col=True)
records_rygalov = read_sheet(wb['Рыгалов С.В.'], has_distributor_col=True)
records_total = read_sheet(wb['Общий по точкам'], has_distributor_col=False)

# Use the "Итог" rows for accurate totals
brozeks_totals = {'akb_2025': 108, 'sales_2025': 2014795.50,
                  'akb_2026': 125, 'sales_2026': 3439930.93}
rygalov_totals = {'akb_2025': 82, 'sales_2025': 2555783.10,
                  'akb_2026': 100, 'sales_2026': 2690941.70}
total_totals = {'akb_2025': 640, 'sales_2025': 25589741.50,
                'akb_2026': 714, 'sales_2026': 21890360.30}

def aggregate(records):
    total_2025 = sum(r['sales_2025'] for r in records)
    total_2026 = sum(r['sales_2026'] for r in records)
    return {
        'rows': len(records),
        'total_2025': total_2025,
        'total_2026': total_2026,
        'growth_pct': ((total_2026 - total_2025) / total_2025 * 100) if total_2025 > 0 else 0,
        'akb_2025': sum(r['akb_2025'] for r in records),
        'akb_2026': sum(r['akb_2026'] for r in records),
        'unique_skus': len(set(r['sku_name'] for r in records)),
        'unique_colors': len(set(r['color_display'] for r in records if r['color_display'])),
        'unique_volumes': len(set(r['volume_display'] for r in records if r['volume_display'])),
    }

data_payload = {
    'brozeks': records_brozeks,
    'rygalov': records_rygalov,
    'total': records_total,
    'aggregates': {
        'brozeks': {**aggregate(records_brozeks), **brozeks_totals},
        'rygalov': {**aggregate(records_rygalov), **rygalov_totals},
        'total': {**aggregate(records_total), **total_totals},
    },
}

data_json = json.dumps(data_payload, ensure_ascii=False)

print(f"Records: brozeks={len(records_brozeks)}, rygalov={len(records_rygalov)}, total={len(records_total)}")

# ── Build HTML ──────────────────────────────────────────────────────────
HTML = r'''<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<title>TITAN · Сравнение БРОЗЭКС vs Рыгалов</title>
<script src="https://cdn.jsdelivr.net/npm/echarts@5.4.3/dist/echarts.min.js"></script>
<style>
  :root {
    --bg: #F5F2EC;
    --card: #FFFFFF;
    --border: #D9D2C2;
    --divider: #EFE9DD;
    --ink: #2F3A48;
    --mute: #8A95A3;
    --red: #C00000;
    --sand: #C9A77C;
    --green: #8BAA8A;
    --blue: #6B8FA8;
    --amber: #D4A574;
  }
  * { box-sizing: border-box; }
  body {
    margin: 0; padding: 24px;
    font-family: -apple-system, "Segoe UI", Inter, Roboto, sans-serif;
    background: var(--bg);
    color: var(--ink);
  }
  .header {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 24px 28px;
    margin-bottom: 24px;
  }
  .header-row { display: flex; justify-content: space-between; align-items: flex-end; }
  .header-title { font-size: 22px; font-weight: 700; letter-spacing: -0.2px; }
  .header-sub { font-size: 13px; color: var(--mute); margin-top: 4px; }
  .pill {
    display: inline-block;
    padding: 6px 16px;
    border-radius: 999px;
    background: #FFF;
    border: 1px solid var(--sand);
    color: var(--sand);
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 1.2px;
  }

  /* Section heading */
  .section-heading {
    font-size: 14px;
    font-weight: 700;
    color: var(--red);
    text-transform: uppercase;
    letter-spacing: 1.5px;
    margin: 28px 0 12px 0;
    padding-bottom: 8px;
    border-bottom: 1px solid var(--border);
  }

  /* KPI comparison grid (2 distributors × 2 metrics) */
  .compare-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 16px;
    margin-bottom: 24px;
  }
  .compare-card {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 20px 24px;
    border-left: 4px solid var(--red);
  }
  .compare-card.rygalov { border-left-color: var(--blue); }
  .compare-card-title {
    font-size: 14px;
    color: var(--mute);
    text-transform: uppercase;
    letter-spacing: 1.2px;
    margin-bottom: 8px;
  }
  .compare-card-name {
    font-size: 18px;
    font-weight: 700;
    margin-bottom: 16px;
  }
  .compare-metric {
    display: flex;
    justify-content: space-between;
    align-items: baseline;
    padding: 8px 0;
    border-bottom: 1px dashed var(--divider);
  }
  .compare-metric:last-child { border-bottom: none; }
  .compare-metric-label {
    font-size: 12px;
    color: var(--mute);
  }
  .compare-metric-values {
    font-family: "Roboto Mono", ui-monospace, monospace;
    font-size: 13px;
  }
  .compare-metric-value-2025 { color: var(--mute); }
  .compare-metric-value-2026 { color: var(--ink); font-weight: 700; margin-left: 12px; }
  .compare-metric-delta {
    margin-left: 12px;
    padding: 2px 8px;
    border-radius: 999px;
    font-size: 11px;
    font-weight: 600;
  }
  .delta-up { background: rgba(139,170,138,0.2); color: #5E7E5C; }
  .delta-down { background: rgba(192,0,0,0.12); color: var(--red); }

  .tabs {
    display: flex;
    gap: 8px;
    margin-bottom: 20px;
    background: var(--card);
    padding: 6px;
    border-radius: 10px;
    border: 1px solid var(--border);
    width: fit-content;
  }
  .tab {
    padding: 10px 18px;
    border-radius: 8px;
    cursor: pointer;
    font-size: 13px;
    font-weight: 600;
    color: var(--mute);
    transition: all 0.15s;
  }
  .tab:hover { background: var(--divider); }
  .tab.active {
    background: var(--ink);
    color: white;
  }

  .charts-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 20px;
    margin-bottom: 20px;
  }
  .charts-full { grid-template-columns: 1fr; }
  .chart-card {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 16px 20px;
  }
  .chart-title { font-size: 14px; font-weight: 700; margin-bottom: 4px; }
  .chart-sub { font-size: 11px; color: var(--mute); margin-bottom: 12px; }
  .chart-body { width: 100%; height: 320px; }
  .chart-body.tall { height: 420px; }

  .summary {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 24px 28px;
    margin-top: 24px;
  }
  .summary h3 {
    font-size: 14px;
    color: var(--red);
    text-transform: uppercase;
    letter-spacing: 1.5px;
    margin: 0 0 12px 0;
  }
  .summary ul {
    margin: 0; padding-left: 20px;
    font-size: 13px;
    line-height: 1.7;
  }
  .summary li { margin-bottom: 6px; }
  .summary strong { color: var(--ink); }

  .footer {
    margin-top: 30px;
    padding: 16px 0;
    font-size: 11px;
    color: var(--mute);
    text-align: center;
    border-top: 1px solid var(--border);
  }
</style>
</head>
<body>

<div class="header">
  <div class="header-row">
    <div>
      <div class="header-title">TITAN · Сравнение БРОЗЭКС vs Рыгалов</div>
      <div class="header-sub">АКБ и товарооборот · 2025 vs 2026 · Период: год к году</div>
    </div>
    <div class="pill">ОТЧЁТ · 02.10.2026</div>
  </div>
</div>

<!-- KEY COMPARISON: 2x2 grid -->
<div class="section-heading">Сравнение по 4 ключевым показателям</div>
<div class="compare-grid">
  <!-- Card 1: БРОЗЭКС АКБ -->
  <div class="compare-card">
    <div class="compare-card-title">Показатель 1</div>
    <div class="compare-card-name">БРОЗЭКС — АКБ (активная клиентская база)</div>
    <div class="compare-metric">
      <span class="compare-metric-label">2025</span>
      <span class="compare-metric-values">
        <span class="compare-metric-value-2025" id="brozeks_akb_2025">—</span>
      </span>
    </div>
    <div class="compare-metric">
      <span class="compare-metric-label">2026</span>
      <span class="compare-metric-values">
        <span class="compare-metric-value-2026" id="brozeks_akb_2026">—</span>
        <span class="compare-metric-delta" id="brozeks_akb_delta">—</span>
      </span>
    </div>
  </div>

  <!-- Card 2: Рыгалов АКБ -->
  <div class="compare-card rygalov">
    <div class="compare-card-title">Показатель 2</div>
    <div class="compare-card-name">Рыгалов — АКБ (активная клиентская база)</div>
    <div class="compare-metric">
      <span class="compare-metric-label">2025</span>
      <span class="compare-metric-values">
        <span class="compare-metric-value-2025" id="rygalov_akb_2025">—</span>
      </span>
    </div>
    <div class="compare-metric">
      <span class="compare-metric-label">2026</span>
      <span class="compare-metric-values">
        <span class="compare-metric-value-2026" id="rygalov_akb_2026">—</span>
        <span class="compare-metric-delta" id="rygalov_akb_delta">—</span>
      </span>
    </div>
  </div>

  <!-- Card 3: БРОЗЭКС товарооборот -->
  <div class="compare-card">
    <div class="compare-card-title">Показатель 3</div>
    <div class="compare-card-name">БРОЗЭКС — Товарооборот (₽, с НДС)</div>
    <div class="compare-metric">
      <span class="compare-metric-label">2025</span>
      <span class="compare-metric-values">
        <span class="compare-metric-value-2025" id="brozeks_sales_2025">—</span>
      </span>
    </div>
    <div class="compare-metric">
      <span class="compare-metric-label">2026</span>
      <span class="compare-metric-values">
        <span class="compare-metric-value-2026" id="brozeks_sales_2026">—</span>
        <span class="compare-metric-delta" id="brozeks_sales_delta">—</span>
      </span>
    </div>
  </div>

  <!-- Card 4: Рыгалов товарооборот -->
  <div class="compare-card rygalov">
    <div class="compare-card-title">Показатель 4</div>
    <div class="compare-card-name">Рыгалов — Товарооборот (₽, с НДС)</div>
    <div class="compare-metric">
      <span class="compare-metric-label">2025</span>
      <span class="compare-metric-values">
        <span class="compare-metric-value-2025" id="rygalov_sales_2025">—</span>
      </span>
    </div>
    <div class="compare-metric">
      <span class="compare-metric-label">2026</span>
      <span class="compare-metric-values">
        <span class="compare-metric-value-2026" id="rygalov_sales_2026">—</span>
        <span class="compare-metric-delta" id="rygalov_sales_delta">—</span>
      </span>
    </div>
  </div>
</div>

<!-- Quick comparison charts -->
<div class="section-heading">Визуальное сравнение</div>
<div class="charts-grid">
  <div class="chart-card">
    <div class="chart-title">АКБ 2025 vs 2026 по дистрибьюторам</div>
    <div class="chart-sub">Количество активных клиентов</div>
    <div id="chart_akb_compare" class="chart-body"></div>
  </div>
  <div class="chart-card">
    <div class="chart-title">Товарооборот 2025 vs 2026 по дистрибьюторам</div>
    <div class="chart-sub">₽ с НДС</div>
    <div id="chart_sales_compare" class="chart-body"></div>
  </div>
</div>
<div class="charts-grid">
  <div class="chart-card">
    <div class="chart-title">Динамика АКБ (%)</div>
    <div class="chart-sub">Процент изменения 2025 → 2026</div>
    <div id="chart_akb_growth" class="chart-body"></div>
  </div>
  <div class="chart-card">
    <div class="chart-title">Динамика товарооборота (%)</div>
    <div class="chart-sub">Процент изменения 2025 → 2026</div>
    <div id="chart_sales_growth" class="chart-body"></div>
  </div>
</div>

<!-- Detailed analysis tabs -->
<div class="section-heading">Детальный анализ по дистрибьютору</div>
<div class="tabs">
  <div class="tab active" data-view="brozeks">🏪 БРОЗЭКС</div>
  <div class="tab" data-view="rygalov">👤 Рыгалов С.В.</div>
  <div class="tab" data-view="compare_sku">⚖️ Сравнение топ-SKU</div>
</div>

<div id="content"></div>

<div class="summary">
  <h3>Ключевые выводы</h3>
  <ul id="insights">—</ul>
</div>

<div class="footer">
  Источник данных: Титан запрос Заколодний Роман.xlsx (3 вкладки) · Анализ подготовлен 02.10.2026
</div>

<script>
const DATA = __DATA_PLACEHOLDER__;

const palette = {
  red: '#C00000', sand: '#C9A77C', green: '#8BAA8A', blue: '#6B8FA8',
  amber: '#D4A574', mute: '#8A95A3', ink: '#2F3A48', gray: '#D9D2C2',
};

function fmt(n) {
  if (!n) return '0';
  if (n >= 1e6) return (n/1e6).toFixed(2) + ' млн';
  if (n >= 1e3) return (n/1e3).toFixed(0) + ' тыс';
  return Math.round(n).toString();
}
function fmtRub(n) {
  return Math.round(n).toLocaleString('ru-RU') + ' ₽';
}
function fmtPct(n) {
  return (n >= 0 ? '+' : '') + n.toFixed(1) + '%';
}
function deltaPct(curr, prev) {
  return (curr - prev) / Math.max(prev, 1) * 100;
}

// Init KPI cards from totals (use exact "Итог" values)
function initKPIs() {
  const br = DATA.aggregates.brozeks;
  const rg = DATA.aggregates.rygalov;

  // БРОЗЭКС АКБ
  document.getElementById('brozeks_akb_2025').textContent = br.akb_2025;
  document.getElementById('brozeks_akb_2026').textContent = br.akb_2026;
  const brAkbDelta = deltaPct(br.akb_2026, br.akb_2025);
  const brAkbEl = document.getElementById('brozeks_akb_delta');
  brAkbEl.textContent = fmtPct(brAkbDelta);
  brAkbEl.className = 'compare-metric-delta ' + (brAkbDelta >= 0 ? 'delta-up' : 'delta-down');

  // Рыгалов АКБ
  document.getElementById('rygalov_akb_2025').textContent = rg.akb_2025;
  document.getElementById('rygalov_akb_2026').textContent = rg.akb_2026;
  const rgAkbDelta = deltaPct(rg.akb_2026, rg.akb_2025);
  const rgAkbEl = document.getElementById('rygalov_akb_delta');
  rgAkbEl.textContent = fmtPct(rgAkbDelta);
  rgAkbEl.className = 'compare-metric-delta ' + (rgAkbDelta >= 0 ? 'delta-up' : 'delta-down');

  // БРОЗЭКС sales
  document.getElementById('brozeks_sales_2025').textContent = fmtRub(br.sales_2025);
  document.getElementById('brozeks_sales_2026').textContent = fmtRub(br.sales_2026);
  const brSalesDelta = deltaPct(br.sales_2026, br.sales_2025);
  const brSalesEl = document.getElementById('brozeks_sales_delta');
  brSalesEl.textContent = fmtPct(brSalesDelta);
  brSalesEl.className = 'compare-metric-delta ' + (brSalesDelta >= 0 ? 'delta-up' : 'delta-down');

  // Рыгалов sales
  document.getElementById('rygalov_sales_2025').textContent = fmtRub(rg.sales_2025);
  document.getElementById('rygalov_sales_2026').textContent = fmtRub(rg.sales_2026);
  const rgSalesDelta = deltaPct(rg.sales_2026, rg.sales_2025);
  const rgSalesEl = document.getElementById('rygalov_sales_delta');
  rgSalesEl.textContent = fmtPct(rgSalesDelta);
  rgSalesEl.className = 'compare-metric-delta ' + (rgSalesDelta >= 0 ? 'delta-up' : 'delta-down');
}

// Comparison charts (top section)
function renderCompareCharts() {
  const br = DATA.aggregates.brozeks;
  const rg = DATA.aggregates.rygalov;

  // АКБ comparison
  echarts.init(document.getElementById('chart_akb_compare')).setOption({
    tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' },
               formatter: function(p) {
                 let s = p[0].name + '<br/>';
                 p.forEach(function(item) { s += item.marker + ' ' + item.seriesName + ': ' + item.value + '<br/>'; });
                 return s;
               }},
    legend: { data: ['2025', '2026'], top: 0, right: 0 },
    grid: { left: 60, right: 30, top: 40, bottom: 40 },
    xAxis: { type: 'category', data: ['БРОЗЭКС', 'Рыгалов'], axisLabel: { color: palette.ink, fontSize: 12 } },
    yAxis: { type: 'value', axisLabel: { color: palette.mute, fontSize: 10 } },
    series: [
      { name: '2025', type: 'bar', data: [br.akb_2025, rg.akb_2025], itemStyle: { color: palette.mute }, label: { show: true, position: 'top', color: palette.ink, fontSize: 11, fontWeight: 600 } },
      { name: '2026', type: 'bar', data: [br.akb_2026, rg.akb_2026], itemStyle: { color: palette.red }, label: { show: true, position: 'top', color: palette.ink, fontSize: 11, fontWeight: 600 } },
    ],
  });

  // Sales comparison
  echarts.init(document.getElementById('chart_sales_compare')).setOption({
    tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' },
               formatter: function(p) {
                 let s = p[0].name + '<br/>';
                 p.forEach(function(item) { s += item.marker + ' ' + item.seriesName + ': ' + fmtRub(item.value) + '<br/>'; });
                 return s;
               }},
    legend: { data: ['2025', '2026'], top: 0, right: 0 },
    grid: { left: 80, right: 30, top: 40, bottom: 40 },
    xAxis: { type: 'category', data: ['БРОЗЭКС', 'Рыгалов'], axisLabel: { color: palette.ink, fontSize: 12 } },
    yAxis: { type: 'value', axisLabel: { color: palette.mute, fontSize: 10, formatter: function(v) { return fmt(v); } } },
    series: [
      { name: '2025', type: 'bar', data: [br.sales_2025, rg.sales_2025], itemStyle: { color: palette.mute }, label: { show: true, position: 'top', color: palette.ink, fontSize: 10, fontWeight: 600, formatter: function(p) { return fmt(p.value); } } },
      { name: '2026', type: 'bar', data: [br.sales_2026, rg.sales_2026], itemStyle: { color: palette.red }, label: { show: true, position: 'top', color: palette.ink, fontSize: 10, fontWeight: 600, formatter: function(p) { return fmt(p.value); } } },
    ],
  });

  // АКБ growth %
  const brAkbGrowth = deltaPct(br.akb_2026, br.akb_2025);
  const rgAkbGrowth = deltaPct(rg.akb_2026, rg.akb_2025);
  echarts.init(document.getElementById('chart_akb_growth')).setOption({
    tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' },
               formatter: function(p) { return p[0].name + ': ' + fmtPct(p[0].value); } },
    grid: { left: 60, right: 80, top: 20, bottom: 40 },
    xAxis: { type: 'category', data: ['БРОЗЭКС', 'Рыгалов'], axisLabel: { color: palette.ink, fontSize: 12 } },
    yAxis: { type: 'value', axisLabel: { color: palette.mute, fontSize: 10, formatter: function(v) { return v + '%'; } } },
    series: [{
      type: 'bar',
      data: [
        { value: brAkbGrowth, itemStyle: { color: brAkbGrowth >= 0 ? palette.green : palette.red } },
        { value: rgAkbGrowth, itemStyle: { color: rgAkbGrowth >= 0 ? palette.green : palette.red } },
      ],
      label: { show: true, position: 'top', formatter: function(p) { return fmtPct(p.value); }, color: palette.ink, fontSize: 12, fontWeight: 700 },
    }],
  });

  // Sales growth %
  const brSalesGrowth = deltaPct(br.sales_2026, br.sales_2025);
  const rgSalesGrowth = deltaPct(rg.sales_2026, rg.sales_2025);
  echarts.init(document.getElementById('chart_sales_growth')).setOption({
    tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' },
               formatter: function(p) { return p[0].name + ': ' + fmtPct(p[0].value); } },
    grid: { left: 60, right: 80, top: 20, bottom: 40 },
    xAxis: { type: 'category', data: ['БРОЗЭКС', 'Рыгалов'], axisLabel: { color: palette.ink, fontSize: 12 } },
    yAxis: { type: 'value', axisLabel: { color: palette.mute, fontSize: 10, formatter: function(v) { return v + '%'; } } },
    series: [{
      type: 'bar',
      data: [
        { value: brSalesGrowth, itemStyle: { color: brSalesGrowth >= 0 ? palette.green : palette.red } },
        { value: rgSalesGrowth, itemStyle: { color: rgSalesGrowth >= 0 ? palette.green : palette.red } },
      ],
      label: { show: true, position: 'top', formatter: function(p) { return fmtPct(p.value); }, color: palette.ink, fontSize: 12, fontWeight: 700 },
    }],
  });
}

// Helpers for detailed analysis
function aggregateBy(records, key) {
  const m25 = {}, m26 = {};
  records.forEach(function(r) {
    const k = r[key] || '—';
    m25[k] = (m25[k] || 0) + r.sales_2025;
    m26[k] = (m26[k] || 0) + r.sales_2026;
  });
  const keys = Array.from(new Set(Object.keys(m25).concat(Object.keys(m26))));
  const arr = keys.map(function(k) {
    const v25 = m25[k] || 0;
    const v26 = m26[k] || 0;
    return { key: k, v25: v25, v26: v26, delta: (v26 - v25) / Math.max(v25, 1) * 100 };
  });
  arr.sort(function(a, b) { return b.v26 - a.v26; });
  return arr;
}

const charts = [];
function disposeCharts() {
  charts.forEach(function(c) { c.dispose(); });
  charts.length = 0;
}
function renderChart(domId, option) {
  const dom = document.getElementById(domId);
  if (!dom) return null;
  const chart = echarts.init(dom);
  chart.setOption(option);
  charts.push(chart);
  return chart;
}
function tooltipRub() {
  return {
    trigger: 'axis', axisPointer: { type: 'shadow' },
    formatter: function(p) {
      let s = p[0].name + '<br/>';
      p.forEach(function(item) { s += item.marker + ' ' + item.seriesName + ': ' + fmtRub(item.value) + '<br/>'; });
      return s;
    }
  };
}

function renderDistributor(viewKey) {
  const records = DATA[viewKey];
  const title = viewKey === 'brozeks' ? 'БРОЗЭКС' : 'Рыгалов С.В.';
  document.getElementById('content').innerHTML =
    '<div class="charts-grid">' +
      '<div class="chart-card">' +
        '<div class="chart-title">' + title + ': топ-10 SKU по выручке 2026</div>' +
        '<div class="chart-sub">С разбивкой по выручке 2025</div>' +
        '<div id="chart_d_sku" class="chart-body tall"></div>' +
      '</div>' +
      '<div class="chart-card">' +
        '<div class="chart-title">Цветовая структура</div>' +
        '<div class="chart-sub">Топ-8 цветов по выручке 2026</div>' +
        '<div id="chart_d_color" class="chart-body"></div>' +
      '</div>' +
    '</div>' +
    '<div class="charts-grid">' +
      '<div class="chart-card">' +
        '<div class="chart-title">Распределение по фасовкам</div>' +
        '<div class="chart-sub">2025 vs 2026, ₽</div>' +
        '<div id="chart_d_vol" class="chart-body"></div>' +
      '</div>' +
      '<div class="chart-card">' +
        '<div class="chart-title">Структура по типам продуктов</div>' +
        '<div class="chart-sub">2026, ₽</div>' +
        '<div id="chart_d_type" class="chart-body"></div>' +
      '</div>' +
    '</div>';

  const skuData = aggregateBy(records, 'sku_name').slice(0, 10).reverse();
  renderChart('chart_d_sku', {
    tooltip: tooltipRub(),
    legend: { data: ['2025', '2026'], top: 0, right: 0 },
    grid: { left: 380, right: 50, top: 40, bottom: 30 },
    xAxis: { type: 'value', axisLabel: { color: palette.mute, fontSize: 10, formatter: function(v) { return fmt(v); } } },
    yAxis: { type: 'category',
             data: skuData.map(function(d) { return d.key.length > 55 ? d.key.slice(0,55)+'…' : d.key; }),
             axisLabel: { color: palette.ink, fontSize: 9 } },
    series: [
      { name: '2025', type: 'bar', data: skuData.map(function(d) { return d.v25; }), itemStyle: { color: palette.mute } },
      { name: '2026', type: 'bar', data: skuData.map(function(d) { return d.v26; }), itemStyle: { color: palette.red } },
    ],
  });

  const colorData = aggregateBy(records, 'color_display').filter(function(r) { return r.key && r.key !== '—'; }).slice(0, 8);
  renderChart('chart_d_color', {
    tooltip: tooltipRub(),
    legend: { data: ['2025', '2026'], top: 0, right: 0 },
    grid: { left: 80, right: 20, top: 40, bottom: 100 },
    xAxis: { type: 'category', data: colorData.map(function(d) { return d.key; }),
             axisLabel: { rotate: 30, color: palette.ink, fontSize: 9 } },
    yAxis: { type: 'value', axisLabel: { color: palette.mute, fontSize: 10, formatter: function(v) { return fmt(v); } } },
    series: [
      { name: '2025', type: 'bar', data: colorData.map(function(d) { return d.v25; }), itemStyle: { color: palette.mute } },
      { name: '2026', type: 'bar', data: colorData.map(function(d) { return d.v26; }), itemStyle: { color: palette.sand } },
    ],
  });

  const volData = aggregateBy(records, 'volume_display').filter(function(r) { return r.key && r.key !== '—'; }).slice(0, 8);
  renderChart('chart_d_vol', {
    tooltip: tooltipRub(),
    legend: { data: ['2025', '2026'], top: 0, right: 0 },
    grid: { left: 80, right: 20, top: 40, bottom: 80 },
    xAxis: { type: 'category', data: volData.map(function(d) { return d.key; }),
             axisLabel: { color: palette.ink, fontSize: 10 } },
    yAxis: { type: 'value', axisLabel: { color: palette.mute, fontSize: 10, formatter: function(v) { return fmt(v); } } },
    series: [
      { name: '2025', type: 'bar', data: volData.map(function(d) { return d.v25; }), itemStyle: { color: palette.mute } },
      { name: '2026', type: 'bar', data: volData.map(function(d) { return d.v26; }), itemStyle: { color: palette.blue } },
    ],
  });

  const typeData = aggregateBy(records, 'type').filter(function(r) { return r.key !== '—'; });
  renderChart('chart_d_type', {
    tooltip: { trigger: 'item', formatter: function(p) { return p.name + '<br/>' + fmtRub(p.value); } },
    legend: { bottom: 0, left: 'center' },
    series: [{
      type: 'pie', radius: ['40%', '70%'], center: ['50%', '45%'],
      data: typeData.map(function(d, i) {
        return { name: d.key, value: d.v26, itemStyle: { color: ['#C00000', '#C9A77C', '#6B8FA8', '#8BAA8A', '#D4A574'][i % 5] } };
      }),
      label: { color: palette.ink, fontSize: 10, formatter: '{b}\n{d}%' },
    }],
  });
}

function renderCompareSku() {
  document.getElementById('content').innerHTML =
    '<div class="charts-grid">' +
      '<div class="chart-card">' +
        '<div class="chart-title">Топ-15 SKU: БРОЗЭКС vs Рыгалов (2026)</div>' +
        '<div class="chart-sub">Сравнение выручки 2026, ₽</div>' +
        '<div id="chart_cmp_sku" class="chart-body tall"></div>' +
      '</div>' +
      '<div class="chart-card">' +
        '<div class="chart-title">Топ-8 цветов: БРОЗЭКС vs Рыгалов (2026)</div>' +
        '<div class="chart-sub">Сравнение выручки 2026, ₽</div>' +
        '<div id="chart_cmp_color" class="chart-body"></div>' +
      '</div>' +
    '</div>';

  // Top SKU comparison
  const brSku = aggregateBy(DATA.brozeks, 'sku_name').slice(0, 15);
  const rgSku = aggregateBy(DATA.rygalov, 'sku_name');
  const rgMap = {};
  rgSku.forEach(function(d) { rgMap[d.key] = d.v26; });
  const cmpSku = brSku.map(function(d) {
    return { name: d.key, br: d.v26, rg: rgMap[d.key] || 0 };
  }).sort(function(a, b) { return (b.br + b.rg) - (a.br + a.rg); }).reverse();

  renderChart('chart_cmp_sku', {
    tooltip: tooltipRub(),
    legend: { data: ['БРОЗЭКС 2026', 'Рыгалов 2026'], top: 0, right: 0 },
    grid: { left: 380, right: 50, top: 40, bottom: 30 },
    xAxis: { type: 'value', axisLabel: { color: palette.mute, fontSize: 10, formatter: function(v) { return fmt(v); } } },
    yAxis: { type: 'category',
             data: cmpSku.map(function(d) { return d.name.length > 55 ? d.name.slice(0,55)+'…' : d.name; }),
             axisLabel: { color: palette.ink, fontSize: 9 } },
    series: [
      { name: 'БРОЗЭКС 2026', type: 'bar', data: cmpSku.map(function(d) { return d.br; }), itemStyle: { color: palette.red } },
      { name: 'Рыгалов 2026', type: 'bar', data: cmpSku.map(function(d) { return d.rg; }), itemStyle: { color: palette.blue } },
    ],
  });

  // Top colors comparison
  const brColor = aggregateBy(DATA.brozeks, 'color_display').filter(function(r) { return r.key && r.key !== '—'; }).slice(0, 8);
  const rgColor = aggregateBy(DATA.rygalov, 'color_display');
  const rgColorMap = {};
  rgColor.forEach(function(d) { rgColorMap[d.key] = d.v26; });
  const allColors = Array.from(new Set(brColor.map(function(d) { return d.key; }).concat(Object.keys(rgColorMap))));
  const cmpColor = allColors.map(function(k) {
    return { name: k, br: 0, rg: rgColorMap[k] || 0 };
  });
  brColor.forEach(function(d) {
    const item = cmpColor.find(function(c) { return c.name === d.key; });
    if (item) item.br = d.v26;
  });
  cmpColor.sort(function(a, b) { return (b.br + b.rg) - (a.br + a.rg); });
  cmpColor.reverse();

  renderChart('chart_cmp_color', {
    tooltip: tooltipRub(),
    legend: { data: ['БРОЗЭКС 2026', 'Рыгалов 2026'], top: 0, right: 0 },
    grid: { left: 200, right: 50, top: 40, bottom: 30 },
    xAxis: { type: 'value', axisLabel: { color: palette.mute, fontSize: 10, formatter: function(v) { return fmt(v); } } },
    yAxis: { type: 'category', data: cmpColor.map(function(d) { return d.name; }),
             axisLabel: { color: palette.ink, fontSize: 10 } },
    series: [
      { name: 'БРОЗЭКС 2026', type: 'bar', data: cmpColor.map(function(d) { return d.br; }), itemStyle: { color: palette.red } },
      { name: 'Рыгалов 2026', type: 'bar', data: cmpColor.map(function(d) { return d.rg; }), itemStyle: { color: palette.blue } },
    ],
  });
}

function renderInsights() {
  const br = DATA.aggregates.brozeks;
  const rg = DATA.aggregates.rygalov;
  const insights = [];

  const brAkbGrowth = deltaPct(br.akb_2026, br.akb_2025);
  const rgAkbGrowth = deltaPct(rg.akb_2026, rg.akb_2025);
  const brSalesGrowth = deltaPct(br.sales_2026, br.sales_2025);
  const rgSalesGrowth = deltaPct(rg.sales_2026, rg.sales_2025);

  insights.push('<strong>1. БРОЗЭКС — АКБ:</strong> ' + br.akb_2025 + ' → ' + br.akb_2026 + ' клиентов (' + fmtPct(brAkbGrowth) + ')');
  insights.push('<strong>2. Рыгалов — АКБ:</strong> ' + rg.akb_2025 + ' → ' + rg.akb_2026 + ' клиентов (' + fmtPct(rgAkbGrowth) + ')');
  insights.push('<strong>3. БРОЗЭКС — Товарооборот:</strong> ' + fmtRub(br.sales_2025) + ' → ' + fmtRub(br.sales_2026) + ' (' + fmtPct(brSalesGrowth) + ')');
  insights.push('<strong>4. Рыгалов — Товарооборот:</strong> ' + fmtRub(rg.sales_2025) + ' → ' + fmtRub(rg.sales_2026) + ' (' + fmtPct(rgSalesGrowth) + ')');

  insights.push('<strong>📊 Сравнение АКБ:</strong> Рыгалов растёт быстрее (+22,0%) чем БРОЗЭКС (+15,7%) в относительном выражении. В абсолютном — БРОЗЭКС добавил 17 клиентов, Рыгалов — 18 клиентов.');

  insights.push('<strong>💰 Сравнение товарооборота:</strong> БРОЗЭКС показал взрывной рост +70,7% (+1,43 млн ₽), Рыгалов — умеренный +5,3% (+135 тыс ₽). Средний чек на клиента у БРОЗЭКС вырос сильнее.');

  if (brAkbGrowth > 0 && brSalesGrowth > brAkbGrowth) {
    insights.push('<strong>💡 Вывод по БРОЗЭКС:</strong> Товарооборот растёт быстрее АКБ (+70,7% vs +15,7%) — означает рост среднего чека на клиента. Эффективная работа с существующей базой + расширение.');
  }
  if (rgAkbGrowth > 0 && rgSalesGrowth < rgAkbGrowth) {
    insights.push('<strong>💡 Вывод по Рыгалову:</strong> АКБ растёт быстрее товарооборота (+22,0% vs +5,3%) — новые клиенты дают меньший чек. Требуется работа по увеличению среднего объёма закупок.');
  }

  document.getElementById('insights').innerHTML = insights.map(function(i) { return '<li>' + i + '</li>'; }).join('');
}

document.querySelectorAll('.tab').forEach(function(tab) {
  tab.addEventListener('click', function() {
    document.querySelectorAll('.tab').forEach(function(t) { t.classList.remove('active'); });
    tab.classList.add('active');
    const view = tab.dataset.view;
    disposeCharts();
    if (view === 'compare_sku') {
      renderCompareSku();
    } else {
      renderDistributor(view);
    }
  });
});

initKPIs();
renderCompareCharts();
renderDistributor('brozeks');
renderInsights();

window.addEventListener('resize', function() {
  charts.forEach(function(c) { c.resize(); });
  // Also resize comparison charts (created before charts[] was populated)
  ['chart_akb_compare', 'chart_sales_compare', 'chart_akb_growth', 'chart_sales_growth'].forEach(function(id) {
    const inst = echarts.getInstanceByDom(document.getElementById(id));
    if (inst) inst.resize();
  });
});
</script>

</body>
</html>
'''

# Replace placeholder — escape </script> if present in JSON
data_safe = data_json.replace('</', '<\\/').replace('<script', '<\\script')
HTML = HTML.replace('__DATA_PLACEHOLDER__', data_safe)

with open(OUT_HTML, 'w', encoding='utf-8') as f:
    f.write(HTML)

print(f"HTML dashboard: {OUT_HTML}")
print(f"Size: {os.path.getsize(OUT_HTML)} bytes")
