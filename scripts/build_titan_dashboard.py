"""
Build interactive HTML dashboard for TITAN distributors analysis — FIXED version.

This version:
- Reads Excel file directly (no separate JSON step)
- Embeds parsed data as JS variable
- All template literals use proper JS syntax (no Python :.0f format strings)
- Self-contained HTML (only ECharts CDN dependency)
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
        'type': sku_type, 'finish': finish, 'color': color,
        'ral': ral, 'color_display': color_display,
        'volume_l': vol_l, 'volume_kg': vol_kg,
        'volume_display': volume_display,
    }


# ── Read all 3 sheets ───────────────────────────────────────────────────
wb = openpyxl.load_workbook(SRC, data_only=True)

def read_sheet(sheet, has_distributor_col=True):
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

        def num(v):
            if v is None or v == '' or v == 0:
                return 0
            try: return float(v)
            except: return 0

        def akb_int(v):
            if v is None or v == '': return 0
            try: return int(v)
            except: return 0

        if not sku_name: continue
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


records_brozeks_rygalov = read_sheet(wb['БРОЗЭКС, Рыгалов'], has_distributor_col=True)
records_rygalov = read_sheet(wb['Рыгалов С.В.'], has_distributor_col=True)
records_total = read_sheet(wb['Общий по точкам'], has_distributor_col=False)


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


agg1 = aggregate(records_brozeks_rygalov)
agg2 = aggregate(records_rygalov)
agg3 = aggregate(records_total)

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

data_json = json.dumps(data_payload, ensure_ascii=False)

# ── Build HTML with embedded JSON ───────────────────────────────────────
HTML = r'''<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<title>TITAN · Анализ продаж по дистрибьюторам</title>
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
  h1 { font-size: 28px; margin: 0 0 8px 0; }
  h2 { font-size: 16px; margin: 0 0 12px 0; color: var(--mute); font-weight: 500; }
  .header {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 24px 28px;
    margin-bottom: 24px;
  }
  .header-row {
    display: flex; justify-content: space-between; align-items: flex-end;
  }
  .header-title {
    font-size: 22px; font-weight: 700; letter-spacing: -0.2px;
  }
  .header-sub {
    font-size: 13px; color: var(--mute); margin-top: 4px;
  }
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
  .kpi-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 16px;
    margin-bottom: 24px;
  }
  .kpi-card {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 16px 18px;
    border-left: 4px solid var(--red);
  }
  .kpi-card.accent { border-left-color: var(--sand); }
  .kpi-card.blue { border-left-color: var(--blue); }
  .kpi-card.green { border-left-color: var(--green); }
  .kpi-label {
    font-size: 11px;
    color: var(--mute);
    text-transform: uppercase;
    letter-spacing: 1.2px;
    margin-bottom: 6px;
  }
  .kpi-value {
    font-size: 22px;
    font-weight: 700;
    font-family: "Roboto Mono", ui-monospace, monospace;
    line-height: 1.1;
  }
  .kpi-delta {
    font-size: 11px;
    font-weight: 600;
    margin-top: 4px;
  }
  .delta-up { color: var(--green); }
  .delta-down { color: var(--red); }
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
  .chart-title {
    font-size: 14px;
    font-weight: 700;
    margin-bottom: 4px;
  }
  .chart-sub {
    font-size: 11px;
    color: var(--mute);
    margin-bottom: 12px;
  }
  .chart-body {
    width: 100%;
    height: 320px;
  }
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
      <div class="header-title">TITAN · Анализ продаж по дистрибьюторам</div>
      <div class="header-sub">БРОЗЭКС · И.П. Рыгалов · Общая статистика · Сравнение 2025 vs 2026</div>
    </div>
    <div class="pill">ОТЧЁТ · 02.10.2026</div>
  </div>
</div>

<div class="kpi-grid">
  <div class="kpi-card">
    <div class="kpi-label">Общая выручка 2026</div>
    <div class="kpi-value" id="kpi_total_2026">—</div>
    <div class="kpi-delta" id="kpi_total_delta">—</div>
  </div>
  <div class="kpi-card accent">
    <div class="kpi-label">БРОЗЭКС + Рыгалов 2026</div>
    <div class="kpi-value" id="kpi_brozeks_2026">—</div>
    <div class="kpi-delta" id="kpi_brozeks_delta">—</div>
  </div>
  <div class="kpi-card blue">
    <div class="kpi-label">Рыгалов 2026</div>
    <div class="kpi-value" id="kpi_rygalov_2026">—</div>
    <div class="kpi-delta" id="kpi_rygalov_delta">—</div>
  </div>
  <div class="kpi-card green">
    <div class="kpi-label">Уникальных SKU (общая)</div>
    <div class="kpi-value" id="kpi_skus">—</div>
    <div class="kpi-delta" id="kpi_skus_delta">—</div>
  </div>
</div>

<div class="tabs">
  <div class="tab active" data-view="total">📊 Общий по точкам</div>
  <div class="tab" data-view="brozeks_rygalov">🏪 БРОЗЭКС + Рыгалов</div>
  <div class="tab" data-view="rygalov">👤 Рыгалов С.В.</div>
  <div class="tab" data-view="compare">⚖️ Сравнение</div>
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
// DATA injected from Python — JSON-encoded
const DATA = __DATA_PLACEHOLDER__;

const palette = {
  red: '#C00000', sand: '#C9A77C', green: '#8BAA8A', blue: '#6B8FA8',
  amber: '#D4A574', mute: '#8A95A3', ink: '#2F3A48', gray: '#D9D2C2',
  bg: '#F5F2EC', card: '#FFFFFF', divider: '#EFE9DD',
};

const colorPalette = [
  '#C00000', '#C9A77C', '#6B8FA8', '#8BAA8A', '#D4A574',
  '#B5614A', '#5A6776', '#A7C0D0', '#8B7355', '#8A95A3',
  '#D9D2C2', '#C7333E', '#5E8B5E', '#8B4513', '#4A6FA5',
];

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
    return {
      key: k,
      v25: v25,
      v26: v26,
      delta: (v26 - v25) / Math.max(v25, 1) * 100,
    };
  });
  arr.sort(function(a, b) { return b.v26 - a.v26; });
  return arr;
}

function initKPIs() {
  const tot = DATA.aggregates.total;
  const br = DATA.aggregates.brozeks_rygalov;
  const rg = DATA.aggregates.rygalov;

  document.getElementById('kpi_total_2026').textContent = fmtRub(tot.total_2026);
  const totDelta = tot.growth_pct;
  const totDeltaEl = document.getElementById('kpi_total_delta');
  totDeltaEl.textContent = fmtPct(totDelta) + ' к 2025';
  totDeltaEl.className = 'kpi-delta ' + (totDelta >= 0 ? 'delta-up' : 'delta-down');

  document.getElementById('kpi_brozeks_2026').textContent = fmtRub(br.total_2026);
  const brDelta = br.growth_pct;
  const brDeltaEl = document.getElementById('kpi_brozeks_delta');
  brDeltaEl.textContent = fmtPct(brDelta) + ' к 2025';
  brDeltaEl.className = 'kpi-delta ' + (brDelta >= 0 ? 'delta-up' : 'delta-down');

  document.getElementById('kpi_rygalov_2026').textContent = fmtRub(rg.total_2026);
  const rgDelta = rg.growth_pct;
  const rgDeltaEl = document.getElementById('kpi_rygalov_delta');
  rgDeltaEl.textContent = fmtPct(rgDelta) + ' к 2025';
  rgDeltaEl.className = 'kpi-delta ' + (rgDelta >= 0 ? 'delta-up' : 'delta-down');

  document.getElementById('kpi_skus').textContent = tot.unique_skus;
  document.getElementById('kpi_skus_delta').textContent = tot.unique_colors + ' цветов · ' + tot.unique_volumes + ' фасовок';
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
    trigger: 'axis',
    axisPointer: { type: 'shadow' },
    formatter: function(p) {
      let s = p[0].name + '<br/>';
      p.forEach(function(item) {
        s += item.marker + ' ' + item.seriesName + ': ' + fmtRub(item.value) + '<br/>';
      });
      return s;
    }
  };
}

function renderTotal() {
  const records = DATA.total;
  document.getElementById('content').innerHTML =
    '<div class="charts-grid">' +
      '<div class="chart-card">' +
        '<div class="chart-title">Структура по типам продуктов</div>' +
        '<div class="chart-sub">Сравнение 2025 vs 2026, ₽</div>' +
        '<div id="chart_type" class="chart-body"></div>' +
      '</div>' +
      '<div class="chart-card">' +
        '<div class="chart-title">Топ-8 цветов</div>' +
        '<div class="chart-sub">Сравнение 2025 vs 2026, ₽</div>' +
        '<div id="chart_color" class="chart-body"></div>' +
      '</div>' +
    '</div>' +
    '<div class="charts-grid">' +
      '<div class="chart-card">' +
        '<div class="chart-title">Распределение по фасовкам</div>' +
        '<div class="chart-sub">Сравнение 2025 vs 2026, ₽</div>' +
        '<div id="chart_volume" class="chart-body"></div>' +
      '</div>' +
      '<div class="chart-card">' +
        '<div class="chart-title">Топ-10 SKU по выручке 2026</div>' +
        '<div class="chart-sub">С разбивкой по выручке 2025</div>' +
        '<div id="chart_top_sku" class="chart-body tall"></div>' +
      '</div>' +
    '</div>' +
    '<div class="charts-grid charts-full">' +
      '<div class="chart-card">' +
        '<div class="chart-title">Продажи по классам торговых точек (ТТ)</div>' +
        '<div class="chart-sub">Сравнение 2025 vs 2026, ₽</div>' +
        '<div id="chart_class" class="chart-body"></div>' +
      '</div>' +
    '</div>';

  const typeData = aggregateBy(records, 'type').filter(function(r) { return r.key !== '—'; });
  renderChart('chart_type', {
    tooltip: tooltipRub(),
    legend: { data: ['2025', '2026'], top: 0, right: 0 },
    grid: { left: 80, right: 20, top: 40, bottom: 80 },
    xAxis: { type: 'category', data: typeData.map(function(d) { return d.key; }),
             axisLabel: { rotate: 15, color: palette.ink, fontSize: 10 } },
    yAxis: { type: 'value',
             axisLabel: { color: palette.mute, fontSize: 10, formatter: function(v) { return fmt(v); } } },
    series: [
      { name: '2025', type: 'bar', data: typeData.map(function(d) { return d.v25; }), itemStyle: { color: palette.mute } },
      { name: '2026', type: 'bar', data: typeData.map(function(d) { return d.v26; }), itemStyle: { color: palette.red } },
    ],
  });

  const colorData = aggregateBy(records, 'color_display').filter(function(r) { return r.key && r.key !== '—'; }).slice(0, 8);
  renderChart('chart_color', {
    tooltip: tooltipRub(),
    legend: { data: ['2025', '2026'], top: 0, right: 0 },
    grid: { left: 80, right: 20, top: 40, bottom: 100 },
    xAxis: { type: 'category', data: colorData.map(function(d) { return d.key; }),
             axisLabel: { rotate: 30, color: palette.ink, fontSize: 9 } },
    yAxis: { type: 'value',
             axisLabel: { color: palette.mute, fontSize: 10, formatter: function(v) { return fmt(v); } } },
    series: [
      { name: '2025', type: 'bar', data: colorData.map(function(d) { return d.v25; }), itemStyle: { color: palette.mute } },
      { name: '2026', type: 'bar', data: colorData.map(function(d) { return d.v26; }), itemStyle: { color: palette.sand } },
    ],
  });

  const volData = aggregateBy(records, 'volume_display').filter(function(r) { return r.key && r.key !== '—'; }).slice(0, 8);
  renderChart('chart_volume', {
    tooltip: tooltipRub(),
    legend: { data: ['2025', '2026'], top: 0, right: 0 },
    grid: { left: 80, right: 20, top: 40, bottom: 80 },
    xAxis: { type: 'category', data: volData.map(function(d) { return d.key; }),
             axisLabel: { color: palette.ink, fontSize: 10 } },
    yAxis: { type: 'value',
             axisLabel: { color: palette.mute, fontSize: 10, formatter: function(v) { return fmt(v); } } },
    series: [
      { name: '2025', type: 'bar', data: volData.map(function(d) { return d.v25; }), itemStyle: { color: palette.mute } },
      { name: '2026', type: 'bar', data: volData.map(function(d) { return d.v26; }), itemStyle: { color: palette.blue } },
    ],
  });

  const skuData = aggregateBy(records, 'sku_name').slice(0, 10).reverse();
  renderChart('chart_top_sku', {
    tooltip: tooltipRub(),
    legend: { data: ['2025', '2026'], top: 0, right: 0 },
    grid: { left: 380, right: 50, top: 40, bottom: 30 },
    xAxis: { type: 'value',
             axisLabel: { color: palette.mute, fontSize: 10, formatter: function(v) { return fmt(v); } } },
    yAxis: { type: 'category',
             data: skuData.map(function(d) { return d.key.length > 55 ? d.key.slice(0,55)+'…' : d.key; }),
             axisLabel: { color: palette.ink, fontSize: 9 } },
    series: [
      { name: '2025', type: 'bar', data: skuData.map(function(d) { return d.v25; }), itemStyle: { color: palette.mute } },
      { name: '2026', type: 'bar', data: skuData.map(function(d) { return d.v26; }), itemStyle: { color: palette.red } },
    ],
  });

  const classData = aggregateBy(records, 'class').filter(function(r) { return r.key && r.key !== '—'; }).slice(0, 10);
  renderChart('chart_class', {
    tooltip: tooltipRub(),
    legend: { data: ['2025', '2026'], top: 0, right: 0 },
    grid: { left: 80, right: 20, top: 40, bottom: 60 },
    xAxis: { type: 'category', data: classData.map(function(d) { return d.key; }),
             axisLabel: { color: palette.ink, fontSize: 11 } },
    yAxis: { type: 'value',
             axisLabel: { color: palette.mute, fontSize: 10, formatter: function(v) { return fmt(v); } } },
    series: [
      { name: '2025', type: 'bar', data: classData.map(function(d) { return d.v25; }), itemStyle: { color: palette.mute } },
      { name: '2026', type: 'bar', data: classData.map(function(d) { return d.v26; }), itemStyle: { color: palette.amber } },
    ],
  });
}

function renderDistributor(viewKey) {
  const records = DATA[viewKey];
  const title = viewKey === 'brozeks_rygalov' ? 'БРОЗЭКС + Рыгалов' : 'Рыгалов С.В.';
  document.getElementById('content').innerHTML =
    '<div class="charts-grid">' +
      '<div class="chart-card">' +
        '<div class="chart-title">' + title + ': топ-10 SKU</div>' +
        '<div class="chart-sub">Сравнение 2025 vs 2026, ₽</div>' +
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
        '<div class="chart-sub">2025 vs 2026, ₽</div>' +
        '<div id="chart_d_type" class="chart-body"></div>' +
      '</div>' +
    '</div>';

  const skuData = aggregateBy(records, 'sku_name').slice(0, 10).reverse();
  renderChart('chart_d_sku', {
    tooltip: tooltipRub(),
    legend: { data: ['2025', '2026'], top: 0, right: 0 },
    grid: { left: 380, right: 50, top: 40, bottom: 30 },
    xAxis: { type: 'value',
             axisLabel: { color: palette.mute, fontSize: 10, formatter: function(v) { return fmt(v); } } },
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
    yAxis: { type: 'value',
             axisLabel: { color: palette.mute, fontSize: 10, formatter: function(v) { return fmt(v); } } },
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
    yAxis: { type: 'value',
             axisLabel: { color: palette.mute, fontSize: 10, formatter: function(v) { return fmt(v); } } },
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
      type: 'pie',
      radius: ['40%', '70%'],
      center: ['50%', '45%'],
      data: typeData.map(function(d, i) {
        return { name: d.key, value: d.v26, itemStyle: { color: colorPalette[i % colorPalette.length] } };
      }),
      label: { color: palette.ink, fontSize: 10, formatter: '{b}\n{d}%' },
    }],
  });
}

function renderCompare() {
  document.getElementById('content').innerHTML =
    '<div class="charts-grid">' +
      '<div class="chart-card">' +
        '<div class="chart-title">Сравнение выручки по каналам</div>' +
        '<div class="chart-sub">2025 vs 2026, ₽</div>' +
        '<div id="chart_cmp_channels" class="chart-body"></div>' +
      '</div>' +
      '<div class="chart-card">' +
        '<div class="chart-title">Доля каналов в общей выручке 2026</div>' +
        '<div class="chart-sub">БРОЗЭКС+Рыгалов vs остальная сеть</div>' +
        '<div id="chart_cmp_share" class="chart-body"></div>' +
      '</div>' +
    '</div>' +
    '<div class="charts-grid">' +
      '<div class="chart-card">' +
        '<div class="chart-title">Динамика роста по каналам</div>' +
        '<div class="chart-sub">Процент изменения 2025→2026</div>' +
        '<div id="chart_cmp_growth" class="chart-body"></div>' +
      '</div>' +
      '<div class="chart-card">' +
        '<div class="chart-title">Топ-10 SKU: БРОЗЭКС+Рыгалов vs общая сеть</div>' +
        '<div class="chart-sub">Сравнение выручки 2026</div>' +
        '<div id="chart_cmp_top" class="chart-body tall"></div>' +
      '</div>' +
    '</div>';

  const tot = DATA.aggregates.total;
  const br = DATA.aggregates.brozeks_rygalov;
  const rg = DATA.aggregates.rygalov;
  const other = {
    total_2025: tot.total_2025 - br.total_2025,
    total_2026: tot.total_2026 - br.total_2026,
  };
  other.growth_pct = ((other.total_2026 - other.total_2025) / Math.max(other.total_2025, 1)) * 100;

  renderChart('chart_cmp_channels', {
    tooltip: tooltipRub(),
    legend: { data: ['2025', '2026'], top: 0, right: 0 },
    grid: { left: 100, right: 30, top: 40, bottom: 40 },
    xAxis: { type: 'value',
             axisLabel: { color: palette.mute, fontSize: 10, formatter: function(v) { return fmt(v); } } },
    yAxis: { type: 'category',
             data: ['Рыгалов', 'БРОЗЭКС+Рыгалов', 'Остальная сеть', 'Общий итог'],
             axisLabel: { color: palette.ink, fontSize: 11 } },
    series: [
      { name: '2025', type: 'bar', data: [rg.total_2025, br.total_2025, other.total_2025, tot.total_2025], itemStyle: { color: palette.mute } },
      { name: '2026', type: 'bar', data: [rg.total_2026, br.total_2026, other.total_2026, tot.total_2026], itemStyle: { color: palette.red } },
    ],
  });

  renderChart('chart_cmp_share', {
    tooltip: { trigger: 'item', formatter: function(p) { return p.name + '<br/>' + fmtRub(p.value) + ' (' + p.percent + '%)'; } },
    legend: { bottom: 0, left: 'center' },
    series: [{
      type: 'pie',
      radius: ['45%', '75%'],
      center: ['50%', '45%'],
      data: [
        { name: 'БРОЗЭКС+Рыгалов', value: br.total_2026, itemStyle: { color: palette.red } },
        { name: 'Рыгалов С.В.', value: rg.total_2026, itemStyle: { color: palette.blue } },
        { name: 'Остальная сеть', value: other.total_2026, itemStyle: { color: palette.mute } },
      ],
      label: { color: palette.ink, fontSize: 11, formatter: '{b}\n{d}%' },
      emphasis: { itemStyle: { shadowBlur: 10, shadowColor: 'rgba(0,0,0,0.2)' } },
    }],
  });

  renderChart('chart_cmp_growth', {
    tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' },
               formatter: function(p) { return p[0].name + ': ' + fmtPct(p[0].value); } },
    grid: { left: 100, right: 80, top: 20, bottom: 40 },
    xAxis: { type: 'value',
             axisLabel: { color: palette.mute, fontSize: 10, formatter: function(v) { return v + '%'; } } },
    yAxis: { type: 'category',
             data: ['Рыгалов', 'БРОЗЭКС+Рыгалов', 'Остальная сеть', 'Общий итог'],
             axisLabel: { color: palette.ink, fontSize: 11 } },
    series: [{
      type: 'bar',
      data: [
        { value: rg.growth_pct, itemStyle: { color: rg.growth_pct >= 0 ? palette.green : palette.red } },
        { value: br.growth_pct, itemStyle: { color: br.growth_pct >= 0 ? palette.green : palette.red } },
        { value: other.growth_pct, itemStyle: { color: other.growth_pct >= 0 ? palette.green : palette.red } },
        { value: tot.growth_pct, itemStyle: { color: tot.growth_pct >= 0 ? palette.green : palette.red } },
      ],
      label: { show: true, position: 'right',
               formatter: function(p) { return fmtPct(p.value); },
               color: palette.ink, fontSize: 11, fontWeight: 600 },
    }],
  });

  const brSkuData = aggregateBy(DATA.brozeks_rygalov, 'sku_name').slice(0, 10);
  const totSkuData = aggregateBy(DATA.total, 'sku_name');
  const totMap = {};
  totSkuData.forEach(function(d) { totMap[d.key] = d.v26; });
  const cmpSku = brSkuData.map(function(d) {
    return { name: d.key, br: d.v26, tot: totMap[d.key] || 0 };
  }).sort(function(a, b) { return b.br - a.br; }).reverse();

  renderChart('chart_cmp_top', {
    tooltip: tooltipRub(),
    legend: { data: ['БРОЗЭКС+Рыгалов 2026', 'Общая сеть 2026'], top: 0, right: 0 },
    grid: { left: 380, right: 50, top: 40, bottom: 30 },
    xAxis: { type: 'value',
             axisLabel: { color: palette.mute, fontSize: 10, formatter: function(v) { return fmt(v); } } },
    yAxis: { type: 'category',
             data: cmpSku.map(function(d) { return d.name.length > 55 ? d.name.slice(0,55)+'…' : d.name; }),
             axisLabel: { color: palette.ink, fontSize: 9 } },
    series: [
      { name: 'БРОЗЭКС+Рыгалов 2026', type: 'bar', data: cmpSku.map(function(d) { return d.br; }), itemStyle: { color: palette.red } },
      { name: 'Общая сеть 2026', type: 'bar', data: cmpSku.map(function(d) { return d.tot; }), itemStyle: { color: palette.sand } },
    ],
  });
}

function renderInsights(view) {
  const tot = DATA.aggregates.total;
  const br = DATA.aggregates.brozeks_rygalov;
  const rg = DATA.aggregates.rygalov;
  const insights = [];

  if (tot.growth_pct < 0) {
    insights.push('<strong>📉 Общая динамика:</strong> выручка TITAN снизилась на ' + Math.abs(tot.growth_pct).toFixed(1) + '% в 2026 г. (с ' + fmtRub(tot.total_2025) + ' до ' + fmtRub(tot.total_2026) + ')');
  } else {
    insights.push('<strong>📈 Общая динамика:</strong> выручка TITAN выросла на ' + tot.growth_pct.toFixed(1) + '% в 2026 г.');
  }

  const brShare = (br.total_2026 / tot.total_2026 * 100).toFixed(1);
  insights.push('<strong>🏪 БРОЗЭКС+Рыгалов:</strong> ' + fmtRub(br.total_2026) + ' (' + fmtPct(br.growth_pct) + ' к 2025) — доля в общей сети: ' + brShare + '%');

  const rgShare = (rg.total_2026 / tot.total_2026 * 100).toFixed(1);
  insights.push('<strong>👤 Рыгалов С.В.:</strong> ' + fmtRub(rg.total_2026) + ' (' + fmtPct(rg.growth_pct) + ' к 2025) — доля: ' + rgShare + '%');

  if (view === 'total' || view === 'compare') {
    const records = DATA.total;
    const skuAgg = aggregateBy(records, 'sku_name').filter(function(d) { return d.v25 > 100000; });
    const growers = skuAgg.filter(function(d) { return d.delta > 50; }).sort(function(a, b) { return b.delta - a.delta; }).slice(0, 3);
    const decliners = skuAgg.filter(function(d) { return d.delta < -30; }).sort(function(a, b) { return a.delta - b.delta; }).slice(0, 3);

    if (growers.length > 0) {
      insights.push('<strong>🚀 Топ-рост 2026:</strong> ' + growers.map(function(d) { return d.key.slice(0, 40) + '… (+' + d.delta.toFixed(0) + '%)'; }).join(' · '));
    }
    if (decliners.length > 0) {
      insights.push('<strong>⚠️ Падение 2026:</strong> ' + decliners.map(function(d) { return d.key.slice(0, 40) + '… (' + d.delta.toFixed(0) + '%)'; }).join(' · '));
    }
  }

  const typeAgg = aggregateBy(DATA.total, 'type').filter(function(d) { return d.key !== '—'; });
  if (typeAgg.length > 0) {
    const bestType = typeAgg[0];
    const typeShare = (bestType.v26 / tot.total_2026 * 100).toFixed(0);
    insights.push('<strong>🎨 Лидер по типам:</strong> ' + bestType.key + ' — ' + fmtRub(bestType.v26) + ' (' + typeShare + '% от общей выручки)');
  }

  const colorAgg = aggregateBy(DATA.total, 'color_display').filter(function(d) { return d.key && d.key !== '—'; });
  if (colorAgg.length > 0) {
    const top3colors = colorAgg.slice(0, 3);
    let top3sum = 0;
    top3colors.forEach(function(d) { top3sum += d.v26; });
    const top3share = (top3sum / tot.total_2026 * 100).toFixed(0);
    insights.push('<strong>🎨 Топ-3 цвета:</strong> ' + top3colors.map(function(d) { return d.key; }).join(' · ') + ' — формируют ' + top3share + '% выручки');
  }

  const volAgg = aggregateBy(DATA.total, 'volume_display').filter(function(d) { return d.key && d.key !== '—'; });
  if (volAgg.length > 0) {
    const topVol = volAgg[0];
    const volShare = (topVol.v26 / tot.total_2026 * 100).toFixed(0);
    insights.push('<strong>📦 Лидер по фасовке:</strong> ' + topVol.key + ' — ' + fmtRub(topVol.v26) + ' (' + volShare + '%)');
  }

  if (rg.growth_pct > 0 && br.growth_pct < 0) {
    insights.push('<strong>💡 Вывод:</strong> Рыгалов — растущий канал (+' + rg.growth_pct.toFixed(1) + '%), БРОЗЭКС — падающий (' + br.growth_pct.toFixed(1) + '%). Необходим аудит причин падения в БРОЗЭКС.');
  } else if (rg.growth_pct > 0 && br.growth_pct > 0) {
    insights.push('<strong>💡 Вывод:</strong> Оба пилотных канала растут. Рекомендуется масштабирование опыта на другие точки.');
  } else if (rg.growth_pct < 0 && br.growth_pct < 0) {
    insights.push('<strong>⚠️ Вывод:</strong> Оба пилотных канала падают. Необходим срочный аудит и корректирующие мероприятия.');
  }

  document.getElementById('insights').innerHTML = insights.map(function(i) { return '<li>' + i + '</li>'; }).join('');
}

document.querySelectorAll('.tab').forEach(function(tab) {
  tab.addEventListener('click', function() {
    document.querySelectorAll('.tab').forEach(function(t) { t.classList.remove('active'); });
    tab.classList.add('active');
    const view = tab.dataset.view;
    disposeCharts();
    if (view === 'total') {
      renderTotal();
    } else if (view === 'compare') {
      renderCompare();
    } else {
      renderDistributor(view);
    }
    renderInsights(view);
  });
});

initKPIs();
renderTotal();
renderInsights('total');

window.addEventListener('resize', function() {
  charts.forEach(function(c) { c.resize(); });
});
</script>

</body>
</html>
'''

# Replace placeholder with data — escape </script> if present
data_safe = data_json.replace('</', '<\\/').replace('<script', '<\\script')
HTML = HTML.replace('__DATA_PLACEHOLDER__', data_safe)

with open(OUT_HTML, 'w', encoding='utf-8') as f:
    f.write(HTML)

print(f"HTML dashboard: {OUT_HTML}")
print(f"Size: {os.path.getsize(OUT_HTML)} bytes")
print(f"Records — total: {len(records_total)}, brozeks+rygalov: {len(records_brozeks_rygalov)}, rygalov: {len(records_rygalov)}")
