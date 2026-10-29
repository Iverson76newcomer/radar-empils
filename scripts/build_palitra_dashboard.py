"""
Parse Палитра sales data with proper category extraction and build interactive HTML dashboard.
"""
import openpyxl
import json
import re
from collections import defaultdict
import os

SRC = "/home/z/my-project/upload/Палитра Продажи Орел Титан сентябрь 2026.xlsx"
OUT_HTML = '/home/z/my-project/download/Палитра_TITAN_дашборд.html'

# ── Parse Excel with proper category extraction ──────────────────────
wb = openpyxl.load_workbook(SRC, data_only=True)
ws = wb['Лист_1']
rows = list(ws.iter_rows(values_only=True))

all_skus = []
current_region = ''
current_manager = ''
current_city = ''
current_client = ''
current_category = ''
current_geo = '—'

for i, row in enumerate(rows[14:], start=15):
    text = row[0]
    qty = row[6]
    rev = row[7]
    if not text:
        continue
    text_str = str(text).strip()
    if text_str == 'Итого':
        continue

    # Region (ALL CAPS, no comma)
    if text_str.isupper() and not text_str.endswith(',') and 'Супервайзер' not in text_str:
        current_region = text_str
        continue
    # Manager
    if 'Супервайзер' in text_str:
        current_manager = text_str
        continue
    # Brand line — skip
    if text_str.startswith('ЭМПИЛС ОРЕОЛ ТИТАН'):
        continue
    # Category (ends with "," — comma without space, e.g. "ГРУНТ-ЭМАЛЬ матовая,")
    if text_str.endswith(',') and not text_str.startswith('ЭМПИЛС') and not re.search(r',\s*\d{5,6}\s*$', text_str):
        current_category = text_str.rstrip(',').strip()
        continue
    # SKU (has article code) — CHECK BEFORE CITY!
    if re.search(r',\s*\d{5,6}\s*$', text_str):
        sku_match = re.search(r',\s*(\d{5,6})\s*$', text_str)
        sku_code = sku_match.group(1) if sku_match else ''
        name_lower = text_str.lower()
        color = ''
        colors = {
            'белая': 'Белый', 'белый': 'Белый',
            'черная': 'Чёрный', 'чёрная': 'Чёрный', 'черный': 'Чёрный',
            'серая': 'Серый', 'серый': 'Серый',
            'графитовый серый': 'Графитовый серый',
            'шоколадно-коричневая': 'Шоколадно-коричневый',
            'шоколадно-коричневый': 'Шоколадно-коричневый',
            'красно-коричневый': 'Красно-коричневый',
            'золотистая': 'Золотистый',
            'коричневая': 'Коричневый',
            'медная': 'Медный',
            'серебристая': 'Серебристый',
            'темно-зеленая': 'Тёмно-зелёный',
            'тёмно-зелёная': 'Тёмно-зелёный',
        }
        for kw, c in colors.items():
            if kw in name_lower:
                color = c
                break
        ral_match = re.search(r'RAL\s*(\d{4})', text_str, re.IGNORECASE)
        ral = f'RAL {ral_match.group(1)}' if ral_match else ''
        vol_match = re.search(r'(\d+(?:,\d+)?)\s*(кг|л)', text_str, re.IGNORECASE)
        volume = vol_match.group(0) if vol_match else ''

        # Determine category from name
        cat = ''
        if 'молотковым эффектом' in name_lower:
            cat = 'Эмали молотковые'
        elif 'быстросохнущая' in name_lower:
            cat = 'Грунт-эмаль быстросохнущая'
        elif 'по ржавчине' in name_lower:
            cat = 'Грунт-эмаль по ржавчине'
        else:
            cat = current_category or 'Прочее'

        all_skus.append({
            'name': text_str,
            'code': sku_code,
            'qty': int(qty) if qty else 0,
            'rev': float(rev) if rev else 0,
            'category': cat,
            'region': current_region,
            'manager': current_manager,
            'city': current_city,
            'client': current_client,
            'color': color,
            'ral': ral,
            'volume': volume,
        })
        continue
    # Address
    if re.search(r'\d{6}|ул\.|д\.|р-н|край|область', text_str):
        if 'Краснодарский край' in text_str:
            current_geo = 'Краснодарский край'
        elif 'Ростовская' in text_str:
            current_geo = 'Ростовская область'
        elif 'Ставропольский' in text_str:
            current_geo = 'Ставропольский край'
        elif 'Адыгея' in text_str:
            current_geo = 'Республика Адыгея'
        continue
    # Client (ИП / ООО / etc) — CHECK BEFORE CITY!
    if re.search(r'ИП|ООО|ЗАО', text_str):
        current_client = text_str
        continue
    # City (mixed case, no comma at end, not ALL CAPS)
    if not text_str.endswith(',') and text_str[0].isupper() and not text_str.isupper():
        current_city = text_str
        continue

# Aggregate
def agg_by(records, key):
    m = defaultdict(lambda: {'qty': 0, 'rev': 0, 'count': 0})
    for r in records:
        k = r.get(key, '—') or '—'
        m[k]['qty'] += r['qty']
        m[k]['rev'] += r['rev']
        m[k]['count'] += 1
    return [{'key': k, 'qty': v['qty'], 'rev': v['rev'], 'count': v['count']} for k, v in m.items()]

total_qty = sum(s['qty'] for s in all_skus)
total_rev = sum(s['rev'] for s in all_skus)

# Aggregate clients separately (track unique client names from SKU records)
clients_agg = defaultdict(lambda: {'qty': 0, 'rev': 0, 'city': '', 'geo': '', 'region': '', 'manager': ''})
for s in all_skus:
    if s['client']:
        c = clients_agg[s['client']]
        c['qty'] += s['qty']
        c['rev'] += s['rev']
        c['city'] = s['city']
        c['region'] = s['region']
        c['manager'] = s['manager']

clients_list = []
for name, d in sorted(clients_agg.items(), key=lambda x: -x[1]['rev']):
    clients_list.append({
        'name': name,
        'city': d['city'],
        'region': d['region'],
        'manager': d['manager'],
        'qty': d['qty'],
        'rev': d['rev'],
    })

data = {
    'total_qty': total_qty,
    'total_rev': total_rev,
    'categories': agg_by(all_skus, 'category'),
    'colors': [d for d in agg_by(all_skus, 'color') if d['key'] != '—'],
    'volumes': [d for d in agg_by(all_skus, 'volume') if d['key']],
    'regions': agg_by(all_skus, 'region'),
    'managers': [d for d in agg_by(all_skus, 'manager') if d['key']],
    'skus': sorted(all_skus, key=lambda x: -x['rev']),
    'clients': clients_list,
    'unique_skus': len(set(s['name'] for s in all_skus)),
    'unique_clients': len(clients_agg),
}

print(f"Total: {data['total_qty']} шт, {data['total_rev']:,.0f} ₽")
print(f"Unique SKUs: {data['unique_skus']}, Clients: {data['unique_clients']}")
print(f"Categories: {[c['key'] for c in data['categories']]}")
print(f"Top 5 clients:")
for c in data['clients'][:5]:
    print(f"  {c['name'][:40]:40s} | rev={c['rev']:>10,.0f}")

data_json = json.dumps(data, ensure_ascii=False)

# ── Build HTML ────────────────────────────────────────────────────────
HTML = r'''<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<title>Палитра · TITAN · Сентябрь 2026</title>
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
    color: var(--mute);
    margin-top: 4px;
  }
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
      <div class="header-title">Палитра · ОРЕОЛ TITAN · Сентябрь 2026</div>
      <div class="header-sub">Анализ продаж у нового дистрибьютора · АКБ · Товарооборот · Структура спроса</div>
    </div>
    <div class="pill">ПЕРВЫЙ МЕСЯЦ · СТАРТ</div>
  </div>
</div>

<div class="kpi-grid">
  <div class="kpi-card">
    <div class="kpi-label">Выручка за месяц</div>
    <div class="kpi-value" id="kpi_rev">—</div>
    <div class="kpi-delta">с НДС, сентябрь 2026</div>
  </div>
  <div class="kpi-card accent">
    <div class="kpi-label">Продажи (шт)</div>
    <div class="kpi-value" id="kpi_qty">—</div>
    <div class="kpi-delta">единиц хранения</div>
  </div>
  <div class="kpi-card blue">
    <div class="kpi-label">АКБ (клиентов)</div>
    <div class="kpi-value" id="kpi_clients">—</div>
    <div class="kpi-delta">активная клиентская база</div>
  </div>
  <div class="kpi-card green">
    <div class="kpi-label">Уникальных SKU</div>
    <div class="kpi-value" id="kpi_skus">—</div>
    <div class="kpi-delta">из ассортимента TITAN</div>
  </div>
</div>

<div class="section-heading">Структура продаж</div>
<div class="charts-grid">
  <div class="chart-card">
    <div class="chart-title">По категориям продуктов</div>
    <div class="chart-sub">Выручка по типам ЛКМ, ₽</div>
    <div id="chart_cat" class="chart-body"></div>
  </div>
  <div class="chart-card">
    <div class="chart-title">По цветам (топ-8)</div>
    <div class="chart-sub">Выручка, ₽</div>
    <div id="chart_color" class="chart-body"></div>
  </div>
</div>

<div class="charts-grid">
  <div class="chart-card">
    <div class="chart-title">По фасовкам</div>
    <div class="chart-sub">Выручка, ₽</div>
    <div id="chart_vol" class="chart-body"></div>
  </div>
  <div class="chart-card">
    <div class="chart-title">По менеджерам</div>
    <div class="chart-sub">Выручка и количество, ₽ / шт</div>
    <div id="chart_mgr" class="chart-body"></div>
  </div>
</div>

<div class="section-heading">Регионы и клиенты</div>
<div class="charts-grid">
  <div class="chart-card">
    <div class="chart-title">По бизнес-регионам Палитры</div>
    <div class="chart-sub">Выручка по регионам дистрибьютора, ₽</div>
    <div id="chart_reg" class="chart-body tall"></div>
  </div>
  <div class="chart-card">
    <div class="chart-title">Топ-15 клиентов (АКБ)</div>
    <div class="chart-sub">Выручка по клиентам, ₽</div>
    <div id="chart_clients" class="chart-body tall"></div>
  </div>
</div>

<div class="charts-grid">
  <div class="chart-card">
    <div class="chart-title">Топ-15 SKU по выручке</div>
    <div class="chart-sub">Лучшие позиции, ₽</div>
    <div id="chart_sku" class="chart-body tall"></div>
  </div>
  <div class="chart-card">
    <div class="chart-title">Цвет × Фасовка</div>
    <div class="chart-sub">Тепловая карта продаж, ₽</div>
    <div id="chart_heatmap" class="chart-body tall"></div>
  </div>
</div>

<div class="summary">
  <h3>Ключевые выводы</h3>
  <ul id="insights">—</ul>
</div>

<div class="footer">
  Источник: Палитра Продажи Орел Титан сентябрь 2026.xlsx · Период: 01.09.2026 — 30.09.2026
</div>

<script>
const DATA = __DATA_PLACEHOLDER__;

const palette = {
  red: '#C00000', sand: '#C9A77C', green: '#8BAA8A', blue: '#6B8FA8',
  amber: '#D4A574', mute: '#8A95A3', ink: '#2F3A48',
};

const colorPalette = [
  '#C00000', '#C9A77C', '#6B8FA8', '#8BAA8A', '#D4A574',
  '#B5614A', '#5A6776', '#A7C0D0', '#8B7355', '#8A95A3',
  '#D9D2C2', '#C7333E',
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

// Init KPIs
(function() {
  document.getElementById('kpi_rev').textContent = fmtRub(DATA.total_rev);
  document.getElementById('kpi_qty').textContent = DATA.total_qty.toLocaleString('ru-RU');
  document.getElementById('kpi_clients').textContent = DATA.unique_clients;
  document.getElementById('kpi_skus').textContent = DATA.unique_skus;
})();

function tooltipRub() {
  return {
    trigger: 'axis', axisPointer: { type: 'shadow' },
    formatter: function(p) {
      let s = p[0].name + '<br/>';
      p.forEach(function(item) {
        s += item.marker + ' ' + item.seriesName + ': ' + fmtRub(item.value) + '<br/>';
      });
      return s;
    }
  };
}

// Categories chart (pie)
const catData = DATA.categories.sort(function(a,b) { return b.rev - a.rev; });
echarts.init(document.getElementById('chart_cat')).setOption({
  tooltip: { trigger: 'item', formatter: function(p) { return p.name + '<br/>' + fmtRub(p.value) + ' (' + p.percent + '%)'; } },
  legend: { bottom: 0, left: 'center', textStyle: { fontSize: 10 } },
  series: [{
    type: 'pie',
    radius: ['40%', '70%'],
    center: ['50%', '45%'],
    data: catData.map(function(d, i) {
      return { name: d.key, value: d.rev, itemStyle: { color: colorPalette[i % colorPalette.length] } };
    }),
    label: { color: palette.ink, fontSize: 9, formatter: '{b}\n{d}%' },
  }],
});

// Colors chart
const colorData = DATA.colors.sort(function(a,b) { return b.rev - a.rev; }).slice(0, 8);
echarts.init(document.getElementById('chart_color')).setOption({
  tooltip: tooltipRub(),
  grid: { left: 80, right: 30, top: 20, bottom: 80 },
  xAxis: { type: 'category', data: colorData.map(function(d) { return d.key; }),
           axisLabel: { rotate: 30, color: palette.ink, fontSize: 9 } },
  yAxis: { type: 'value', axisLabel: { color: palette.mute, fontSize: 10, formatter: function(v) { return fmt(v); } } },
  series: [{
    type: 'bar',
    data: colorData.map(function(d, i) { return { value: d.rev, itemStyle: { color: colorPalette[i % colorPalette.length] } }; }),
    label: { show: true, position: 'top', formatter: function(p) { return fmt(p.value); }, color: palette.ink, fontSize: 9, fontWeight: 600 },
  }],
});

// Volumes chart
const volData = DATA.volumes.sort(function(a,b) { return b.rev - a.rev; });
echarts.init(document.getElementById('chart_vol')).setOption({
  tooltip: tooltipRub(),
  grid: { left: 80, right: 30, top: 20, bottom: 60 },
  xAxis: { type: 'category', data: volData.map(function(d) { return d.key; }),
           axisLabel: { color: palette.ink, fontSize: 10 } },
  yAxis: { type: 'value', axisLabel: { color: palette.mute, fontSize: 10, formatter: function(v) { return fmt(v); } } },
  series: [{
    type: 'bar',
    data: volData.map(function(d) { return d.rev; }),
    itemStyle: { color: palette.blue },
    label: { show: true, position: 'top', formatter: function(p) { return fmt(p.value); }, color: palette.ink, fontSize: 9, fontWeight: 600 },
  }],
});

// Managers chart (dual axis: rev + qty)
const mgrData = DATA.managers.sort(function(a,b) { return b.rev - a.rev; });
echarts.init(document.getElementById('chart_mgr')).setOption({
  tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' },
    formatter: function(p) {
      let s = p[0].name + '<br/>';
      p.forEach(function(item) {
        if (item.seriesName === 'Выручка') s += item.marker + ' ' + item.seriesName + ': ' + fmtRub(item.value) + '<br/>';
        else s += item.marker + ' ' + item.seriesName + ': ' + item.value + ' шт<br/>';
      });
      return s;
    }
  },
  legend: { data: ['Выручка', 'Количество'], top: 0, right: 0 },
  grid: { left: 80, right: 60, top: 40, bottom: 60 },
  xAxis: { type: 'category', data: mgrData.map(function(d) { return d.key.replace('Супервайзер ', ''); }),
           axisLabel: { color: palette.ink, fontSize: 10 } },
  yAxis: [
    { type: 'value', name: '₽', axisLabel: { color: palette.mute, fontSize: 10, formatter: function(v) { return fmt(v); } } },
    { type: 'value', name: 'шт', axisLabel: { color: palette.mute, fontSize: 10 } },
  ],
  series: [
    { name: 'Выручка', type: 'bar', data: mgrData.map(function(d) { return d.rev; }), itemStyle: { color: palette.red },
      label: { show: true, position: 'top', formatter: function(p) { return fmt(p.value); }, color: palette.ink, fontSize: 9, fontWeight: 600 } },
    { name: 'Количество', type: 'bar', yAxisIndex: 1, data: mgrData.map(function(d) { return d.qty; }), itemStyle: { color: palette.sand },
      label: { show: true, position: 'top', color: palette.ink, fontSize: 9, fontWeight: 600 } },
  ],
});

// Regions chart (horizontal bar)
const regData = DATA.regions.sort(function(a,b) { return a.rev - b.rev; });
echarts.init(document.getElementById('chart_reg')).setOption({
  tooltip: tooltipRub(),
  grid: { left: 200, right: 80, top: 20, bottom: 30 },
  xAxis: { type: 'value', axisLabel: { color: palette.mute, fontSize: 10, formatter: function(v) { return fmt(v); } } },
  yAxis: { type: 'category',
           data: regData.map(function(d) { return d.key.length > 30 ? d.key.slice(0,30)+'…' : d.key; }),
           axisLabel: { color: palette.ink, fontSize: 9 } },
  series: [{
    type: 'bar',
    data: regData.map(function(d, i) { return { value: d.rev, itemStyle: { color: colorPalette[i % colorPalette.length] } }; }),
    label: { show: true, position: 'right', formatter: function(p) { return fmtRub(p.value) + ' (' + (p.value/DATA.total_rev*100).toFixed(1) + '%)'; }, color: palette.ink, fontSize: 9, fontWeight: 600 },
  }],
});

// Top 15 SKUs
const skuData = DATA.skus.slice(0, 15).reverse();
echarts.init(document.getElementById('chart_sku')).setOption({
  tooltip: tooltipRub(),
  grid: { left: 380, right: 50, top: 20, bottom: 30 },
  xAxis: { type: 'value', axisLabel: { color: palette.mute, fontSize: 10, formatter: function(v) { return fmt(v); } } },
  yAxis: { type: 'category',
           data: skuData.map(function(d) { return d.name.length > 55 ? d.name.slice(0,55)+'…' : d.name; }),
           axisLabel: { color: palette.ink, fontSize: 8 } },
  series: [{
    type: 'bar',
    data: skuData.map(function(d) { return d.rev; }),
    itemStyle: { color: palette.red },
    label: { show: true, position: 'right', formatter: function(p) { return fmt(p.value); }, color: palette.ink, fontSize: 9, fontWeight: 600 },
  }],
});

// Top 15 clients
const clientData = (DATA.clients || []).slice(0, 15).reverse();
echarts.init(document.getElementById('chart_clients')).setOption({
  tooltip: tooltipRub(),
  grid: { left: 280, right: 80, top: 20, bottom: 30 },
  xAxis: { type: 'value', axisLabel: { color: palette.mute, fontSize: 10, formatter: function(v) { return fmt(v); } } },
  yAxis: { type: 'category',
           data: clientData.map(function(d) { return d.name.length > 40 ? d.name.slice(0,40)+'…' : d.name; }),
           axisLabel: { color: palette.ink, fontSize: 9 } },
  series: [{
    type: 'bar',
    data: clientData.map(function(d, i) { return { value: d.rev, itemStyle: { color: colorPalette[i % colorPalette.length] } }; }),
    label: { show: true, position: 'right', formatter: function(p) { return fmtRub(p.value); }, color: palette.ink, fontSize: 9, fontWeight: 600 },
  }],
});

// Heatmap: Color × Volume
const heatColors = DATA.colors.map(function(d) { return d.key; }).filter(function(k) { return k && k !== '—'; });
const heatVols = DATA.volumes.map(function(d) { return d.key; }).filter(function(k) { return k; });
const heatData = [];
let heatMax = 0;
heatColors.forEach(function(c, i) {
  heatVols.forEach(function(v, j) {
    let sum = 0;
    DATA.skus.forEach(function(s) {
      if (s.color === c && s.volume === v) sum += s.rev;
    });
    heatData.push([j, i, sum]);
    if (sum > heatMax) heatMax = sum;
  });
});
echarts.init(document.getElementById('chart_heatmap')).setOption({
  tooltip: { position: 'top', formatter: function(p) {
    return heatColors[p.value[1]] + ' / ' + heatVols[p.value[0]] + '<br/>' + fmtRub(p.value[2]);
  }},
  grid: { left: 120, right: 30, top: 20, bottom: 80 },
  xAxis: { type: 'category', data: heatVols, axisLabel: { color: palette.ink, fontSize: 10 } },
  yAxis: { type: 'category', data: heatColors, axisLabel: { color: palette.ink, fontSize: 9 } },
  visualMap: { min: 0, max: heatMax, calculable: true, orient: 'horizontal', left: 'center', bottom: 10,
    textStyle: { color: palette.ink, fontSize: 9 },
    inRange: { color: ['#F5F2EC', '#C9A77C', '#C00000'] } },
  series: [{
    type: 'heatmap', data: heatData,
    label: { show: true, formatter: function(p) { return p.value[2] > 0 ? fmt(p.value[2]) : ''; }, color: palette.ink, fontSize: 8 },
    emphasis: { itemStyle: { shadowBlur: 10, shadowColor: 'rgba(0, 0, 0, 0.5)' } },
  }],
});

// Insights
function renderInsights() {
  const insights = [];
  insights.push('<strong>Стартовый месяц:</strong> ' + fmtRub(DATA.total_rev) + ' за сентябрь 2026 — первый месяц продаж TITAN у дистрибьютора «Палитра»');
  insights.push('<strong>АКБ:</strong> ' + DATA.unique_clients + ' клиентов купили TITAN, средний чек на клиента — ' + fmtRub(DATA.total_rev / DATA.unique_clients));
  insights.push('<strong>Ассортимент:</strong> ' + DATA.unique_skus + ' SKU продано за месяц — MML TITAN полностью представлен');

  // Top region
  const topReg = DATA.regions.sort(function(a,b) { return b.rev - a.rev; })[0];
  if (topReg) {
    insights.push('<strong>Бизнес-регион-лидер:</strong> ' + topReg.key + ' — ' + fmtRub(topReg.rev) + ' (' + (topReg.rev/DATA.total_rev*100).toFixed(1) + '%)');
  }

  // Top color
  const topColor = DATA.colors.sort(function(a,b) { return b.rev - a.rev; })[0];
  if (topColor) {
    insights.push('<strong>Лидер по цвету:</strong> ' + topColor.key + ' — ' + fmtRub(topColor.rev) + ' (' + (topColor.rev/DATA.total_rev*100).toFixed(1) + '%)');
  }

  // Top volume
  const topVol = DATA.volumes.sort(function(a,b) { return b.rev - a.rev; })[0];
  if (topVol) {
    insights.push('<strong>Лидер по фасовке:</strong> ' + topVol.key + ' — ' + fmtRub(topVol.rev) + ' (' + (topVol.rev/DATA.total_rev*100).toFixed(1) + '%)');
  }

  // Manager comparison
  const mgrs = DATA.managers.sort(function(a,b) { return b.rev - a.rev; });
  if (mgrs.length >= 2) {
    insights.push('<strong>Менеджеры:</strong> ' + mgrs[0].key.replace('Супервайзер ', '') + ' — ' + fmtRub(mgrs[0].rev) + ' (' + (mgrs[0].rev/DATA.total_rev*100).toFixed(1) + '%), ' + mgrs[1].key.replace('Супервайзер ', '') + ' — ' + fmtRub(mgrs[1].rev) + ' (' + (mgrs[1].rev/DATA.total_rev*100).toFixed(1) + '%)');
  }

  // Top SKU
  const topSku = DATA.skus[0];
  if (topSku) {
    insights.push('<strong>Топ-SKU:</strong> ' + topSku.name.slice(0, 60) + '… — ' + fmtRub(topSku.rev));
  }

  // Concentration
  const top5Rev = DATA.skus.slice(0, 5).reduce(function(s, d) { return s + d.rev; }, 0);
  insights.push('<strong>Концентрация SKU:</strong> топ-5 SKU дают ' + (top5Rev/DATA.total_rev*100).toFixed(1) + '% выручки, топ-15 — ' + (DATA.skus.slice(0,15).reduce(function(s,d){return s+d.rev;},0)/DATA.total_rev*100).toFixed(1) + '%');

  document.getElementById('insights').innerHTML = insights.map(function(i) { return '<li>' + i + '</li>'; }).join('');
}
renderInsights();

window.addEventListener('resize', function() {
  ['chart_cat', 'chart_color', 'chart_vol', 'chart_mgr', 'chart_reg', 'chart_sku', 'chart_clients', 'chart_heatmap'].forEach(function(id) {
    const inst = echarts.getInstanceByDom(document.getElementById(id));
    if (inst) inst.resize();
  });
});
</script>

</body>
</html>
'''

data_safe = data_json.replace('</', '<\\/').replace('<script', '<\\script')
HTML = HTML.replace('__DATA_PLACEHOLDER__', data_safe)

with open(OUT_HTML, 'w', encoding='utf-8') as f:
    f.write(HTML)

print(f"\nHTML dashboard: {OUT_HTML} ({os.path.getsize(OUT_HTML)} bytes)")
