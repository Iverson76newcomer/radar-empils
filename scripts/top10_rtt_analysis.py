"""
Analyze top-10 retail outlets (РТТ) by:
1. Number of unique TITAN SKUs sold
2. Total turnover from TITAN SKUs

Source: 'Титан Брозэкс, Рыгалов.xlsx', sheet 'выгрузка'
- Column A: Дистрибьютор (БРОЗЭКС / Рыгалов)
- Column B: СерияТорговойМарки (ОРЕОЛ TITAN — only need these rows)
- Column F: Фактическая сумма отгрузки, с НДС, руб
- Column I: Название клиента (РТТ)

Output:
- /home/z/my-project/download/TITAN_топ_10_РТТ.xlsx — formatted Excel with two sheets
- /home/z/my-project/download/TITAN_топ_10_РТТ.md — analysis report
"""
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict
import os

SRC = "/home/z/my-project/upload/Титан Брозэкс, Рыгалов.xlsx"
OUT_XLSX = "/home/z/my-project/download/TITAN_топ_10_РТТ.xlsx"
OUT_MD = "/home/z/my-project/download/TITAN_топ_10_РТТ.md"

# ── Read raw data ───────────────────────────────────────────────────
wb_src = openpyxl.load_workbook(SRC, data_only=True)
ws = wb_src['выгрузка']

# Aggregate per РТТ: unique SKU count + total turnover (separately by distributor)
distributors = set()
rtt_data = defaultdict(lambda: {
    'distributor': '',
    'sku_set': set(),
    'turnover_2025': 0.0,
    'turnover_2026': 0.0,
    'turnover_total': 0.0,
    'weight_total': 0.0,
    'class_tt': set(),
})

# Skip header (row 1)
for row in ws.iter_rows(min_row=2, values_only=True):
    if not row or len(row) < 10:
        continue
    distributor = row[0]
    series = row[1]
    sku_name = row[2]
    sku_code = row[3]
    weight = row[4] if row[4] else 0
    amount = row[5] if row[5] else 0
    year = row[6]
    month = row[7]
    rtt_name = row[8]
    class_tt = row[9] or ''

    # Only TITAN series
    if not series or 'TITAN' not in str(series).upper():
        continue
    if not rtt_name:
        continue

    # Normalize РТТ name (some have extra whitespace)
    rtt_name = str(rtt_name).strip()
    distributor = str(distributor).strip() if distributor else ''

    key = (distributor, rtt_name)
    rec = rtt_data[key]
    rec['distributor'] = distributor
    rec['sku_set'].add((sku_code, sku_name) if sku_code else sku_name)
    rec['turnover_total'] += float(amount) if amount else 0
    rec['weight_total'] += float(weight) if weight else 0
    if year == 2025:
        rec['turnover_2025'] += float(amount) if amount else 0
    elif year == 2026:
        rec['turnover_2026'] += float(amount) if amount else 0
    if class_tt:
        rec['class_tt'].add(class_tt)

    distributors.add(distributor)

print(f"Unique distributors: {len(distributors)}")
print(f"  - {distributors}")
print(f"Unique РТТ with TITAN sales: {len(rtt_data)}")

# Convert to list with computed SKU count
rtt_list = []
for (distributor, rtt_name), d in rtt_data.items():
    rtt_list.append({
        'distributor': distributor,
        'rtt_name': rtt_name,
        'sku_count': len(d['sku_set']),
        'turnover_2025': d['turnover_2025'],
        'turnover_2026': d['turnover_2026'],
        'turnover_total': d['turnover_total'],
        'weight_total': d['weight_total'],
        'class_tt': ', '.join(sorted(d['class_tt'])) if d['class_tt'] else '',
        'avg_check': d['turnover_total'] / max(len(d['sku_set']), 1),
    })

# ── Build output Excel ───────────────────────────────────────────────
wb = Workbook()

# Style helpers
THIN = Side(border_style="thin", color="D9D2C2")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill(start_color="2F3A48", end_color="2F3A48", fill_type="solid")
HEADER_FONT = Font(name="Inter", size=11, bold=True, color="FFFFFF")
TITLE_FONT = Font(name="Inter", size=14, bold=True, color="2F3A48")
SUBTITLE_FONT = Font(name="Inter", size=10, italic=True, color="8A95A3")
BODY_FONT = Font(name="Inter", size=10, color="2F3A48")
BODY_FONT_BOLD = Font(name="Inter", size=10, bold=True, color="2F3A48")
ACCENT_FONT = Font(name="Roboto Mono", size=10, bold=True, color="C00000")
ALT_FILL = PatternFill(start_color="F5F2EC", end_color="F5F2EC", fill_type="solid")
WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

# ── Sheet 1: Top 10 by SKU count ─────────────────────────────────────
ws1 = wb.active
ws1.title = "Топ-10 по кол-ву SKU"

ws1.merge_cells('A1:I1')
ws1['A1'] = "Топ-10 РТТ по количеству SKU ОРЕОЛ TITAN"
ws1['A1'].font = TITLE_FONT
ws1['A1'].alignment = Alignment(horizontal='left', vertical='center')

ws1.merge_cells('A2:I2')
ws1['A2'] = "Дистрибьюторы: БРОЗЭКС + Рыгалов С.В. · Период: 2025-2026"
ws1['A2'].font = SUBTITLE_FONT

# Header row 4
headers1 = ['#', 'Дистрибьютор', 'РТТ', 'Класс ТТ', 'Кол-во SKU', 'Оборот 2025 (₽)', 'Оборот 2026 (₽)', 'Оборот всего (₽)', 'Ср. чек на SKU (₽)']
for col, h in enumerate(headers1, 1):
    cell = ws1.cell(row=4, column=col)
    cell.value = h
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = BORDER

# Sort by SKU count desc, then turnover desc
top_sku = sorted(rtt_list, key=lambda x: (-x['sku_count'], -x['turnover_total']))[:10]

for i, r in enumerate(top_sku):
    row_num = 5 + i
    fill = ALT_FILL if i % 2 == 0 else WHITE_FILL
    values = [
        i + 1,
        r['distributor'],
        r['rtt_name'],
        r['class_tt'],
        r['sku_count'],
        r['turnover_2025'],
        r['turnover_2026'],
        r['turnover_total'],
        r['avg_check'],
    ]
    for col, v in enumerate(values, 1):
        cell = ws1.cell(row=row_num, column=col)
        cell.value = v
        cell.border = BORDER
        cell.fill = fill
        if col == 1:
            cell.font = BODY_FONT_BOLD
            cell.alignment = Alignment(horizontal='center', vertical='center')
        elif col == 5:
            cell.font = Font(name="Roboto Mono", size=11, bold=True, color="C00000")
            cell.alignment = Alignment(horizontal='center', vertical='center')
        elif col in (6, 7, 8, 9):
            cell.font = Font(name="Roboto Mono", size=10, color="2F3A48")
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '#,##0 ₽'
        else:
            cell.font = BODY_FONT
            cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

# Column widths
widths = [4, 24, 45, 10, 12, 18, 18, 18, 18]
for col, w in enumerate(widths, 1):
    ws1.column_dimensions[get_column_letter(col)].width = w

# Row heights
ws1.row_dimensions[1].height = 22
ws1.row_dimensions[2].height = 16
ws1.row_dimensions[4].height = 32
for r in range(5, 15):
    ws1.row_dimensions[r].height = 30

# Freeze header
ws1.freeze_panes = 'A5'

# Summary box below
ws1['A16'] = "ИТОГ:"
ws1['A16'].font = Font(name="Inter", size=11, bold=True, color="2F3A48")
total_sku_sum = sum(r['sku_count'] for r in top_sku)
total_turnover_top10 = sum(r['turnover_total'] for r in top_sku)
ws1['E16'] = total_sku_sum
ws1['E16'].font = Font(name="Roboto Mono", size=11, bold=True, color="C00000")
ws1['H16'] = total_turnover_top10
ws1['H16'].font = Font(name="Roboto Mono", size=11, bold=True, color="C00000")
ws1['H16'].number_format = '#,##0 ₽'

# Grand total (all РТТ)
ws1['A18'] = "Все РТТ (для сравнения):"
ws1['A18'].font = SUBTITLE_FONT
all_sku_sum = sum(r['sku_count'] for r in rtt_list)
all_turnover = sum(r['turnover_total'] for r in rtt_list)
ws1['E18'] = all_sku_sum
ws1['E18'].font = Font(name="Roboto Mono", size=10, color="8A95A3")
ws1['H18'] = all_turnover
ws1['H18'].font = Font(name="Roboto Mono", size=10, color="8A95A3")
ws1['H18'].number_format = '#,##0 ₽'

ws1['A19'] = f"Доля топ-10 в обороте: {total_turnover_top10/all_turnover*100:.1f}%"
ws1['A19'].font = SUBTITLE_FONT


# ── Sheet 2: Top 10 by turnover ──────────────────────────────────────
ws2 = wb.create_sheet("Топ-10 по обороту")

ws2.merge_cells('A1:I1')
ws2['A1'] = "Топ-10 РТТ по обороту ОРЕОЛ TITAN"
ws2['A1'].font = TITLE_FONT

ws2.merge_cells('A2:I2')
ws2['A2'] = "Дистрибьюторы: БРОЗЭКС + Рыгалов С.В. · Период: 2025-2026"
ws2['A2'].font = SUBTITLE_FONT

for col, h in enumerate(headers1, 1):
    cell = ws2.cell(row=4, column=col)
    cell.value = h
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = BORDER

# Sort by turnover desc
top_rev = sorted(rtt_list, key=lambda x: -x['turnover_total'])[:10]

for i, r in enumerate(top_rev):
    row_num = 5 + i
    fill = ALT_FILL if i % 2 == 0 else WHITE_FILL
    values = [
        i + 1,
        r['distributor'],
        r['rtt_name'],
        r['class_tt'],
        r['sku_count'],
        r['turnover_2025'],
        r['turnover_2026'],
        r['turnover_total'],
        r['avg_check'],
    ]
    for col, v in enumerate(values, 1):
        cell = ws2.cell(row=row_num, column=col)
        cell.value = v
        cell.border = BORDER
        cell.fill = fill
        if col == 1:
            cell.font = BODY_FONT_BOLD
            cell.alignment = Alignment(horizontal='center', vertical='center')
        elif col == 5:
            cell.font = Font(name="Roboto Mono", size=11, bold=True, color="2F3A48")
            cell.alignment = Alignment(horizontal='center', vertical='center')
        elif col == 8:
            cell.font = Font(name="Roboto Mono", size=11, bold=True, color="C00000")
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '#,##0 ₽'
        elif col in (6, 7, 9):
            cell.font = Font(name="Roboto Mono", size=10, color="2F3A48")
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '#,##0 ₽'
        else:
            cell.font = BODY_FONT
            cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

for col, w in enumerate(widths, 1):
    ws2.column_dimensions[get_column_letter(col)].width = w

ws2.row_dimensions[1].height = 22
ws2.row_dimensions[2].height = 16
ws2.row_dimensions[4].height = 32
for r in range(5, 15):
    ws2.row_dimensions[r].height = 30

ws2.freeze_panes = 'A5'

# Summary
ws2['A16'] = "ИТОГ:"
ws2['A16'].font = Font(name="Inter", size=11, bold=True, color="2F3A48")
total_sku_sum2 = sum(r['sku_count'] for r in top_rev)
total_turnover_top10_2 = sum(r['turnover_total'] for r in top_rev)
ws2['E16'] = total_sku_sum2
ws2['E16'].font = Font(name="Roboto Mono", size=11, bold=True, color="C00000")
ws2['H16'] = total_turnover_top10_2
ws2['H16'].font = Font(name="Roboto Mono", size=11, bold=True, color="C00000")
ws2['H16'].number_format = '#,##0 ₽'

ws2['A18'] = "Все РТТ (для сравнения):"
ws2['A18'].font = SUBTITLE_FONT
ws2['E18'] = all_sku_sum
ws2['E18'].font = Font(name="Roboto Mono", size=10, color="8A95A3")
ws2['H18'] = all_turnover
ws2['H18'].font = Font(name="Roboto Mono", size=10, color="8A95A3")
ws2['H18'].number_format = '#,##0 ₽'

ws2['A19'] = f"Доля топ-10 в обороте: {total_turnover_top10_2/all_turnover*100:.1f}%"
ws2['A19'].font = SUBTITLE_FONT


# ── Sheet 3: All РТТ — combined ranking ─────────────────────────────
ws3 = wb.create_sheet("Полный рейтинг РТТ")
ws3.merge_cells('A1:J1')
ws3['A1'] = "Полный рейтинг всех РТТ по продажам ОРЕОЛ TITAN"
ws3['A1'].font = TITLE_FONT

ws3.merge_cells('A2:J2')
ws3['A2'] = f"Всего РТТ: {len(rtt_list)} · Период: 2025-2026"
ws3['A2'].font = SUBTITLE_FONT

headers3 = ['Ранг по SKU', 'Ранг по обороту', 'Дистрибьютор', 'РТТ', 'Класс', 'SKU', 'Оборот 2025 (₽)', 'Оборот 2026 (₽)', 'Оборот всего (₽)', 'Ср. чек/SKU (₽)']
for col, h in enumerate(headers3, 1):
    cell = ws3.cell(row=4, column=col)
    cell.value = h
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = BORDER

# Sort all by turnover desc
all_sorted = sorted(rtt_list, key=lambda x: -x['turnover_total'])
# Compute rank by SKU count
sorted_by_sku = sorted(rtt_list, key=lambda x: (-x['sku_count'], -x['turnover_total']))
rank_sku = {id(r): i + 1 for i, r in enumerate(sorted_by_sku)}
rank_rev = {id(r): i + 1 for i, r in enumerate(all_sorted)}

# Show top 30 by turnover
for i, r in enumerate(all_sorted[:30]):
    row_num = 5 + i
    fill = ALT_FILL if i % 2 == 0 else WHITE_FILL
    values = [
        rank_sku[id(r)],
        rank_rev[id(r)],
        r['distributor'],
        r['rtt_name'],
        r['class_tt'],
        r['sku_count'],
        r['turnover_2025'],
        r['turnover_2026'],
        r['turnover_total'],
        r['avg_check'],
    ]
    for col, v in enumerate(values, 1):
        cell = ws3.cell(row=row_num, column=col)
        cell.value = v
        cell.border = BORDER
        cell.fill = fill
        if col in (1, 2):
            cell.font = Font(name="Roboto Mono", size=10, bold=True, color="2F3A48")
            cell.alignment = Alignment(horizontal='center', vertical='center')
            # Highlight top-10 in both rankings
            if v <= 10:
                cell.fill = PatternFill(start_color="FFF5E6", end_color="FFF5E6", fill_type="solid")
                cell.font = Font(name="Roboto Mono", size=10, bold=True, color="C00000")
        elif col == 6:
            cell.font = Font(name="Roboto Mono", size=10, color="2F3A48")
            cell.alignment = Alignment(horizontal='center', vertical='center')
        elif col in (7, 8, 9, 10):
            cell.font = Font(name="Roboto Mono", size=10, color="2F3A48")
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '#,##0 ₽'
        else:
            cell.font = BODY_FONT
            cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

widths3 = [10, 10, 22, 45, 8, 8, 18, 18, 18, 18]
for col, w in enumerate(widths3, 1):
    ws3.column_dimensions[get_column_letter(col)].width = w

ws3.row_dimensions[1].height = 22
ws3.row_dimensions[2].height = 16
ws3.row_dimensions[4].height = 32
for r in range(5, 35):
    ws3.row_dimensions[r].height = 28

ws3.freeze_panes = 'A5'


# ── Save ────────────────────────────────────────────────────────────
wb.save(OUT_XLSX)
print(f"\nExcel saved: {OUT_XLSX} ({os.path.getsize(OUT_XLSX)} bytes)")


# ── Generate markdown analysis report ───────────────────────────────
md = []
md.append("# Топ-10 РТТ по продажам ОРЕОЛ TITAN\n")
md.append("**Источник данных:** `Титан Брозэкс, Рыгалов.xlsx`, вкладка «выгрузка» (3 055 строк)\n")
md.append("**Дистрибьюторы:** БРОЗЭКС (Екатеринбург) · Рыгалов С.В. (Пенза)\n")
md.append(f"**Всего уникальных РТТ с продажами TITAN:** {len(rtt_list)}\n")
md.append(f"**Общий оборот TITAN за 2025-2026:** {all_turnover:,.0f} ₽\n")
md.append("\n---\n\n")

# Top 10 by SKU
md.append("## 1. Топ-10 РТТ по количеству SKU ОРЕОЛ TITAN\n\n")
md.append("| # | Дистрибьютор | РТТ | Класс | SKU | Оборот 2025 (₽) | Оборот 2026 (₽) | Оборот всего (₽) | Ср. чек/SKU (₽) |\n")
md.append("|---:|---|---|---|---:|---:|---:|---:|---:||\n")
for i, r in enumerate(top_sku, 1):
    md.append(f"| {i} | {r['distributor']} | {r['rtt_name'][:60]} | {r['class_tt']} | **{r['sku_count']}** | {r['turnover_2025']:,.0f} | {r['turnover_2026']:,.0f} | {r['turnover_total']:,.0f} | {r['avg_check']:,.0f} |")
md.append(f"\n**Итог топ-10:** {total_sku_sum} SKU · {total_turnover_top10:,.0f} ₽ ({total_turnover_top10/all_turnover*100:.1f}% от общего оборота TITAN)\n")
md.append("\n---\n\n")

# Top 10 by turnover
md.append("## 2. Топ-10 РТТ по обороту ОРЕОЛ TITAN\n\n")
md.append("| # | Дистрибьютор | РТТ | Класс | SKU | Оборот 2025 (₽) | Оборот 2026 (₽) | Оборот всего (₽) | Ср. чек/SKU (₽) |\n")
md.append("|---:|---|---|---|---:|---:|---:|---:|---:||\n")
for i, r in enumerate(top_rev, 1):
    md.append(f"| {i} | {r['distributor']} | {r['rtt_name'][:60]} | {r['class_tt']} | {r['sku_count']} | {r['turnover_2025']:,.0f} | {r['turnover_2026']:,.0f} | **{r['turnover_total']:,.0f}** | {r['avg_check']:,.0f} |")
md.append(f"\n**Итог топ-10:** {total_sku_sum2} SKU · {total_turnover_top10_2:,.0f} ₽ ({total_turnover_top10_2/all_turnover*100:.1f}% от общего оборота TITAN)\n")
md.append("\n---\n\n")

# Cross-analysis
md.append("## 3. Анализ закономерностей\n\n")

# РТТ которые попали в оба топ-10
top_sku_set = set(r['rtt_name'] for r in top_sku)
top_rev_set = set(r['rtt_name'] for r in top_rev)
intersection = top_sku_set & top_rev_set
md.append(f"### 3.1. РТТ, попавшие в оба топ-10 ({len(intersection)} шт.)\n\n")
if intersection:
    md.append("Это «золотые» точки — лидируют и по ассортименту, и по обороту:\n\n")
    for r in rtt_list:
        if r['rtt_name'] in intersection:
            md.append(f"- **{r['rtt_name'][:60]}** ({r['distributor']}) — {r['sku_count']} SKU, {r['turnover_total']:,.0f} ₽\n")
md.append("\n")

# Distributor analysis
md.append("### 3.2. Распределение топ-10 по дистрибьюторам\n\n")
br_in_top_sku = sum(1 for r in top_sku if 'БРОЗЭКС' in r['distributor'])
rg_in_top_sku = sum(1 for r in top_sku if 'Рыгалов' in r['distributor'])
br_in_top_rev = sum(1 for r in top_rev if 'БРОЗЭКС' in r['distributor'])
rg_in_top_rev = sum(1 for r in top_rev if 'Рыгалов' in r['distributor'])
md.append(f"| Категория | БРОЗЭКС | Рыгалов |\n|---|---:|---:|\n")
md.append(f"| Топ-10 по SKU | {br_in_top_sku} | {rg_in_top_sku} |\n")
md.append(f"| Топ-10 по обороту | {br_in_top_rev} | {rg_in_top_rev} |\n")
md.append("\n")

# Class analysis
md.append("### 3.3. Распределение топ-10 по классам ТТ\n\n")
class_sku = defaultdict(int)
class_rev = defaultdict(int)
for r in top_sku:
    for c in r['class_tt'].split(','):
        c = c.strip()
        if c: class_sku[c] += 1
for r in top_rev:
    for c in r['class_tt'].split(','):
        c = c.strip()
        if c: class_rev[c] += 1
md.append(f"| Класс ТТ | Топ-10 по SKU | Топ-10 по обороту |\n|---|---:|---:|\n")
all_classes = sorted(set(list(class_sku.keys()) + list(class_rev.keys())))
for c in all_classes:
    md.append(f"| {c} | {class_sku.get(c, 0)} | {class_rev.get(c, 0)} |\n")
md.append("\n")

# Average check
md.append("### 3.4. Средний чек на SKU\n\n")
md.append("Показатель «Ср. чек/SKU» позволяет оценить глубину ассортимента в точке:\n\n")
avg_check_top_sku = sum(r['avg_check'] for r in top_sku) / 10
avg_check_top_rev = sum(r['avg_check'] for r in top_rev) / 10
avg_check_all = sum(r['avg_check'] for r in rtt_list) / len(rtt_list)
md.append(f"- Средний чек/SKU в **топ-10 по SKU**: {avg_check_top_sku:,.0f} ₽\n")
md.append(f"- Средний чек/SKU в **топ-10 по обороту**: {avg_check_top_rev:,.0f} ₽\n")
md.append(f"- Средний чек/SKU **по всем РТТ**: {avg_check_all:,.0f} ₽\n\n")

if avg_check_top_rev > avg_check_top_sku:
    md.append("Топ-10 по обороту имеет более высокий средний чек на SKU — это означает, что точки с меньшим ассортиментом, но более высокой оборачиваемостью каждой позиции, генерируют больше выручки.\n\n")
else:
    md.append("Топ-10 по SKU имеет более высокий средний чек на SKU — это означает, что ассортиментные точки также генерируют высокую выручку на каждую позицию.\n\n")

# Strategic insights
md.append("## 4. Стратегические выводы\n\n")
md.append("**На основе анализа топ-10 РТТ по двум критериям:**\n\n")

md.append("1. **Концентрация оборота.** ")
md.append(f"Топ-10 РТТ по обороту генерируют {total_turnover_top10_2/all_turnover*100:.1f}% общей выручки TITAN при {(sum(r['sku_count'] for r in top_rev)/all_sku_sum*100):.1f}% SKU. ")
md.append(f"Это означает, что {len(rtt_list) - 10} РТТ из {len(rtt_list)} генерируют лишь {(1 - total_turnover_top10_2/all_turnover)*100:.1f}% оборота — **высокая концентрация спроса**.\n\n")

md.append("2. **Ассортимент vs оборот.** ")
md.append(f"Топ-10 по количеству SKU имеет средний чек/SKU = {avg_check_top_sku:,.0f} ₽, ")
md.append(f"топ-10 по обороту — {avg_check_top_rev:,.0f} ₽. ")
if avg_check_top_rev > avg_check_top_sku * 1.2:
    md.append("**Топ по обороту сильно опережает топ по SKU по чеку** — это фокус-точки, где каждый артикул продаётся в больших объёмах.\n\n")
else:
    md.append("Средний чек сопоставим — ассортиментные точки и оборотные точки в значительной степени пересекаются.\n\n")

md.append("3. **Золотые РТТ.** ")
md.append(f"В оба топ-10 попали {len(intersection)} РТТ — это **приоритетные точки для удержания и расширения**. ")
md.append("Рекомендуется: 100% наличие топ-SKU, расширенная выкладка, приоритет в промо-программах.\n\n")

md.append("4. **Профиль идеальной РТТ.** ")
top_class = max(class_rev, key=class_rev.get) if class_rev else '?'
md.append(f"Преобладающий класс в топ-10 по обороту: **{top_class}**. ")
md.append("При расширении сети приоритет — точки этого класса с похожим профилем.\n\n")

md.append("---\n\n*Анализ подготовлен на основе выгрузки продаж БРОЗЭКС + Рыгалов С.В., 2025-2026*\n")

with open(OUT_MD, 'w', encoding='utf-8') as f:
    f.write('\n'.join(md))
print(f"Markdown report: {OUT_MD} ({os.path.getsize(OUT_MD)} bytes)")


# Console summary
print("\n" + "=" * 70)
print("TOP-10 BY SKU COUNT")
print("=" * 70)
print(f"{'#':3s} {'Дистрибьютор':25s} {'РТТ':50s} {'Кл':3s} {'SKU':4s} {'Оборот':12s}")
for i, r in enumerate(top_sku, 1):
    print(f"{i:3d} {r['distributor'][:25]:25s} {r['rtt_name'][:50]:50s} {r['class_tt'][:3]:3s} {r['sku_count']:4d} {r['turnover_total']:>12,.0f}")

print("\n" + "=" * 70)
print("TOP-10 BY TURNOVER")
print("=" * 70)
print(f"{'#':3s} {'Дистрибьютор':25s} {'РТТ':50s} {'Кл':3s} {'SKU':4s} {'Оборот':12s}")
for i, r in enumerate(top_rev, 1):
    print(f"{i:3d} {r['distributor'][:25]:25s} {r['rtt_name'][:50]:50s} {r['class_tt'][:3]:3s} {r['sku_count']:4d} {r['turnover_total']:>12,.0f}")

print(f"\nIntersection (both top-10): {len(intersection)} РТТ")
print(f"  → {intersection}")
