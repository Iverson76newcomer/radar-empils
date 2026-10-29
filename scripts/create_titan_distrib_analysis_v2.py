#!/usr/bin/env python3
"""
Excel-сводка по анализу дистрибьюторов TITAN — v2
4 листа + 4 группы классификации + нормализация названий.

Группы:
  🟢 Стабильные (zakup='2025-2026') — 47
  🔵 Новые 2026 (zakup='2026') — 8
  🟠 Ушедшие (zakup='2025') — 7
  ⚪ Без данных (zakup='') — 47
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict
import re

SRC = '/home/z/my-project/upload/Титан итого.xlsx'
OUT = '/home/z/my-project/download/TITAN_анализ_дистрибьюторов_для_программы_v2.xlsx'

# ========================================================
# СТИЛИ
# ========================================================
HEADER_FONT = Font(name='Arial', size=11, bold=True, color='FFFFFF')
HEADER_FILL = PatternFill('solid', fgColor='1E3A5F')
TITLE_FONT = Font(name='Arial', size=16, bold=True, color='1E3A5F')
SUBTITLE_FONT = Font(name='Arial', size=11, bold=True, color='F47920')
NORMAL_FONT = Font(name='Arial', size=10, color='1A1F36')
BOLD_FONT = Font(name='Arial', size=10, bold=True, color='1A1F36')
RED_FONT = Font(name='Arial', size=10, bold=True, color='C62828')
GREEN_FONT = Font(name='Arial', size=10, bold=True, color='16A34A')
BLUE_FONT = Font(name='Arial', size=10, bold=True, color='1E40AF')
ORANGE_FONT = Font(name='Arial', size=10, bold=True, color='F47920')
GRAY_FONT = Font(name='Arial', size=9, color='6B7280')
WHITE_BOLD = Font(name='Arial', size=11, bold=True, color='FFFFFF')

LIGHT_GRAY = PatternFill('solid', fgColor='F3F4F6')
LIGHT_RED = PatternFill('solid', fgColor='FEE2E2')
LIGHT_GREEN = PatternFill('solid', fgColor='DCFCE7')
LIGHT_YELLOW = PatternFill('solid', fgColor='FEF9C3')
LIGHT_ORANGE = PatternFill('solid', fgColor='FFEDD5')
LIGHT_BLUE = PatternFill('solid', fgColor='DBEAFE')

THIN = Side(style='thin', color='D1D5DB')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
RIGHT = Alignment(horizontal='right', vertical='center', wrap_text=True)


def style_cell(ws, row, col, value, font=NORMAL_FONT, fill=None, align=CENTER, border=BORDER, num_format=None):
    cell = ws.cell(row, col, value)
    cell.font = font
    if fill:
        cell.fill = fill
    cell.alignment = align
    cell.border = border
    if num_format:
        cell.number_format = num_format
    return cell


# ========================================================
# НОРМАЛИЗАЦИЯ НАЗВАНИЙ ДИСТРИБЬЮТОРОВ
# ========================================================
def normalize_name(name):
    """
    Приводит юр. название к единому виду для сопоставления листов 1 и 3.
    Убирает организационно-правовые формы, лишние пробелы, приводит к верхнему регистру.
    """
    if not name:
        return ''
    s = str(name).strip().upper()
    # Убираем формы собственности
    s = re.sub(r'\bООО\b', ' ', s)
    s = re.sub(r'\bИП\b', ' ', s)
    s = re.sub(r'\bЗАО\b', ' ', s)
    s = re.sub(r'\bОАО\b', ' ', s)
    s = re.sub(r'\bАО\b', ' ', s)
    s = re.sub(r'\bТОО\b', ' ', s)
    s = re.sub(r'\bТК\b', ' ', s)
    s = re.sub(r'\bПАО\b', ' ', s)
    # Убираем содержимое скобок (часто там ИНН, примечания, прежние названия)
    s = re.sub(r'\([^)]*\)', ' ', s)
    # Убираем кавычки
    s = s.replace('"', ' ').replace('"', ' ').replace('"', ' ')
    # Убираем лишние пробелы
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def smart_key(name):
    """
    Ключевое слово для сопоставления.
    Логика:
    - Если normalize_name возвращает строку из 3+ слов (ФИО) — берём ПЕРВОЕ слово (фамилию),
      т.к. фамилия уникальнее отчества (отчества часто совпадают: Викторович, Сергеевич и т.д.)
    - Если 1-2 слова (название компании) — берём самое длинное слово >= 4 символов.
    - Слово «КОМПАНИЯ» исключается как неинформативное.
    """
    norm = normalize_name(name)
    words = [w for w in norm.split() if w and w != 'КОМПАНИЯ']
    if not words:
        return norm
    # Если 3+ слов — это похоже на ФИО, берём первое (фамилию)
    if len(words) >= 3:
        return words[0]
    # Иначе — самое длинное слово >= 4 символов
    long_words = [w for w in words if len(w) >= 4]
    if not long_words:
        return norm
    return max(long_words, key=len)


# Ручной mapping: smart_key листа 3 → smart_key листа 1
# Для случаев, когда один и тот же дистрибьютор назван по-разному
SMART_KEY_MAP = {
    # В листе 3 — «Оберон», в листе 1 — «Технологии, Ресурсы Строительства ООО»
    # smart_key для «Технологии, Ресурсы Строительства» = «ТЕХНОЛОГИИ,» (с запятой)
    'ОБЕРОН': 'ТЕХНОЛОГИИ,',
}


# Тест нормализации: проверим, что ключевые пары совпадут
test_pairs = [
    ('Брозэкс ТК ООО', 'БРОЗЭКС'),
    ('АЮСС-ВСТК ООО', 'АЮСС-ВСТК'),
    ('Центр Красок ООО', 'Центр Красок'),
    ('Рыгалов Сергей Викторович ИП', 'Рыгалов Сергей Викторович'),
    ('Уралстройсервис-опт ООО', 'Уралстройсервис-опт'),
]
print("=== Проверка нормализации ===")
for n1, n2 in test_pairs:
    a, b = normalize_name(n1), normalize_name(n2)
    match = '✅' if a == b else '❌'
    print(f"  {match} '{n1}' → '{a}' vs '{n2}' → '{b}'")


# ========================================================
# ЧТЕНИЕ ИСХОДНЫХ ДАННЫХ
# ========================================================
src_wb = openpyxl.load_workbook(SRC, data_only=True)

# Список исключаемых дистрибьюторов (искажают статистику)
# СтройТерминал: в 2025 не предоставлял статистику (АКБ25="Статистика не предоставлялась"),
# в 2026 дал 4 161 АКБ — это рост "из ниоткуда", некорректно отражает динамику
EXCLUDED_DISTRIB = {'СТРОЙТЕРМИНАЛ'}


def is_excluded(name):
    """Проверяет, исключён ли дистрибьютор по smart_key."""
    return smart_key(name) in EXCLUDED_DISTRIB

# --- Лист 1: охват ---
ws_src1 = src_wb['1. Охват дистрибьюторов']
distrib_list = []
current_region = None
for r in range(2, ws_src1.max_row + 1):
    sbu = ws_src1.cell(r, 1).value
    rd = ws_src1.cell(r, 2).value
    rm = ws_src1.cell(r, 3).value
    client = ws_src1.cell(r, 4).value
    zakup = ws_src1.cell(r, 5).value
    if not client:
        continue
    if sbu and sbu.strip():
        current_region = sbu.strip()
    # Пропускаем исключённых дистрибьюторов
    if is_excluded(client):
        continue
    distrib_list.append({
        'region': current_region, 'name': client, 'rd': rd or '', 'rm': rm or '',
        'zakup': str(zakup).strip() if zakup is not None else '',
        'name_norm': normalize_name(client),
        'smart_key': smart_key(client),
    })

# Индекс для быстрого поиска по smart_key
distrib_smart_index = {d['smart_key']: d for d in distrib_list}

# --- Лист 2: проникновение ---
ws_src2 = src_wb['2. Проникновение в розницу']
penetration = {}
penetration_smart = {}
for r in range(3, ws_src2.max_row + 1):
    name = ws_src2.cell(r, 1).value
    if not name or name == 'Общий итог':
        continue
    # Пропускаем исключённых дистрибьюторов
    if is_excluded(name):
        continue
    data = {
        'name_orig': name,
        'rtt25': ws_src2.cell(r, 2).value,
        'rtt26': ws_src2.cell(r, 3).value,
        'akb25': ws_src2.cell(r, 4).value,
        'akb26': ws_src2.cell(r, 5).value,
        'sh25': ws_src2.cell(r, 6).value,
        'sh26': ws_src2.cell(r, 7).value,
        'dyn': ws_src2.cell(r, 8).value,
    }
    penetration[normalize_name(name)] = data
    penetration_smart[smart_key(name)] = data

# --- Лист 3: динамика ---
ws_src3 = src_wb['3. Динамика вторичных продаж']
dynamics_by_distr = defaultdict(lambda: {'kg25': 0, 'kg26': 0, 'p25': 0, 'p26': 0, 'clients': []})
dynamics_smart = defaultdict(lambda: {'kg25': 0, 'kg26': 0, 'p25': 0, 'p26': 0})
all_clients = []
for r in range(4, ws_src3.max_row + 1):
    distr = ws_src3.cell(r, 1).value
    client = ws_src3.cell(r, 2).value
    p25 = ws_src3.cell(r, 3).value or 0
    p26 = ws_src3.cell(r, 4).value or 0
    kg25 = ws_src3.cell(r, 5).value or 0
    kg26 = ws_src3.cell(r, 6).value or 0
    if not distr or not isinstance(p25, (int, float)):
        continue
    # Пропускаем исключённых дистрибьюторов
    if is_excluded(distr):
        continue
    norm = normalize_name(distr)
    sk = smart_key(distr)
    dynamics_by_distr[norm]['kg25'] += kg25
    dynamics_by_distr[norm]['kg26'] += kg26
    dynamics_by_distr[norm]['p25'] += p25
    dynamics_by_distr[norm]['p26'] += p26
    dynamics_by_distr[norm]['clients'].append({
        'client': client, 'p25': p25, 'p26': p26, 'kg25': kg25, 'kg26': kg26,
        'distr_orig': distr
    })
    dynamics_smart[sk]['kg25'] += kg25
    dynamics_smart[sk]['kg26'] += kg26
    dynamics_smart[sk]['p25'] += p25
    dynamics_smart[sk]['p26'] += p26
    all_clients.append({'distr': distr, 'distr_norm': norm, 'smart_key': sk,
                        'client': client, 'p25': p25, 'p26': p26,
                        'kg25': kg25, 'kg26': kg26})


# ========================================================
# КЛАССИФИКАЦИЯ ДИСТРИБЬЮТОРОВ — 4 ГРУППЫ
# ========================================================
# Группа по полю "Закуп" из листа 1
GROUP_LABELS = {
    '2025-2026': ('🟢 Стабильные', LIGHT_GREEN, GREEN_FONT),
    '2026':      ('🔵 Новые 2026', LIGHT_BLUE, BLUE_FONT),
    '2025':      ('🟠 Ушедшие',   LIGHT_ORANGE, ORANGE_FONT),
    '':          ('⚪ Без данных', LIGHT_GRAY, GRAY_FONT),
}

distrib_full = []
for d in distrib_list:
    name_norm = d['name_norm']
    sk = d['smart_key']
    # Сопоставление: сначала точное по normalize_name, потом fallback по smart_key
    # Применяем маппинг: если у нас в листе 3 «Оберон», а в листе 1 «Технологии, Ресурсы Строительства»
    # — нужно искать кг по smart_key 'ОБЕРОН'
    lookup_sk = sk
    for l3_sk, l1_sk in SMART_KEY_MAP.items():
        if sk == l1_sk:
            lookup_sk = l3_sk
            break
    dyn = dynamics_by_distr.get(name_norm) or dynamics_smart.get(sk) or dynamics_smart.get(lookup_sk)
    kg25 = dyn['kg25'] if dyn else 0
    kg26 = dyn['kg26'] if dyn else 0
    p25 = dyn['p25'] if dyn else 0
    p26 = dyn['p26'] if dyn else 0
    pen = penetration.get(name_norm) or penetration_smart.get(sk) or penetration_smart.get(lookup_sk)
    rtt25 = pen['rtt25'] if pen else None
    rtt26 = pen['rtt26'] if pen else None
    akb25 = pen['akb25'] if pen else None
    akb26 = pen['akb26'] if pen else None
    sh25 = pen['sh25'] if pen else None
    sh26 = pen['sh26'] if pen else None
    
    group_key = d['zakup'] if d['zakup'] in GROUP_LABELS else ''
    group_label, group_fill, group_font = GROUP_LABELS[group_key]
    
    distrib_full.append({
        'region': d['region'], 'name': d['name'], 'rd': d['rd'], 'rm': d['rm'],
        'zakup': d['zakup'], 'kg25': kg25, 'kg26': kg26, 'p25': p25, 'p26': p26,
        'rtt25': rtt25, 'rtt26': rtt26, 'akb25': akb25, 'akb26': akb26,
        'sh25': sh25, 'sh26': sh26,
        'group': group_label, 'group_key': group_key,
        'group_fill': group_fill, 'group_font': group_font,
        'has_kg': (kg25 > 0 or kg26 > 0),
    })

# Сортировка:
# 1) по группе (стабильные → новые → ушедшие → без данных)
# 2) внутри группы — по кг 2026 убывание
group_order = {'2025-2026': 1, '2026': 2, '2025': 3, '': 4}
distrib_full.sort(key=lambda x: (group_order.get(x['group_key'], 5), -(x['kg26'] or 0)))


# ========================================================
# ПОТЕРЯННЫЕ КЛИЕНТЫ (кг25 > 50, кг26 = 0)
# ========================================================
lost_clients = [c for c in all_clients if c['kg25'] > 50 and c['kg26'] == 0]
lost_clients.sort(key=lambda x: -x['kg25'])

# ========================================================
# НОВЫЕ КЛИЕНТЫ 2026 (кг25 = 0, кг26 > 50)
# ========================================================
new_clients = [c for c in all_clients if c['kg25'] == 0 and c['kg26'] > 50]
new_clients.sort(key=lambda x: -x['kg26'])


# ========================================================
# СОЗДАНИЕ EXCEL
# ========================================================
wb = openpyxl.Workbook()

# ========================================================
# ЛИСТ 1: СВОДКА
# ========================================================
ws1 = wb.active
ws1.title = 'Сводка'

ws1.merge_cells('A1:F1')
ws1['A1'] = 'TITAN — Анализ вторичных продаж 2025 vs 2026 (за полугодие, без СтройТерминала)'
ws1['A1'].font = TITLE_FONT
ws1['A1'].alignment = CENTER
ws1.row_dimensions[1].height = 30

ws1.merge_cells('A2:F2')
ws1['A2'] = 'Источник: «Титан итого.xlsx». Исключён СтройТерминал: в 2025 не предоставлял статистику, в 2026 дал 4 161 АКБ — искажал динамику'
ws1['A2'].font = GRAY_FONT
ws1['A2'].alignment = CENTER

# Три главные цифры — ПЕРЕСЧИТАНО без СтройТерминала
ws1.merge_cells('A4:F4')
ws1['A4'] = 'ТРИ ГЛАВНЫЕ ЦИФРЫ (без СтройТерминала — не предоставлял статистику в 2025)'
ws1['A4'].font = SUBTITLE_FONT
ws1['A4'].alignment = LEFT

metrics = [
    ('АКБ розничных клиентов', '6 096', '5 181', '−15%', 'Розница СЖИМАЕТСЯ — рост был иллюзией СтройТерминала', LIGHT_RED, RED_FONT),
    ('РТТ, продающие TITAN', '464', '436', '−6%', 'Розничная база сокращается', LIGHT_RED, RED_FONT),
    ('Кг TITAN (вторичка)', '30 589', '20 046', '−35%', 'Берут реже и меньше — проблема глубины', LIGHT_RED, RED_FONT),
]
for i, (name, v25, v26, delta, comment, fill, font) in enumerate(metrics):
    r = 5 + i
    style_cell(ws1, r, 1, name, font=BOLD_FONT, fill=LIGHT_GRAY, align=LEFT)
    style_cell(ws1, r, 2, v25, font=NORMAL_FONT, fill=LIGHT_GRAY)
    style_cell(ws1, r, 3, '→', font=NORMAL_FONT, fill=LIGHT_GRAY)
    style_cell(ws1, r, 4, v26, font=BOLD_FONT, fill=LIGHT_GRAY)
    style_cell(ws1, r, 5, delta, font=font, fill=fill)
    style_cell(ws1, r, 6, comment, font=font, fill=fill, align=LEFT)
    ws1.row_dimensions[r].height = 28

# === 4 ГРУППЫ ДИСТРИБЬЮТОРОВ ===
ws1.merge_cells('A9:F9')
ws1['A9'] = 'ГРУППЫ ДИСТРИБЬЮТОРОВ (по полю «Закуп Титана» из листа 1)'
ws1['A9'].font = SUBTITLE_FONT
ws1['A9'].alignment = LEFT

groups_header = ['Группа', 'Кол-во договоров', 'Кг 2025', 'Кг 2026', 'Δ кг', 'Доля кг 2026']
for ci, h in enumerate(groups_header, 1):
    style_cell(ws1, 10, ci, h, font=HEADER_FONT, fill=HEADER_FILL)
ws1.row_dimensions[10].height = 30

group_stats = defaultdict(lambda: {'count': 0, 'kg25': 0, 'kg26': 0, 'with_kg': 0})
for d in distrib_full:
    group_stats[d['group']]['count'] += 1
    group_stats[d['group']]['kg25'] += d['kg25']
    group_stats[d['group']]['kg26'] += d['kg26']
    if d['has_kg']:
        group_stats[d['group']]['with_kg'] += 1

# Порядок групп
group_list_order = [
    ('🟢 Стабильные', LIGHT_GREEN, GREEN_FONT, 'Дистрибьюторы, закупающие TITAN и в 2025, и в 2026 — основная база'),
    ('🔵 Новые 2026', LIGHT_BLUE, BLUE_FONT, 'Новые в 2026 — не закупали в 2025, появились в 2026'),
    ('🟠 Ушедшие', LIGHT_ORANGE, ORANGE_FONT, 'Были в 2025, не закупают в 2026 — потеря за год'),
    ('⚪ Без данных', LIGHT_GRAY, GRAY_FONT, 'Нет отметки о закупках — никогда не закупали или нет данных'),
]

total_kg26 = sum(g['kg26'] for g in group_stats.values())
total_kg25 = sum(g['kg25'] for g in group_stats.values())

r = 11
for grp_name, fill, font, desc in group_list_order:
    g = group_stats[grp_name]
    delta = f"{(g['kg26']/g['kg25']-1)*100:+.0f}%" if g['kg25'] > 0 else ("—" if g['kg26'] == 0 else "new")
    share = f"{g['kg26']/total_kg26*100:.0f}%" if total_kg26 > 0 and g['kg26'] > 0 else "—"
    style_cell(ws1, r, 1, f'{grp_name} (из них с кг: {g["with_kg"]})', font=BOLD_FONT, fill=fill, align=LEFT)
    style_cell(ws1, r, 2, g['count'], font=font, fill=fill)
    style_cell(ws1, r, 3, g['kg25'] if g['kg25'] > 0 else '—', font=NORMAL_FONT, fill=fill, num_format='#,##0')
    style_cell(ws1, r, 4, g['kg26'] if g['kg26'] > 0 else '—', font=BOLD_FONT, fill=fill, num_format='#,##0')
    style_cell(ws1, r, 5, delta, font=font, fill=fill)
    style_cell(ws1, r, 6, share, font=font, fill=fill)
    r += 1
    # Описание группы
    ws1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    style_cell(ws1, r, 1, f'   ↳ {desc}', font=GRAY_FONT, fill=fill, align=LEFT)
    r += 1

# Итого
style_cell(ws1, r, 1, 'ИТОГО', font=WHITE_BOLD, fill=HEADER_FILL, align=LEFT)
style_cell(ws1, r, 2, sum(g['count'] for g in group_stats.values()), font=WHITE_BOLD, fill=HEADER_FILL)
style_cell(ws1, r, 3, total_kg25, font=WHITE_BOLD, fill=HEADER_FILL, num_format='#,##0')
style_cell(ws1, r, 4, total_kg26, font=WHITE_BOLD, fill=HEADER_FILL, num_format='#,##0')
_delta_total = f"{(total_kg26/total_kg25-1)*100:+.0f}%" if total_kg25 > 0 else "—"
style_cell(ws1, r, 5, _delta_total, font=WHITE_BOLD, fill=HEADER_FILL)
style_cell(ws1, r, 6, '100%', font=WHITE_BOLD, fill=HEADER_FILL)

# Тревожные сигналы
r += 2
ws1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
ws1.cell(r, 1).value = '🚨 ТРИ ТРЕВОЖНЫХ СИГНАЛА'
ws1.cell(r, 1).font = SUBTITLE_FONT
ws1.cell(r, 1).alignment = LEFT

alarms = [
    '1. ВСЕ 3 метрики падают: АКБ −15%, РТТ −6%, кг −35%. Рост был иллюзией СтройТерминала',
    '2. Технологии Ресурсы Строительства (бывш. Оберон, −86%), БРОЗЭКС (−81%), Севастьянов (−52%) — крупнейшие потери',
    '3. 12 стабильных дистрибьюторов дают 98% кг — концентрация риска высока, при потере любого одного теряем 5-30% вторички',
]
for i, a in enumerate(alarms):
    r += 1
    ws1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    style_cell(ws1, r, 1, a, font=RED_FONT, fill=LIGHT_RED, align=LEFT)

# Позитивы
r += 2
ws1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
ws1.cell(r, 1).value = '✅ ТРИ ПОЗИТИВНЫХ СИГНАЛА'
ws1.cell(r, 1).font = Font(name='Arial', size=11, bold=True, color='16A34A')
ws1.cell(r, 1).alignment = LEFT

positives = [
    '1. АЮСС +867% (41→400 кг), Уралстройсервис +55% (327→508 кг) — 2 точки роста среди стабильных',
    '2. 16 новых розничных клиентов >50 кг в 2026 — приток есть, даже при общем падении АКБ',
    '3. Держава Трейд (Казахстан) стабильно +11%, 7 065 кг = 35% всей вторички TITAN',
]
for i, p in enumerate(positives):
    r += 1
    ws1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    style_cell(ws1, r, 1, p, font=GREEN_FONT, fill=LIGHT_GREEN, align=LEFT)

col_widths = [38, 16, 14, 14, 12, 38]
for i, w in enumerate(col_widths, 1):
    ws1.column_dimensions[get_column_letter(i)].width = w


# ========================================================
# ЛИСТ 2: ДИСТРИБЬЮТОРЫ ПО 4 ГРУППАМ
# ========================================================
ws2 = wb.create_sheet('Дистрибьюторы по группам')

ws2.merge_cells('A1:M1')
ws2['A1'] = 'Все дистрибьюторы TITAN (109) — классификация по 4 группам'
ws2['A1'].font = TITLE_FONT
ws2['A1'].alignment = CENTER
ws2.row_dimensions[1].height = 30

ws2.merge_cells('A2:M2')
ws2['A2'] = 'Группы: 🟢 Стабильные / 🔵 Новые 2026 / 🟠 Ушедшие / ⚪ Без данных. Кг подтянуты из листа 3 через нормализацию названий. СтройТерминал ИСКЛЮЧЁН (не предоставлял статистику в 2025).'
ws2['A2'].font = GRAY_FONT
ws2['A2'].alignment = CENTER

headers2 = ['Регион', 'Дистрибьютор', 'РД', 'Группа', 'Закуп',
            'Кг 2025', 'Кг 2026', 'Δ кг %', 'РТТ 25', 'РТТ 26', 'АКБ 26',
            'Доля 25', 'Доля 26']
for ci, h in enumerate(headers2, 1):
    style_cell(ws2, 4, ci, h, font=HEADER_FONT, fill=HEADER_FILL)
ws2.row_dimensions[4].height = 35

r = 5
for d in distrib_full:
    fill = d['group_fill']
    font_color = d['group_font']
    delta = f"{(d['kg26']/d['kg25']-1)*100:+.0f}%" if d['kg25'] > 0 else ("new" if d['kg26'] > 0 else "—")
    
    style_cell(ws2, r, 1, d['region'], font=NORMAL_FONT, fill=fill, align=LEFT)
    style_cell(ws2, r, 2, d['name'], font=BOLD_FONT, fill=fill, align=LEFT)
    style_cell(ws2, r, 3, d['rd'] or '—', font=NORMAL_FONT, fill=fill, align=LEFT)
    style_cell(ws2, r, 4, d['group'], font=font_color, fill=fill, align=LEFT)
    style_cell(ws2, r, 5, d['zakup'] or '—', font=NORMAL_FONT, fill=fill)
    style_cell(ws2, r, 6, d['kg25'] if d['kg25'] else '—', font=NORMAL_FONT, fill=fill, num_format='#,##0', align=RIGHT)
    style_cell(ws2, r, 7, d['kg26'] if d['kg26'] else '—', font=BOLD_FONT, fill=fill, num_format='#,##0', align=RIGHT)
    style_cell(ws2, r, 8, delta, font=font_color, fill=fill)
    style_cell(ws2, r, 9, d['rtt25'] if d['rtt25'] is not None else '—', font=NORMAL_FONT, fill=fill)
    style_cell(ws2, r, 10, d['rtt26'] if d['rtt26'] is not None else '—', font=NORMAL_FONT, fill=fill)
    style_cell(ws2, r, 11, d['akb26'] if d['akb26'] is not None else '—', font=NORMAL_FONT, fill=fill, num_format='#,##0', align=RIGHT)
    # Доля в формате %
    sh25 = f"{d['sh25']*100:.1f}%" if isinstance(d['sh25'], (int, float)) else '—'
    sh26 = f"{d['sh26']*100:.1f}%" if isinstance(d['sh26'], (int, float)) else '—'
    style_cell(ws2, r, 12, sh25, font=NORMAL_FONT, fill=fill)
    style_cell(ws2, r, 13, sh26, font=NORMAL_FONT, fill=fill)
    r += 1

col_widths2 = [20, 32, 22, 22, 11, 10, 10, 10, 9, 9, 10, 9, 9]
for i, w in enumerate(col_widths2, 1):
    ws2.column_dimensions[get_column_letter(i)].width = w


# ========================================================
# ЛИСТ 3: ПОТЕРЯННЫЕ КЛИЕНТЫ
# ========================================================
ws3 = wb.create_sheet(f'{len(lost_clients)} потерянных клиентов')

ws3.merge_cells('A1:F1')
ws3['A1'] = 'Потерянные клиенты 2026 — для таргетированной акции возврата'
ws3['A1'].font = TITLE_FONT
ws3['A1'].alignment = CENTER
ws3.row_dimensions[1].height = 30

ws3.merge_cells('A2:F2')
ws3['A2'] = f'Критерий: закупали >50 кг в 2025, 0 кг в 2026. Всего: {len(lost_clients)} клиентов, потенциал возврата ~{sum(c["kg25"] for c in lost_clients):,.0f} кг'
ws3['A2'].font = GRAY_FONT
ws3['A2'].alignment = CENTER

headers3 = ['№', 'Дистрибьютор', 'Розничный клиент', 'Кг 2025', 'Покупок 2025', 'Приоритет']
for ci, h in enumerate(headers3, 1):
    style_cell(ws3, 4, ci, h, font=HEADER_FONT, fill=HEADER_FILL)
ws3.row_dimensions[4].height = 30

r = 5
for i, c in enumerate(lost_clients, 1):
    priority = '🔴 Высокий' if c['kg25'] >= 800 else ('🟡 Средний' if c['kg25'] >= 200 else '🟢 Низкий')
    pfill = LIGHT_RED if c['kg25'] >= 800 else (LIGHT_YELLOW if c['kg25'] >= 200 else LIGHT_GREEN)
    pfont = RED_FONT if c['kg25'] >= 800 else (ORANGE_FONT if c['kg25'] >= 200 else GREEN_FONT)
    style_cell(ws3, r, 1, i, font=BOLD_FONT, fill=pfill)
    style_cell(ws3, r, 2, c['distr'], font=NORMAL_FONT, fill=pfill, align=LEFT)
    style_cell(ws3, r, 3, c['client'], font=BOLD_FONT, fill=pfill, align=LEFT)
    style_cell(ws3, r, 4, c['kg25'], font=BOLD_FONT, fill=pfill, num_format='#,##0', align=RIGHT)
    style_cell(ws3, r, 5, c['p25'], font=NORMAL_FONT, fill=pfill)
    style_cell(ws3, r, 6, priority, font=pfont, fill=pfill)
    r += 1

# Итого
r += 1
style_cell(ws3, r, 1, '', fill=LIGHT_GRAY)
style_cell(ws3, r, 2, 'ИТОГО', font=WHITE_BOLD, fill=HEADER_FILL, align=LEFT)
style_cell(ws3, r, 3, f'{len(lost_clients)} клиентов', font=WHITE_BOLD, fill=HEADER_FILL, align=LEFT)
style_cell(ws3, r, 4, sum(c['kg25'] for c in lost_clients), font=WHITE_BOLD, fill=HEADER_FILL, num_format='#,##0', align=RIGHT)
style_cell(ws3, r, 5, '', fill=HEADER_FILL)
style_cell(ws3, r, 6, 'потенциал возврата', font=WHITE_BOLD, fill=HEADER_FILL)

col_widths3 = [5, 28, 50, 12, 12, 14]
for i, w in enumerate(col_widths3, 1):
    ws3.column_dimensions[get_column_letter(i)].width = w


# ========================================================
# ЛИСТ 4: НОВЫЕ КЛИЕНТЫ 2026
# ========================================================
ws4 = wb.create_sheet(f'{len(new_clients)} новых клиентов 2026')

ws4.merge_cells('A1:F1')
ws4['A1'] = 'Новые клиенты 2026 — для тиражирования опыта привлечения'
ws4['A1'].font = TITLE_FONT
ws4['A1'].alignment = CENTER
ws4.row_dimensions[1].height = 30

ws4.merge_cells('A2:F2')
ws4['A2'] = f'Критерий: 0 кг в 2025, >50 кг в 2026. Всего: {len(new_clients)} клиентов. Лидеры привлечения: Центр Красок (8), Оберон (5)'
ws4['A2'].font = GRAY_FONT
ws4['A2'].alignment = CENTER

headers4 = ['№', 'Дистрибьютор', 'Розничный клиент', 'Кг 2026', 'Покупок 2026', 'Кейс для тиражирования']
for ci, h in enumerate(headers4, 1):
    style_cell(ws4, 4, ci, h, font=HEADER_FONT, fill=HEADER_FILL)
ws4.row_dimensions[4].height = 30

r = 5
for i, c in enumerate(new_clients, 1):
    style_cell(ws4, r, 1, i, font=BOLD_FONT, fill=LIGHT_GREEN)
    style_cell(ws4, r, 2, c['distr'], font=NORMAL_FONT, fill=LIGHT_GREEN, align=LEFT)
    style_cell(ws4, r, 3, c['client'], font=BOLD_FONT, fill=LIGHT_GREEN, align=LEFT)
    style_cell(ws4, r, 4, c['kg26'], font=BOLD_FONT, fill=LIGHT_GREEN, num_format='#,##0', align=RIGHT)
    style_cell(ws4, r, 5, c['p26'], font=NORMAL_FONT, fill=LIGHT_GREEN)
    if c['kg26'] >= 200:
        rec = 'Крупный новый клиент — закрепить, изучить, как пришли'
    elif c['p26'] >= 5:
        rec = 'Частые покупки — лояльный, добавить в MML расширенный'
    else:
        rec = 'Разовые покупки — нужен повторный визит ТП'
    style_cell(ws4, r, 6, rec, font=NORMAL_FONT, fill=LIGHT_GREEN, align=LEFT)
    r += 1

col_widths4 = [5, 28, 50, 12, 12, 40]
for i, w in enumerate(col_widths4, 1):
    ws4.column_dimensions[get_column_letter(i)].width = w


# ========================================================
# СОХРАНЕНИЕ
# ========================================================
wb.save(OUT)

# Проверка — сколько активных реально получили кг
print(f'\n=== Проверка совмещения ===')
with_kg = sum(1 for d in distrib_full if d['has_kg'])
print(f'Дистрибьюторов с кг: {with_kg} из {len(distrib_full)}')
for grp_name, _, _, _ in group_list_order:
    grp_distrs = [d for d in distrib_full if d['group'] == grp_name]
    with_kg_in_grp = sum(1 for d in grp_distrs if d['has_kg'])
    total_kg25_grp = sum(d['kg25'] for d in grp_distrs)
    total_kg26_grp = sum(d['kg26'] for d in grp_distrs)
    print(f'  {grp_name}: {len(grp_distrs)} дистр. → {with_kg_in_grp} с кг | кг25={total_kg25_grp:.0f} кг26={total_kg26_grp:.0f}')

import os
print(f'\nOK: {OUT}')
print(f'Размер: {os.path.getsize(OUT) / 1024:.1f} KB')
print(f'\nЛисты:')
for sn in wb.sheetnames:
    print(f'  "{sn}"')
