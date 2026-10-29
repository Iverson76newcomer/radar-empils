#!/usr/bin/env python3
"""
Анализ ассортимента Петрович и создание технологических цепочек.
"""
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill, numbers
from openpyxl.utils import get_column_letter
from copy import copy

# ============================================================
# 1. PARSE ASSORTMENT
# ============================================================
wb_src = openpyxl.load_workbook('/home/z/my-project/upload/Петрович.xlsx')
ws_src = wb_src.worksheets[0]

products = []
seen_names = set()

# Left block (columns A-P)
for row in range(3, ws_src.max_row + 1):
    name = ws_src.cell(row=row, column=2).value
    if name and name not in seen_names:
        code = ws_src.cell(row=row, column=3).value
        status = ws_src.cell(row=row, column=4).value
        seen_names.add(name)
        products.append({
            'name': name.strip(),
            'code': code,
            'status': status,
        })

# Right block (columns Q-AF)
for row in range(3, ws_src.max_row + 1):
    name = ws_src.cell(row=row, column=18).value
    if name and name not in seen_names:
        code = ws_src.cell(row=row, column=19).value
        seen_names.add(name)
        products.append({
            'name': name.strip(),
            'code': code,
            'status': None,
        })

print(f"Total unique products: {len(products)}")

# ============================================================
# 2. CLASSIFY PRODUCTS
# ============================================================
def classify(name):
    n = name.upper()
    if 'ГРУНТ-ЭМАЛЬ' in n and '3В1' in n:
        return 'Грунт-эмаль 3в1'
    if 'ГРУНТОВКА ГФ-021' in n or 'ГФ-021' in n:
        return 'Грунтовка ГФ-021'
    if 'ПФ-115' in n and 'ЭКОНОМ' in n:
        return 'Эмаль ПФ-115 ЭКОНОМ'
    if 'ПФ-115' in n and 'КМ' in n:
        return 'Эмаль ПФ-115 КМ'
    if 'ПФ-115' in n and 'РАСЦВЕТ' in n:
        return 'Эмаль ПФ-115 Расцвет'
    if 'ПФ-115' in n and 'СТРОЙКОМПЛЕКТ' in n:
        return 'Эмаль ПФ-115 Стройкомплект'
    if 'ПФ-115' in n and 'МАТОВАЯ' in n:
        return 'Эмаль ПФ-115 матовая'
    if 'ПФ-115' in n:
        return 'Эмаль ПФ-115'
    if 'НЦ-132' in n:
        return 'Эмаль НЦ-132КП'
    if 'КРАСКА МАСЛЯНАЯ' in n:
        return 'Краска масляная'
    if 'ЛАК' in n and 'ПАРКЕТ' in n:
        return 'Лак паркетный'
    if 'БЕТОНН' in n or 'ПОЛ' in n and 'АК' in n:
        return 'Эмаль для пола'
    return 'Прочее'

for p in products:
    p['category'] = classify(p['name'])

# Count by category
from collections import Counter
cat_count = Counter(p['category'] for p in products)
print("\nCategory breakdown:")
for cat, cnt in cat_count.most_common():
    print(f"  {cat}: {cnt}")

# ============================================================
# 3. EXTRACT DETAILS (weight, color, brand line)
# ============================================================
import re

def extract_details(name):
    n = name
    # Weight
    weight_match = re.search(r'(\d+[\.,]?\d*)\s*кг', n)
    weight = float(weight_match.group(1).replace(',', '.')) if weight_match else None
    
    # Color
    colors_map = {
        'белая': 'Белая', 'белый': 'Белый', 'белое': 'Белое',
        'чёрная': 'Чёрная', 'черная': 'Чёрная', 'чёрный': 'Чёрный',
        'серая': 'Серая', 'серый': 'Серый',
        'синяя': 'Синяя', 'синий': 'Синий',
        'зелёная': 'Зелёная', 'зеленая': 'Зелёная', 'зелёный': 'Зелёный', 'зеленый': 'Зелёный',
        'красная': 'Красная', 'красный': 'Красный',
        'жёлтая': 'Жёлтая', 'желтая': 'Жёлтая', 'жёлтый': 'Жёлтый',
        'коричневая': 'Коричневая', 'коричневый': 'Коричневый',
        'красно-коричневая': 'Красно-коричневая',
        'шоколадная': 'Шоколадная',
        'голубая': 'Голубая',
        'золотисто-желтая': 'Золотисто-жёлтая',
        'светло-серая': 'Светло-серая',
        'сурик': 'Сурик',
        'бежевая': 'Бежевая',
        'матовая': 'Матовая',
    }
    color = None
    for cword, cname in colors_map.items():
        if cword in n.lower():
            color = cname
            break
    
    return {'weight': weight, 'color': color}

for p in products:
    det = extract_details(p['name'])
    p['weight'] = det['weight']
    p['color'] = det['color']

# ============================================================
# 4. BUILD TECHNOLOGICAL CHAINS
# ============================================================
chains = []

# Group products by category for easy access
by_cat = {}
for p in products:
    cat = p['category']
    if cat not in by_cat:
        by_cat[cat] = []
    by_cat[cat].append(p)

# --- CHAIN 1: Металл улица — полная цепочка с ГФ-021 ---
if 'Грунтовка ГФ-021' in by_cat and 'Эмаль ПФ-115 КМ' in by_cat:
    gf = [p for p in by_cat['Грунтовка ГФ-021'] if 'красно-коричнев' in (p.get('color') or '').lower() or 'красно-коричнев' in p['name'].lower()]
    pf115_km = by_cat['Эмаль ПФ-115 КМ']
    if gf and pf115_km:
        chains.append({
            'id': 1,
            'name': 'Металл на улице: полная защита',
            'surface': 'Металл (улица)',
            'purpose': 'Защита металлических конструкций от коррозии в атмосферных условиях',
            'layers': [
                {'step': 1, 'product': 'Грунтовка ГФ-021 Расцвет красно-коричневая 12 кг', 'role': 'Грунт', 'consumption': '60-80 г/м²', 'layers': 1, 'drying': '24 ч'},
                {'step': 2, 'product': 'Эмаль ПФ-115 КМ (цвет на выбор)', 'role': 'Покрытие', 'consumption': '100-140 г/м²', 'layers': 2, 'drying': '24 ч/слой'},
            ],
            'tco_note': 'Без грунта: перекраска через 1-2 года. С грунтом: 4-5 лет. Экономия до 40% TCO.',
            'segment': 'Универсал',
            'priority': 'Высокая',
        })

# --- CHAIN 2: Металл — быстрое решение 3в1 ---
if 'Грунт-эмаль 3в1' in by_cat:
    ge = by_cat['Грунт-эмаль 3в1']
    colors = list(set(p['color'] for p in ge if p['color']))
    chains.append({
        'id': 2,
        'name': 'Металл: быстрое решение (3в1 ПРОСТОКРАШЕНО!)',
        'surface': 'Металл (улица/помещение)',
        'purpose': 'Быстрая защита металла — грунт, преобразователь ржавчины и эмаль в одном слое',
        'layers': [
            {'step': 1, 'product': f'Грунт-эмаль ПРОСТОКРАШЕНО! 3в1 ({", ".join(colors)})', 'role': 'Грунт+эмаль', 'consumption': '120-160 г/м²', 'layers': 2, 'drying': '2-4 ч/слой'},
        ],
        'tco_note': '3в1 = экономия времени 50%. Не требует грунтовки. Идеально для заборов, ворот, труб.',
        'segment': 'Эконом / Универсал',
        'priority': 'Высокая',
    })

# --- CHAIN 3: Металл — двойная защита (ГФ-021 + ПФ-115 ЭКОНОМ) ---
if 'Грунтовка ГФ-021' in by_cat and 'Эмаль ПФ-115 ЭКОНОМ' in by_cat:
    gf = [p for p in by_cat['Грунтовка ГФ-021'] if 'красно-коричнев' in (p.get('color') or '').lower() or 'красно-коричнев' in p['name'].lower()]
    econom = by_cat['Эмаль ПФ-115 ЭКОНОМ']
    if gf and econom:
        chains.append({
            'id': 3,
            'name': 'Металл: двойная защита по выгодной цене',
            'surface': 'Металл (улица)',
            'purpose': 'Экономичное антикоррозийное покрытие — грунт + эмаль ЭКОНОМ',
            'layers': [
                {'step': 1, 'product': 'Грунтовка ГФ-021 Расцвет красно-коричневая 12 кг', 'role': 'Грунт', 'consumption': '60-80 г/м²', 'layers': 1, 'drying': '24 ч'},
                {'step': 2, 'product': 'Эмаль ПФ-115 ЭКОНОМ (цвет на выбор, 20 кг)', 'role': 'Покрытие', 'consumption': '120-160 г/м²', 'layers': 2, 'drying': '24 ч/слой'},
            ],
            'tco_note': 'Комбинация ГФ-021 + ЭКОНОМ = минимальная цена за надёжный результат. TCO на 20% ниже «только эмаль».',
            'segment': 'Эконом',
            'priority': 'Средняя',
        })

# --- CHAIN 4: Дерево улица — защита ---
if 'Эмаль ПФ-115 КМ' in by_cat:
    pf_km = by_cat['Эмаль ПФ-115 КМ']
    chains.append({
        'id': 4,
        'name': 'Дерево на улице: атмосферостойкая защита',
        'surface': 'Дерево (улица)',
        'purpose': 'Защита деревянных конструкций от атмосферных воздействий',
        'layers': [
            {'step': 1, 'product': 'Олифа (рекомендация, нет в ассортименте)', 'role': 'Пропитка', 'consumption': '80-100 г/м²', 'layers': 1, 'drying': '24 ч'},
            {'step': 2, 'product': 'Эмаль ПФ-115 КМ (цвет на выбор)', 'role': 'Покрытие', 'consumption': '100-140 г/м²', 'layers': 2, 'drying': '24 ч/слой'},
        ],
        'tco_note': 'Без пропитки: краска отслаивается через 1-2 сезона. С пропиткой: 4-6 лет. ⚠️ В ассортименте нет олифы — рекомендовать к завозу.',
        'segment': 'Универсал',
        'priority': 'Средняя',
    })

# --- CHAIN 5: Паркет — восстановление ---
if 'Лак паркетный' in by_cat:
    chains.append({
        'id': 5,
        'name': 'Паркет: восстановление и защита',
        'surface': 'Дерево (пол/паркет)',
        'purpose': 'Восстановление и защита паркетного покрытия',
        'layers': [
            {'step': 1, 'product': 'Шпатлёвка по дереву (рекомендация, нет в ассортименте)', 'role': 'Шпатлёвка', 'consumption': '200-300 г/м²', 'layers': 1, 'drying': '4-6 ч'},
            {'step': 2, 'product': 'Лак для паркета Ореол 1.776 кг', 'role': 'Покрытие', 'consumption': '80-120 г/м²', 'layers': 2-3, 'drying': '4-6 ч/слой'},
        ],
        'tco_note': 'Комплексный ремонт паркета вместо замены. Экономия 60-70% vs замена покрытия. ⚠️ Нет шпатлёвки — рекомендовать к завозу.',
        'segment': 'Премиум',
        'priority': 'Средняя',
    })

# --- CHAIN 6: Бетонный пол — защита ---
if 'Эмаль для пола' in by_cat:
    chains.append({
        'id': 6,
        'name': 'Бетонный пол: защита от истирания',
        'surface': 'Бетон (пол)',
        'purpose': 'Защита бетонных полов от пыления, истирания и химических воздействий',
        'layers': [
            {'step': 1, 'product': 'Грунтовка глубокого проникновения (рекомендация, нет в ассортименте)', 'role': 'Грунт', 'consumption': '100-150 г/м²', 'layers': 1, 'drying': '4-6 ч'},
            {'step': 2, 'product': 'Эмаль Расцвет для бетонных полов АК серая 20 кг', 'role': 'Покрытие', 'consumption': '120-180 г/м²', 'layers': 2, 'drying': '6-8 ч/слой'},
        ],
        'tco_note': 'Без грунта: отслаивание через 6-12 мес. С грунтом: 3-5 лет. ⚠️ Нет грунта для бетона — рекомендовать к завозу.',
        'segment': 'Универсал',
        'priority': 'Средняя',
    })

# --- CHAIN 7: Стены — масляная краска ---
if 'Краска масляная' in by_cat:
    chains.append({
        'id': 7,
        'name': 'Стены и окна: масляная покраска',
        'surface': 'Дерево/металл (окна, двери, стены)',
        'purpose': 'Традиционная масляная покраска окон, дверей, стен',
        'layers': [
            {'step': 1, 'product': 'Грунтовка ГФ-021 (для металла) или олифа (для дерева)', 'role': 'Грунт', 'consumption': '60-80 г/м²', 'layers': 1, 'drying': '24 ч'},
            {'step': 2, 'product': 'Краска масляная Расцвет (цвет на выбор)', 'role': 'Покрытие', 'consumption': '120-160 г/м²', 'layers': 2, 'drying': '24 ч/слой'},
        ],
        'tco_note': 'Масляная краска + грунт = 3-4 года службы vs 1-2 года без подготовки.',
        'segment': 'Эконом',
        'priority': 'Средняя',
    })

# --- CHAIN 8: Металл — нитроэмаль (быстросохнущая) ---
if 'Эмаль НЦ-132КП' in by_cat:
    nc = by_cat['Эмаль НЦ-132КП']
    colors = list(set(p['color'] for p in nc if p['color']))
    chains.append({
        'id': 8,
        'name': 'Металл: быстросохнущее покрытие (НЦ-132КП)',
        'surface': 'Металл (помещение)',
        'purpose': 'Быстросохнущее декоративное покрытие для внутренних работ',
        'layers': [
            {'step': 1, 'product': 'Грунтовка ГФ-021 Расцвет (красно-коричневая/серая)', 'role': 'Грунт', 'consumption': '60-80 г/м²', 'layers': 1, 'drying': '24 ч'},
            {'step': 2, 'product': f'Эмаль НЦ-132КП ({", ".join(colors)})', 'role': 'Покрытие', 'consumption': '80-120 г/м²', 'layers': 3-4, 'drying': '1-2 ч/слой'},
        ],
        'tco_note': 'НЦ-132КП сохнет в 10 раз быстрее ПФ-115. Идеально для срочных работ. Срок службы 3-4 года.',
        'segment': 'Универсал',
        'priority': 'Низкая',
    })

# --- CHAIN 9: ПФ-115 Стройкомплект — эконом для больших объёмов ---
if 'Эмаль ПФ-115 Стройкомплект' in by_cat:
    chains.append({
        'id': 9,
        'name': 'Металл: эконом-покраска больших объёмов',
        'surface': 'Металл (улица/помещение)',
        'purpose': 'Экономичное покрытие для больших объёмов металлоконструкций',
        'layers': [
            {'step': 1, 'product': 'Грунтовка ГФ-021 Расцвет красно-коричневая 12 кг', 'role': 'Грунт', 'consumption': '60-80 г/м²', 'layers': 1, 'drying': '24 ч'},
            {'step': 2, 'product': 'Эмаль ПФ-115 СТРОЙКОМПЛЕКТ синяя 20 кг', 'role': 'Покрытие', 'consumption': '120-160 г/м²', 'layers': 2, 'drying': '24 ч/слой'},
        ],
        'tco_note': 'СТРОЙКОМПЛЕКТ — самая доступная линейка. Грунт + эмаль = минимальный TCO.',
        'segment': 'Эконом',
        'priority': 'Низкая',
    })

# --- CHAIN 10: ПФ-115 матовая — декоративное покрытие ---
if 'Эмаль ПФ-115 матовая' in by_cat or any('матовая' in p['name'].lower() for p in products):
    chains.append({
        'id': 10,
        'name': 'Металл/дерево: матовое декоративное покрытие',
        'surface': 'Металл/дерево (помещение)',
        'purpose': 'Декоративное матовое покрытие для интерьерных работ',
        'layers': [
            {'step': 1, 'product': 'Грунтовка ГФ-021 / олифа (по поверхности)', 'role': 'Грунт', 'consumption': '60-80 г/м²', 'layers': 1, 'drying': '24 ч'},
            {'step': 2, 'product': 'Эмаль ПФ-115 белая матовая 20 кг', 'role': 'Покрытие', 'consumption': '100-140 г/м²', 'layers': 2, 'drying': '24 ч/слой'},
        ],
        'tco_note': 'Матовая эмаль — премиальный внешний вид без бликов. Срок службы 4-5 лет с грунтом.',
        'segment': 'Премиум',
        'priority': 'Низкая',
    })

# --- CHAIN 11: 3в1 + ПФ-115 — максимальная защита ---
if 'Грунт-эмаль 3в1' in by_cat and 'Эмаль ПФ-115 КМ' in by_cat:
    chains.append({
        'id': 11,
        'name': 'Металл: максимальная защита (3в1 + ПФ-115)',
        'surface': 'Металл (улица)',
        'purpose': 'Максимальная защита: грунт-эмаль как база + ПФ-115 как финишный слой',
        'layers': [
            {'step': 1, 'product': 'Грунт-эмаль ПРОСТОКРАШЕНО! 3в1 (красно-коричневая)', 'role': 'Грунт+база', 'consumption': '120-160 г/м²', 'layers': 1, 'drying': '2-4 ч'},
            {'step': 2, 'product': 'Эмаль ПФ-115 КМ (цвет на выбор)', 'role': 'Финишное покрытие', 'consumption': '100-140 г/м²', 'layers': 1-2, 'drying': '24 ч/слой'},
        ],
        'tco_note': '3в1 как грунт + ПФ-115 как финиш = 6-8 лет без перекраски. TCO на 50% ниже «только ПФ-115».',
        'segment': 'Универсал+',
        'priority': 'Высокая',
    })

# --- CHAIN 12: Расцвет ГОСТ — гарантированное качество ---
pf115_ras = [p for p in products if 'РАСЦВЕТ' in p['name'].upper() and 'ПФ-115' in p['name'].upper() and 'ГОСТ' in p['name'].upper()]
if pf115_ras:
    chains.append({
        'id': 12,
        'name': 'Металл: ГОСТ-качество по доступной цене',
        'surface': 'Металл (улица)',
        'purpose': 'Покраска по ГОСТ — гарантия качества для ответственных конструкций',
        'layers': [
            {'step': 1, 'product': 'Грунтовка ГФ-021 Расцвет красно-коричневая 12 кг', 'role': 'Грунт', 'consumption': '60-80 г/м²', 'layers': 1, 'drying': '24 ч'},
            {'step': 2, 'product': 'Эмаль ПФ-115 белая РАСЦВЕТ ГОСТ 20 кг', 'role': 'Покрытие', 'consumption': '100-140 г/м²', 'layers': 2, 'drying': '24 ч/слой'},
        ],
        'tco_note': 'ГОСТ = предсказуемый результат. Грунт + ГОСТ-эмаль = 4-5 лет службы.',
        'segment': 'Универсал',
        'priority': 'Средняя',
    })

print(f"\nTotal chains: {len(chains)}")
for c in chains:
    print(f"  #{c['id']}: {c['name']} [{c['surface']}] — {c['priority']} приоритет")

# ============================================================
# 5. CREATE EXCEL WORKBOOK
# ============================================================
wb = openpyxl.Workbook()

# Color palette (Empils brand)
NAVY = '0A2540'
CYAN = '00B4D8'
ORANGE = 'E86C00'
GREEN = '1B7D46'
RED = 'C0392B'
LIGHT_BLUE = 'D6EAF8'
LIGHT_ORANGE = 'FDEBD0'
LIGHT_GREEN = 'D5F5E3'
LIGHT_GRAY = 'F2F4F6'
WHITE = 'FFFFFF'

header_font = Font(name='Carlito', size=11, bold=True, color=WHITE)
header_fill_navy = PatternFill(start_color=NAVY, end_color=NAVY, fill_type='solid')
header_fill_orange = PatternFill(start_color=ORANGE, end_color=ORANGE, fill_type='solid')
header_fill_green = PatternFill(start_color=GREEN, end_color=GREEN, fill_type='solid')
body_font = Font(name='Carlito', size=10)
bold_font = Font(name='Carlito', size=10, bold=True)
title_font = Font(name='Carlito', size=14, bold=True, color=NAVY)
subtitle_font = Font(name='Carlito', size=11, bold=True, color=NAVY)
thin_border = Border(
    left=Side(style='thin', color='D0D8E0'),
    right=Side(style='thin', color='D0D8E0'),
    top=Side(style='thin', color='D0D8E0'),
    bottom=Side(style='thin', color='D0D8E0'),
)
center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)

def set_col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

def write_header_row(ws, row, headers, fill=header_fill_navy):
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col, value=h)
        cell.font = header_font
        cell.fill = fill
        cell.alignment = center_align
        cell.border = thin_border

def write_data_row(ws, row, data, bold_first=False):
    for col, val in enumerate(data, 1):
        cell = ws.cell(row=row, column=col, value=val)
        cell.font = bold_font if (col == 1 and bold_first) else body_font
        cell.alignment = left_align if col <= 2 else center_align
        cell.border = thin_border
        # Alternating rows
        if row % 2 == 0:
            cell.fill = PatternFill(start_color=LIGHT_GRAY, end_color=LIGHT_GRAY, fill_type='solid')

# ====== SHEET 1: Обзор цепочек ======
ws1 = wb.active
ws1.title = 'Обзор цепочек'
ws1.sheet_properties.tabColor = NAVY

# Title
ws1.merge_cells('A1:I1')
ws1.cell(row=1, column=1, value='ТЕХНОЛОГИЧЕСКИЕ ЦЕПОЧКИ ПЕТРОВИЧ — ОБЗОР').font = title_font
ws1.cell(row=1, column=1).alignment = Alignment(horizontal='left', vertical='center')
ws1.row_dimensions[1].height = 30

ws1.merge_cells('A2:I2')
ws1.cell(row=2, column=1, value=f'Ассортимент: {len(products)} SKU | Цепочек: {len(chains)} | Анализ: июнь 2026').font = Font(name='Carlito', size=10, color='506880')

# Headers
set_col_widths(ws1, [5, 40, 22, 40, 12, 14, 55, 18, 12])
headers = ['№', 'Название цепочки', 'Поверхность', 'Назначение', 'Шагов', 'Сегмент', 'Ключевой TCO-аргумент', 'Приоритет', 'Продукты']
write_header_row(ws1, 4, headers)

# Data
for i, c in enumerate(chains):
    r = 5 + i
    prods = ', '.join(l['product'][:50] for l in c['layers'])
    write_data_row(ws1, r, [
        c['id'], c['name'], c['surface'], c['purpose'],
        len(c['layers']), c['segment'], c['tco_note'], c['priority'], prods
    ], bold_first=True)
    # Color priority
    prio_cell = ws1.cell(row=r, column=8)
    if c['priority'] == 'Высокая':
        prio_cell.fill = PatternFill(start_color='E8F8F5', end_color='E8F8F5', fill_type='solid')
        prio_cell.font = Font(name='Carlito', size=10, bold=True, color=GREEN)
    elif c['priority'] == 'Средняя':
        prio_cell.fill = PatternFill(start_color=LIGHT_ORANGE, end_color=LIGHT_ORANGE, fill_type='solid')
        prio_cell.font = Font(name='Carlito', size=10, bold=True, color=ORANGE)
    else:
        prio_cell.font = Font(name='Carlito', size=10, color='808080')

# ====== SHEET 2: Детали цепочек ======
ws2 = wb.create_sheet('Детали цепочек')
ws2.sheet_properties.tabColor = ORANGE

ws2.merge_cells('A1:H1')
ws2.cell(row=1, column=1, value='ДЕТАЛИ ТЕХНОЛОГИЧЕСКИХ ЦЕПОЧЕК — ПОШАГОВЫЙ СОСТАВ').font = title_font
ws2.row_dimensions[1].height = 30

set_col_widths(ws2, [5, 35, 8, 50, 18, 22, 12, 18])
write_header_row(ws2, 3, ['Цепочка №', 'Название цепочки', 'Шаг', 'Продукт', 'Роль', 'Расход', 'Слоёв', 'Сушка'], fill=header_fill_orange)

r = 4
for c in chains:
    for j, layer in enumerate(c['layers']):
        write_data_row(ws2, r, [
            c['id'] if j == 0 else '',
            c['name'] if j == 0 else '',
            layer['step'],
            layer['product'],
            layer['role'],
            layer['consumption'],
            layer['layers'],
            layer['drying'],
        ], bold_first=True)
        # Role color
        role_cell = ws2.cell(row=r, column=5)
        if 'Грунт' in layer['role'] and '3в1' not in layer['role']:
            role_cell.fill = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type='solid')
        elif 'Покрытие' in layer['role'] or 'Финиш' in layer['role']:
            role_cell.fill = PatternFill(start_color=LIGHT_ORANGE, end_color=LIGHT_ORANGE, fill_type='solid')
        elif '3в1' in layer['role']:
            role_cell.fill = PatternFill(start_color=LIGHT_GREEN, end_color=LIGHT_GREEN, fill_type='solid')
        r += 1

# ====== SHEET 3: Матрица сегменты × продукты ======
ws3 = wb.create_sheet('Матрица сегменты')
ws3.sheet_properties.tabColor = GREEN

ws3.merge_cells('A1:H1')
ws3.cell(row=1, column=1, value='МАТРИЦА: СЕГМЕНТЫ × КАТЕГОРИИ ПРОДУКТОВ').font = title_font
ws3.row_dimensions[1].height = 30

# Get unique segments and categories
segments = sorted(set(c['segment'] for c in chains))
categories_order = ['Грунтовка ГФ-021', 'Грунт-эмаль 3в1', 'Эмаль ПФ-115 КМ', 'Эмаль ПФ-115 ЭКОНОМ',
                    'Эмаль ПФ-115 Расцвет', 'Эмаль ПФ-115 Стройкомплект', 'Эмаль ПФ-115 матовая',
                    'Эмаль НЦ-132КП', 'Краска масляная', 'Лак паркетный', 'Эмаль для пола']
available_cats = [c for c in categories_order if c in cat_count]

set_col_widths(ws3, [20] + [16] * len(available_cats))
write_header_row(ws3, 3, ['Сегмент'] + available_cats, fill=header_fill_green)

# Build matrix
for si, seg in enumerate(segments):
    r = 4 + si
    ws3.cell(row=r, column=1, value=seg).font = bold_font
    ws3.cell(row=r, column=1).border = thin_border
    
    for ci, cat in enumerate(available_cats):
        cell = ws3.cell(row=r, column=2 + ci)
        cell.border = thin_border
        cell.alignment = center_align
        # Check if this segment uses this category
        seg_chains = [c for c in chains if c['segment'] == seg]
        used = False
        for sc in seg_chains:
            for layer in sc['layers']:
                prod = layer['product'].upper()
                cat_upper = cat.upper()
                if cat_upper in prod or any(kw in prod for kw in cat_upper.split()):
                    used = True
                    break
            if used:
                break
        if used:
            cell.value = '●'
            cell.font = Font(name='Carlito', size=12, bold=True, color=GREEN)
            cell.fill = PatternFill(start_color='E8F8F5', end_color='E8F8F5', fill_type='solid')
        elif cat_count.get(cat, 0) > 0:
            cell.value = cat_count.get(cat, 0)
            cell.font = Font(name='Carlito', size=9, color='808080')
        else:
            cell.value = '—'
            cell.font = Font(name='Carlito', size=9, color='C0C0C0')

# ====== SHEET 4: Ассортиментная карта ======
ws4 = wb.create_sheet('Ассортиментная карта')
ws4.sheet_properties.tabColor = CYAN

ws4.merge_cells('A1:G1')
ws4.cell(row=1, column=1, value='АССОРТИМЕНТНАЯ КАРТА ПЕТРОВИЧ — ВСЕ ПРОДУКТЫ ЭМПИЛС').font = title_font
ws4.row_dimensions[1].height = 30

ws4.merge_cells('A2:G2')
ws4.cell(row=2, column=1, value=f'Всего SKU: {len(products)} | Категорий: {len(cat_count)}').font = Font(name='Carlito', size=10, color='506880')

set_col_widths(ws4, [5, 50, 20, 12, 12, 15, 50])
write_header_row(ws4, 4, ['№', 'Продукт', 'Категория', 'Вес, кг', 'Цвет', 'Код 1С', 'Цепочка'], fill=PatternFill(start_color=CYAN, end_color=CYAN, fill_type='solid'))
# Fix: white text on cyan might be hard, use navy
for col in range(1, 8):
    c = ws4.cell(row=4, column=col)
    c.font = Font(name='Carlito', size=11, bold=True, color=WHITE)

# Sort products by category
sorted_products = sorted(products, key=lambda p: (p['category'], p['name']))

for i, p in enumerate(sorted_products):
    r = 5 + i
    # Find which chains use this product
    used_in = []
    for c in chains:
        for layer in c['layers']:
            if p['name'][:30].upper() in layer['product'].upper() or any(
                kw.upper() in layer['product'].upper() 
                for kw in [p['name'].split()[0]] if len(p['name'].split()) > 0
            ):
                used_in.append(f"#{c['id']}")
                break
    chain_str = ', '.join(set(used_in)) if used_in else '—'
    
    write_data_row(ws4, r, [
        i + 1, p['name'], p['category'],
        p['weight'] if p['weight'] else '—',
        p['color'] if p['color'] else '—',
        p['code'] if p['code'] else '—',
        chain_str,
    ], bold_first=True)
    
    # Category color coding
    cat_cell = ws4.cell(row=r, column=3)
    if 'Грунт' in p['category'] and '3в1' not in p['category']:
        cat_cell.fill = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type='solid')
    elif '3в1' in p['category']:
        cat_cell.fill = PatternFill(start_color=LIGHT_GREEN, end_color=LIGHT_GREEN, fill_type='solid')
    elif 'ПФ-115' in p['category']:
        cat_cell.fill = PatternFill(start_color=LIGHT_ORANGE, end_color=LIGHT_ORANGE, fill_type='solid')
    
    # Chain usage highlight
    chain_cell = ws4.cell(row=r, column=7)
    if chain_str != '—':
        chain_cell.font = Font(name='Carlito', size=10, bold=True, color=GREEN)

# ====== SHEET 5: Рекомендации по ассортименту ======
ws5 = wb.create_sheet('Рекомендации')
ws5.sheet_properties.tabColor = RED

ws5.merge_cells('A1:D1')
ws5.cell(row=1, column=1, value='РЕКОМЕНДАЦИИ ПО ДОПОЛНЕНИЮ АССОРТИМЕНТА').font = title_font
ws5.row_dimensions[1].height = 30

ws5.merge_cells('A2:D2')
ws5.cell(row=2, column=1, value='Продукты, которых не хватает для полных технологических цепочек').font = Font(name='Carlito', size=10, color=RED)

set_col_widths(ws5, [5, 35, 55, 40])
write_header_row(ws5, 4, ['№', 'Продукт', 'Обоснование', 'Цепочки, где нужен'], fill=PatternFill(start_color=RED, end_color=RED, fill_type='solid'))

recommendations = [
    ('Олифа натуральная/оксоль', 'Необходима для подготовки дерева перед покраской. Без неё краска отслаивается через 1-2 сезона. Ключевой продукт для цепочки «Дерево на улице».', 'Цепочка #4 (Дерево улица), #7 (Стены/окна)'),
    ('Шпатлёвка по дереву', 'Необходима для ремонта паркета перед лакировкой. Без шпатлёвки невозможна качественная реставрация. Открывает сегмент премиум-ремонта полов.', 'Цепочка #5 (Паркет)'),
    ('Грунтовка глубокого проникновения для бетона', 'Критически важна для адгезии эмали по бетону. Без грунта эмаль для полов отслаивается через 6-12 месяцев.', 'Цепочка #6 (Бетонный пол)'),
    ('Краска ВД-АК (водно-дисперсионная)', 'Отсутствует весь сегмент ВД-красок для стен. Это самый быстрорастущий сегмент рынка. Потеря 30-40% потенциальных покупателей стеновых красок.', 'Новая цепочка: Стены внутри помещений'),
    ('Бетоноконтакт', 'Грунтовка для подготовки гладких поверхностей (плитка, старая краска) перед нанесением шпатлёвки/краски. Важен для ремонтного сегмента.', 'Новая цепочка: Ремонт стен'),
    ('Шпатлёвка гипсовая/полимерная', 'Отсутствует полностью. Без шпатлёвки невозможна подготовка стен — самый частый сценарий ремонта.', 'Новая цепочка: Стены внутри помещений'),
    ('Разбавитель/уайт-спирит', 'Сопутствующий товар с высокой маржой. Необходим для разведения ПФ-115 и мытья инструментов. Отсутствует как позиция.', 'Все цепочки с ПФ-115 и масляной краской'),
    ('Грунтовка ГФ-021 малые фасовки (1.8-2.7 кг)', 'Есть только 12 кг — слишком много для розничного покупателя. Малые фасовки увеличат конверсию в технологическую цепочку.', 'Цепочки #1, #3, #9, #12'),
]

for i, (prod, reason, chains_where) in enumerate(recommendations):
    r = 5 + i
    write_data_row(ws5, r, [i+1, prod, reason, chains_where], bold_first=True)
    ws5.cell(row=r, column=2).font = Font(name='Carlito', size=10, bold=True, color=RED)

# ====== SAVE ======
output = '/home/z/my-project/download/Технологические_цепочки_Петрович.xlsx'
wb.save(output)
print(f"\nFile saved: {output}")
