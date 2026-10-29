#!/usr/bin/env python3
"""
Анализ ассортимента Лемана ПРО + построение технологических цепочек Эмпилс.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re

# === СТИЛИ ===
NAVY = '0A2540'; CYAN = '00B4D8'; ORANGE = 'E86C00'; GREEN = '1B7D46'; RED = 'C0392B'
WHITE = 'FFFFFF'; LG = 'F2F2F2'; BG = 'D9D9D9'; YELLOW = 'FFF2CC'

ft = lambda sz, b=False, c=NAVY: Font(name='Carlito', size=sz, bold=b, color=c)
pf = lambda c: PatternFill(start_color=c, end_color=c, fill_type='solid')
ac = Alignment(horizontal='center', vertical='center', wrap_text=True)
al = Alignment(horizontal='left', vertical='center', wrap_text=True)
bd = Border(left=Side('thin', BG), right=Side('thin', BG), top=Side('thin', BG), bottom=Side('thin', BG))

def sc(ws, r, c, v, font=ft(10), fill=None, align=ac, border=bd, fmt=None):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = font; cell.alignment = align; cell.border = border
    if fill: cell.fill = fill
    if fmt: cell.number_format = fmt
    return cell

# === ЧТЕНИЕ ДАННЫХ ===
src = '/home/z/my-project/upload/ЛеманаПро.xlsx'
wb_src = openpyxl.load_workbook(src)
ws_src = wb_src['Лемана ПРО']

# Собираем все уникальные продукты
products = {}  # код → данные
for row in range(3, ws_src.max_row + 1):
    for name_col, code_col, sales_cols in [
        (2, 3, [4,5,6,7,8]),   # Блок A-H
        (18, 19, [20,21,22,23,24])  # Блок Q-X
    ]:
        name = ws_src.cell(row=row, column=name_col).value
        code = ws_src.cell(row=row, column=code_col).value
        if name and code:
            code_str = str(code).strip()
            if code_str not in products:
                # Считаем продажи
                sales = []
                for sc_col in sales_cols:
                    v = ws_src.cell(row=row, column=sc_col).value
                    try: sales.append(float(v) if v else 0)
                    except: sales.append(0)
                products[code_str] = {
                    'name': str(name).strip(),
                    'code': code_str,
                    'sales': sales,
                    'total': sum(sales),
                }

print(f"Уникальных SKU: {len(products)}")

# === КЛАССИФИКАЦИЯ ===
def classify_product(name):
    """Классифицировать продукт по названию"""
    n = name.lower()
    
    # Бренд
    brand = 'Другой'
    if 'avantgarde' in n: brand = 'AVANTGARDE'
    elif 'newtone' in n or 'modern pro' in n: brand = 'NEWTONE'
    elif 'ореол' in n or 'oreol' in n: brand = 'ОРЕОЛ'
    elif 'простокраска' in n: brand = 'ПростоКРАСКА'
    elif 'расцвет' in n: brand = 'Расцвет'
    elif 'эконом' in n: brand = 'ЭКОНОМ'
    elif 'лучшая цена' in n: brand = 'Лучшая цена'
    elif 'леруа' in n or 'pl леруа' in n: brand = 'PL ЛЕРУА'
    elif 'monlid' in n: brand = 'MONLID'
    elif 'luxens' in n: brand = 'Luxens'
    elif 'empils paint lab' in n or 'titan' in n: brand = 'Paint Lab / TITAN'
    elif 'дисконт' in n: brand = 'ДИСКОНТ'
    
    # Категория
    category = 'Прочее'
    subcategory = ''
    
    if 'грунт-эмаль' in n and 'ржавч' in n:
        category = 'Грунт-эмаль по ржавчине'
        if '3в1' in n: subcategory = '3в1'
        elif 'titan' in n: subcategory = 'TITAN'
        elif 'paint lab' in n: subcategory = 'Paint Lab'
    elif 'грунтовк' in n or ('грунт' in n and 'эмаль' not in n):
        category = 'Грунтовка'
        if 'гф-021' in n: subcategory = 'ГФ-021'
        elif 'вд' in n or 'универсальн' in n: subcategory = 'ВД универсальная'
        elif 'бетоноконтакт' in n or 'бетон-контакт' in n: subcategory = 'Бетоноконтакт'
        elif 'глубокого проникновен' in n: subcategory = 'Глубокого проникновения'
        elif 'влагоизолир' in n or 'влаг' in n: subcategory = 'Влагоизолирующий'
        elif 'антисепт' in n: subcategory = 'С антисептиком'
    elif 'шпатлёвк' in n or 'шпатлевк' in n:
        category = 'Шпатлёвка'
        if 'финиш' in n: subcategory = 'Финишная'
        elif 'стартов' in n: subcategory = 'Стартовая'
        elif 'универсальн' in n: subcategory = 'Универсальная'
        elif 'влажн' in n: subcategory = 'Для влажных помещений'
    elif 'краска вд' in n or ('краска' in n and 'вд' in n):
        category = 'Краска ВД'
        if 'потолк' in n: subcategory = 'Потолочная'
        elif 'фасад' in n: subcategory = 'Фасадная'
        elif 'интерьер' in n or 'моющ' in n: subcategory = 'Интерьерная моющаяся'
        elif 'стен' in n: subcategory = 'Для стен и потолков'
        elif 'фактурн' in n: subcategory = 'Фактурная'
    elif 'краска фасадн' in n or ('фасад' in n and 'краск' in n):
        category = 'Краска ВД'
        subcategory = 'Фасадная'
    elif 'эмаль пф-115' in n or 'эмаль пф-115' in n:
        category = 'Эмаль ПФ-115'
    elif 'эмаль для пол' in n:
        category = 'Эмаль для пола'
    elif 'эмаль' in n and 'радиатор' in n:
        category = 'Эмаль для радиаторов'
    elif 'эмаль' in n and ('алкидн' in n or 'универсальн' in n):
        category = 'Эмаль акриловая универсальная'
    elif 'лак' in n and 'яхтн' in n:
        category = 'Лак яхтный'
    elif 'лак' in n and 'паркетн' in n:
        category = 'Лак паркетный'
    elif 'лак' in n:
        category = 'Лак'
    elif 'эмаль' in n:
        category = 'Эмаль'
    elif 'краска' in n and 'osb' in n:
        category = 'Краска ВД'
        subcategory = 'Для OSB'
    elif 'декоративн' in n or 'штукатурк' in n:
        category = 'Декоративная штукатурка'
    
    # Назначение (для какой поверхности/задачи)
    purpose = ''
    if 'ржавчин' in n: purpose = 'Металл / ржавчина'
    elif 'потолк' in n: purpose = 'Потолки'
    elif 'фасад' in n: purpose = 'Фасад'
    elif 'пол' in n or 'паркет' in n: purpose = 'Полы / паркет'
    elif 'радиатор' in n or 'отоплен' in n: purpose = 'Радиаторы'
    elif 'окн' in n: purpose = 'Окна'
    elif 'влажн' in n or 'влаг' in n: purpose = 'Влажные зоны'
    elif 'дерев' in n or 'яхтн' in n: purpose = 'Дерево / яхтный'
    elif 'osb' in n or 'sip' in n: purpose = 'OSB / SIP'
    elif 'интерьер' in n or 'стен' in n: purpose = 'Интерьер'
    elif 'бетон' in n or 'контакт' in n: purpose = 'Проблемные основания'
    
    # Фасовка
    weight = ''
    m = re.search(r'(\d+[,.]?\d*)\s*(кг|л|мл)', n)
    if m: weight = m.group(0).replace(',', '.')
    
    return {
        'brand': brand,
        'category': category,
        'subcategory': subcategory,
        'purpose': purpose,
        'weight': weight,
    }

# Классифицируем все продукты
for code, p in products.items():
    cls = classify_product(p['name'])
    p.update(cls)

# Статистика по категориям
cats = {}
for code, p in products.items():
    cat = p['category']
    if cat not in cats:
        cats[cat] = {'count': 0, 'sales': 0, 'brands': set()}
    cats[cat]['count'] += 1
    cats[cat]['sales'] += p['total']
    cats[cat]['brands'].add(p['brand'])

print("\nКатегории:")
for cat, data in sorted(cats.items(), key=lambda x: -x[1]['sales']):
    print(f"  {cat}: {data['count']} SKU, продажи={data['sales']:.0f}, бренды={', '.join(data['brands'])}")

# === ТЕХНОЛОГИЧЕСКИЕ ЦЕПОЧКИ ===
# Строим на основе логики из PPTX + того, что реально есть в ассортименте Лемана ПРО

technologies = [
    {
        'num': 1,
        'name': 'Фасад, который самоочищается',
        'problem': 'Фасад пачкается, подтёки, грибок, трещины от влажности',
        'chain': [
            ('Грунтовка', 'Грунтовка глубокого проникновения с антисептиком', 'NEWTONE', 'Укрепляет, убивает плесень'),
            ('Финишное покрытие', 'Краска фасадная силиконовая MODERN PRO', 'NEWTONE', 'Эффект лотоса, W3, 12 лет'),
        ],
        'benefit': 'Силиконовая краска + антисептический грунт = самоочищение и 12 лет без ремонта',
        'priority': 'ВЫСОКИЙ',
        'in_assortment': 'Грунтовка ВД + Краска ВД фасадная — ЧАСТИЧНО (нет MODERN PRO)',
    },
    {
        'num': 2,
        'name': 'Крепость во влажной зоне',
        'problem': 'Конденсат, плесень, краска размокает при мытье',
        'chain': [
            ('Грунтовка', 'Грунт-концентрат влагоизолирующий', 'ОРЕОЛ', 'Гидроизолирующий барьер'),
            ('Шпатлёвка', 'Шпатлёвка для влажных помещений с антисептиком', 'ОРЕОЛ', 'Не разрушается от пара'),
            ('Финишное покрытие', 'Краска для влажных помещений суперсстойкая / Lotus 20', 'ОРЕОЛ / NEWTONE', '1-й класс по DIN'),
        ],
        'benefit': 'Сэндвич-система: влагоизолирующий грунт + шпатлёвка с антисептиком + суперсстойкая краска',
        'priority': 'ВЫСОКИЙ',
        'in_assortment': 'Грунтовка ВД + Краска ВД моющаяся — ЧАСТИЧНО (нет спец. шпатлёвки)',
    },
    {
        'num': 3,
        'name': 'Идеальный интерьер под шелкографию',
        'problem': 'Дефекты проступают, краска ложится полосами',
        'chain': [
            ('Грунтовка', 'Грунтовка глубокого проникновения с антисептиком', 'NEWTONE', 'Выравнивает впитывание'),
            ('Шпатлёвка', 'Шпатлёвка суперфинишная акриловая', 'NEWTONE', 'Ювелирная гладкость от 0 мм'),
            ('Финишное покрытие', 'Краска интерьерная моющаяся AVANTGARDE', 'AVANTGARDE', 'Гибрид, 10x прочнее, без полос'),
        ],
        'benefit': 'Суперфиниш + гибридная краска = эффект немецкого качества без маляров-премиум',
        'priority': 'ВЫСОКИЙ',
        'in_assortment': 'Лак паркетный AVANTGARDE есть, краска ВД AVANTGARDE — НЕТ',
    },
    {
        'num': 4,
        'name': 'Деревянный пол без циклёвки',
        'problem': 'Пол истирается, царапины, лак желтеет',
        'chain': [
            ('Подготовка', 'Грунтовка глубокого проникновения', 'NEWTONE', 'Блокировка смол, укрепление'),
            ('Финишное покрытие', 'Лак высокой твердости паркетный AVANTGARDE', 'AVANTGARDE', '1-й класс стойкости, не желтеет'),
        ],
        'benefit': 'Гибридный лак: экологичность воды + прочность алкида, 3 часа межслойная сушка',
        'priority': 'ВЫСОКИЙ',
        'in_assortment': 'Лак паркетный AVANTGARDE — ЕСТЬ (6 SKU)',
    },
    {
        'num': 5,
        'name': 'Окна, которые не слипаются',
        'problem': 'Створки прилипают, трещины, желтеет',
        'chain': [
            ('Грунт (металл)', 'Эмаль акриловая универсальная (как грунт)', 'ОРЕОЛ', 'Ингибитор коррозии'),
            ('Финишное покрытие', 'Эмаль для окон полиуретановая акриловая', 'ОРЕОЛ', 'Anti-blocking, эластичность'),
        ],
        'benefit': 'Термошов + anti-blocking = окна не слипаются и не желтеют',
        'priority': 'СРЕДНИЙ',
        'in_assortment': 'Эмаль для радиаторов есть, эмаль для окон — НЕТ',
    },
    {
        'num': 6,
        'name': 'Морской климат и экстерьер',
        'problem': 'УФ, дожди, мороз — дерево серееет, трескается',
        'chain': [
            ('Грунтовка', 'Грунтовка глубокого проникновения с антисептиком', 'NEWTONE', 'Подготовка дерева'),
            ('Финишное покрытие', 'Лак высокой твердости яхтный AVANTGARDE', 'AVANTGARDE', 'Не желтеет, морской климат'),
        ],
        'benefit': 'Гибридный яхтный лак: не желтеет + не пахнет + стоек к воде и маслам',
        'priority': 'ВЫСОКИЙ',
        'in_assortment': 'Лак яхтный MONLID/Luxens — ЕСТЬ, AVANTGARDE яхтный — НЕТ',
    },
    {
        'num': 7,
        'name': 'Покраска проблемных оснований',
        'problem': 'Старая краска, плитка — ВД краска отваливается',
        'chain': [
            ('Грунтовка', 'Бетоноконтакт', 'ОРЕОЛ', 'Шероховатость + розовый индикатор'),
            ('Финишное покрытие', 'Краска для влажных помещений / Эмаль акриловая', 'ОРЕОЛ', 'Цепляется за шероховатость'),
        ],
        'benefit': 'Не нужно удалять старую плитку/краску — огромная экономия',
        'priority': 'СРЕДНИЙ',
        'in_assortment': 'ГФ-021 есть, Бетоноконтакт — НЕТ',
    },
    {
        'num': 8,
        'name': 'Быстрый ремонт радиаторов без запаха',
        'problem': 'Ржавые батареи, алкидные воняют и желтеют',
        'chain': [
            ('Защита металла', 'Эмаль акриловая универсальная (как грунт)', 'ОРЕОЛ', 'Блокирует ржавчину'),
            ('Финишное покрытие', 'Эмаль для радиаторов акриловая', 'ОРЕОЛ', 'До 120°C, не желтеет'),
        ],
        'benefit': 'Без запаха, красить зимой, не выселяя семью',
        'priority': 'СРЕДНИЙ',
        'in_assortment': 'Эмаль для радиаторов — ЕСТЬ (1 SKU)',
    },
    {
        'num': 9,
        'name': 'Система для OSB и SIP-домов',
        'problem': 'OSB впитывает пятнами, краска трескается',
        'chain': [
            ('Грунтовка', 'Грунт-концентрат влагоизолирующий', 'ОРЕОЛ', 'Запечатывает поры OSB'),
            ('Финишное покрытие', 'Краска для OSB база А', 'ОРЕОЛ', 'Эластичная, с фунгицидом'),
        ],
        'benefit': 'Ровный цвет + защита без микротрещин',
        'priority': 'СРЕДНИЙ',
        'in_assortment': 'Грунтовка ВД есть, Краска для OSB — НЕТ',
    },
    {
        'num': 10,
        'name': 'Эконом-ремонт съёмной квартиры/дачи',
        'problem': 'Быстро и дёшево освежить ремонт',
        'chain': [
            ('Грунтовка', 'Грунтовка универсальная с антисептиком', 'ДИСКОНТ', 'Базовая подготовка'),
            ('Шпатлёвка', 'Шпатлёвка масляно-клеевая', 'ДИСКОНТ', 'Быстрое выравнивание'),
            ('Финишное покрытие', 'Краска интерьерная моющаяся', 'ДИСКОНТ', 'Можно мыть'),
        ],
        'benefit': 'Минимальный бюджет, но покрытие не отваливается',
        'priority': 'НИЗКИЙ',
        'in_assortment': 'Грунтовка ВД ДИСКОНТ + Краска ВД ОРЕОЛ — ЧАСТИЧНО',
    },
    # Новые технологии на основе ассортимента Лемана ПРО
    {
        'num': 11,
        'name': 'Защита металла 3в1 — быстро и навсегда',
        'problem': 'Ржавый металл, нужен быстрый результат без грунтования',
        'chain': [
            ('Финишное покрытие', 'Грунт-эмаль по ржавчине 3в1', 'ОРЕОЛ / Paint Lab / TITAN', 'Грунт+эмаль+защита в одном'),
        ],
        'benefit': '1 продукт вместо 3 — нет грунта, нет преобразователя. Самый продаваемый формат в DIY',
        'priority': 'ВЫСОКИЙ',
        'in_assortment': 'ЕСТЬ — 34 SKU (3в1 + Paint Lab + TITAN)',
    },
    {
        'num': 12,
        'name': 'Эмаль ПФ-115 — универсальная защита',
        'problem': 'Нужна краска по металлу и дереву для наружных/внутренних работ',
        'chain': [
            ('Подготовка', 'Грунтовка ГФ-021', 'ОРЕОЛ', 'Антикоррозионная подготовка'),
            ('Финишное покрытие', 'Эмаль ПФ-115', 'ПростоКРАСКА / ОРЕОЛ Premium / ЭКОНОМ', 'Защита + цвет'),
        ],
        'benefit': 'Классическая советская связка: ГФ-021 + ПФ-115 = гарантированная защита металла',
        'priority': 'ВЫСОКИЙ',
        'in_assortment': 'ЕСТЬ — ГФ-021 (9 SKU) + ПФ-115 (60+ SKU)',
    },
    {
        'num': 13,
        'name': 'Полы, которые не стираются',
        'problem': 'Бетонный пол пылит, стирается, впитывает масло',
        'chain': [
            ('Подготовка', 'Грунтовка ГФ-021', 'ОРЕОЛ', 'Упрочнение поверхности'),
            ('Финишное покрытие', 'Эмаль Расцвет для бетонных полов', 'Расцвет', 'Износостойкая, маслобензостойкая'),
        ],
        'benefit': 'ГФ-021 блокирует пыль + Расцвет даёт износостойкое покрытие',
        'priority': 'СРЕДНИЙ',
        'in_assortment': 'ЕСТЬ — ГФ-021 + Расцвет (10 SKU)',
    },
    {
        'num': 14,
        'name': 'Яхтный лак — защита дерева на улице',
        'problem': 'Деревянные конструкции на улице гниют и трескаются',
        'chain': [
            ('Подготовка', 'Грунтовка глубокого проникновения', 'NEWTONE', 'Антисептик + укрепление'),
            ('Финишное покрытие', 'Лак яхтный атмосферостойкий', 'MONLID / Luxens', 'Атмосферостойкость, алкидно-уретан'),
        ],
        'benefit': 'Алкидно-уретановый лак — максимальная защита от дождя, УФ, мороза',
        'priority': 'ВЫСОКИЙ',
        'in_assortment': 'ЕСТЬ — Лак яхтный MONLID/Luxens (11 SKU)',
    },
    {
        'num': 15,
        'name': 'Потолок без разводов',
        'problem': 'Потолочная краска ложится полосами, видны разводы',
        'chain': [
            ('Грунтовка', 'Грунтовка ВД универсальная', 'ОРЕОЛ ДИСКОНТ', 'Выравнивание впитывания'),
            ('Финишное покрытие', 'Краска ВД для потолков PL ЛЕРУА', 'PL ЛЕРУА', 'Без разводов, глубокоматовая'),
        ],
        'benefit': 'Грунтовка предотвращает пятна, потолочная краска — без разводов',
        'priority': 'ВЫСОКИЙ',
        'in_assortment': 'ЕСТЬ — Грунтовка ВД + Краска потолочная (5 SKU)',
    },
]

# === СОЗДАНИЕ EXCEL ===
dst = '/home/z/my-project/download/Анализ_ЛеманаПРО_технологии.xlsx'
wb = openpyxl.Workbook()

# ---- ЛИСТ 1: АССОРТИМЕНТ ----
ws1 = wb.active
ws1.title = 'Ассортимент'
ws1.sheet_properties.tabColor = CYAN

# Ширина
widths = [3, 10, 40, 10, 15, 18, 18, 15, 15, 15, 15]
for i, w in enumerate(widths, 1):
    ws1.column_dimensions[get_column_letter(i)].width = w

# Заголовок
ws1.merge_cells('B2:K2')
sc(ws1, 2, 2, 'АССОРТИМЕНТ ЛЕМАНА ПРО — АНАЛИЗ И КЛАССИФИКАЦИЯ', ft(16, True, WHITE), pf(NAVY), ac)

# Заголовки
hdrs = ['№', 'Номенклатура', 'Код 1С', 'Бренд', 'Категория', 'Подкатегория', 'Назначение', 'Фасовка', 'Продажи (тыс. ₽)', 'Доля']
for i, h in enumerate(hdrs):
    sc(ws1, 4, i+1, h, ft(10, True, WHITE), pf(CYAN))

# Сортируем по продажам
sorted_prods = sorted(products.values(), key=lambda x: -x['total'])
total_sales = sum(p['total'] for p in sorted_prods)

for idx, p in enumerate(sorted_prods):
    row = 5 + idx
    sc(ws1, row, 1, idx+1, ft(9))
    sc(ws1, row, 2, p['name'], ft(9), None, al)
    sc(ws1, row, 3, p['code'], ft(9))
    sc(ws1, row, 4, p['brand'], ft(9, True, CYAN))
    sc(ws1, row, 5, p['category'], ft(9))
    sc(ws1, row, 6, p['subcategory'], ft(9))
    sc(ws1, row, 7, p['purpose'], ft(9))
    sc(ws1, row, 8, p['weight'], ft(9))
    sc(ws1, row, 9, round(p['total']/1000, 1), ft(9), None, ac, bd, '#,##0.0')
    pct = p['total']/total_sales if total_sales else 0
    sc(ws1, row, 10, pct, ft(9), None, ac, bd, '0.0%')
    # Подсветка по бренду
    brand_colors = {
        'AVANTGARDE': 'C00000', 'NEWTONE': '4472C4', 'ОРЕОЛ': '00B050',
        'ПростоКРАСКА': 'FFC000', 'Расцвет': '7030A0', 'ЭКОНОМ': '808080',
        'MONLID': '002060', 'Luxens': '00B0F0', 'Paint Lab / TITAN': 'ED7D31',
        'PL ЛЕРУА': 'BF8F00', 'ДИСКОНТ': 'A6A6A6',
    }
    bc = brand_colors.get(p['brand'], 'FFFFFF')
    ws1.cell(row=row, column=4).fill = PatternFill(start_color=bc, end_color=bc, fill_type='solid')
    ws1.cell(row=row, column=4).font = ft(9, True, WHITE if bc not in ['FFC000', '00B0F0', '00B050'] else NAVY)

# ---- ЛИСТ 2: СВОДКА ПО КАТЕГОРИЯМ ----
ws2 = wb.create_sheet('Сводка по категориям')
ws2.sheet_properties.tabColor = GREEN

ws2.column_dimensions['A'].width = 3
ws2.column_dimensions['B'].width = 30
ws2.column_dimensions['C'].width = 10
ws2.column_dimensions['D'].width = 18
ws2.column_dimensions['E'].width = 12
ws2.column_dimensions['F'].width = 40

ws2.merge_cells('B2:F2')
sc(ws2, 2, 2, 'СВОДКА ПО КАТЕГОРИЯМ И БРЕНДАМ', ft(16, True, WHITE), pf(NAVY), ac)

for i, h in enumerate(['Категория', 'SKU', 'Продажи (тыс. ₽)', 'Доля', 'Бренды в категории']):
    sc(ws2, 4, i+2, h, ft(10, True, WHITE), pf(CYAN))

sorted_cats = sorted(cats.items(), key=lambda x: -x[1]['sales'])
for idx, (cat, data) in enumerate(sorted_cats):
    row = 5 + idx
    sc(ws2, row, 2, cat, ft(10, True))
    sc(ws2, row, 3, data['count'], ft(10))
    sc(ws2, row, 4, round(data['sales']/1000, 1), ft(10), None, ac, bd, '#,##0.0')
    sc(ws2, row, 5, data['sales']/total_sales if total_sales else 0, ft(10), None, ac, bd, '0.0%')
    sc(ws2, row, 6, ', '.join(sorted(data['brands'])), ft(9), None, al)

# ---- ЛИСТ 3: ТЕХНОЛОГИЧЕСКИЕ ЦЕПОЧКИ ----
ws3 = wb.create_sheet('Технологические цепочки')
ws3.sheet_properties.tabColor = ORANGE

ws3.column_dimensions['A'].width = 3
ws3.column_dimensions['B'].width = 5
ws3.column_dimensions['C'].width = 35
ws3.column_dimensions['D'].width = 15
ws3.column_dimensions['E'].width = 45
ws3.column_dimensions['F'].width = 15
ws3.column_dimensions['G'].width = 35
ws3.column_dimensions['H'].width = 14
ws3.column_dimensions['I'].width = 40

ws3.merge_cells('B2:I2')
sc(ws3, 2, 2, 'ТЕХНОЛОГИЧЕСКИЕ ЦЕПОЧКИ ЭМПИЛС ДЛЯ ЛЕМАНА ПРО', ft(16, True, WHITE), pf(NAVY), ac)

ws3.merge_cells('B3:I3')
sc(ws3, 3, 2, 'Группирование по логике: Проблема → Система продуктов (Грунт→Шпатлёвка→Покрытие) → Выгода', ft(11, False, CYAN), pf(NAVY), ac)

# Заголовки
hdrs3 = ['№', 'Технология', 'Приоритет', 'Проблема', 'Слой', 'Продукт', 'Бренд', 'Роль в цепочке']
for i, h in enumerate(hdrs3):
    sc(ws3, 5, i+2, h, ft(10, True, WHITE), pf(NAVY))

row = 6
for tech in technologies:
    # Заголовок технологии
    priority_color = {'ВЫСОКИЙ': 'C00000', 'СРЕДНИЙ': 'ED7D31', 'НИЗКИЙ': '808080'}
    pc = priority_color.get(tech['priority'], '808080')
    
    sc(ws3, row, 2, tech['num'], ft(12, True, WHITE), PatternFill(start_color=pc, end_color=pc, fill_type='solid'))
    sc(ws3, row, 3, tech['name'], ft(11, True, NAVY), None, al)
    sc(ws3, row, 4, tech['priority'], ft(10, True, WHITE), PatternFill(start_color=pc, end_color=pc, fill_type='solid'))
    sc(ws3, row, 5, tech['problem'], ft(9), None, al)
    sc(ws3, row, 8, tech['in_assortment'], ft(9, False, RED if 'НЕТ' in tech.get('in_assortment','') else GREEN), None, al)
    sc(ws3, row, 9, tech['benefit'], ft(9, True), None, al)
    row += 1
    
    # Продукты в цепочке
    for layer, product, brand, role in tech['chain']:
        sc(ws3, row, 6, layer, ft(9, True, CYAN), None, al)
        sc(ws3, row, 7, product, ft(9), None, al)
        sc(ws3, row, 8, brand, ft(9, True), None, al)
        sc(ws3, row, 9, role, ft(9), None, al)
        row += 1
    
    row += 1  # Пустая строка между технологиями

# ---- ЛИСТ 4: МАТРИЦА ПОКРЫТИЯ ----
ws4 = wb.create_sheet('Матрица покрытия')
ws4.sheet_properties.tabColor = RED

ws4.column_dimensions['A'].width = 3
ws4.column_dimensions['B'].width = 5
ws4.column_dimensions['C'].width = 30
ws4.column_dimensions['D'].width = 12
ws4.column_dimensions['E'].width = 12
ws4.column_dimensions['F'].width = 12
ws4.column_dimensions['G'].width = 12
ws4.column_dimensions['H'].width = 12
ws4.column_dimensions['I'].width = 12
ws4.column_dimensions['J'].width = 12
ws4.column_dimensions['K'].width = 14

ws4.merge_cells('B2:K2')
sc(ws4, 2, 2, 'МАТРИЦА ПОКРЫТИЯ: ЧТО ЕСТЬ vs ЧТО НУЖНО', ft(16, True, WHITE), pf(NAVY), ac)

ws4.merge_cells('B3:K3')
sc(ws4, 3, 2, 'Зелёный = есть в ассортименте, Красный = отсутствует (добавить!), Жёлтый = частично', ft(11, False, CYAN), pf(NAVY), ac)

# Заголовки
matrix_hdrs = ['№', 'Технология', 'Грунтовка', 'Шпатлёвка старт.', 'Шпатлёвка финиш.', 'Краска ВД', 'Лак/эмаль', 'Грунт-эмаль', 'Декор', 'Оценка']
for i, h in enumerate(matrix_hdrs):
    sc(ws4, 5, i+2, h, ft(10, True, WHITE), pf(CYAN))

# Матрица: для каждой технологии показываем наличие
matrix_data = [
    (1, 'Фасад самоочищающийся', '🟡', '—', '—', '🟡', '—', '—', '—', 'ЧАСТИЧНО'),
    (2, 'Влажная зона', '🟡', '🔴', '🔴', '🟡', '—', '—', '—', 'СЛАБО'),
    (3, 'Идеальный интерьер', '🟡', '🔴', '🔴', '🔴', '—', '—', '—', 'НЕТ AVANTGARDE'),
    (4, 'Деревянный пол', '🟡', '—', '—', '—', '🟢', '—', '—', 'ЕСТЬ'),
    (5, 'Окна не слипаются', '🔴', '—', '—', '—', '🔴', '—', '—', 'НЕТ'),
    (6, 'Морской климат', '🟡', '—', '—', '—', '🟡', '—', '—', 'ЧАСТИЧНО'),
    (7, 'Проблемные основания', '🔴', '—', '—', '🟡', '—', '—', '—', 'НЕТ БЕТОНОКОНТАКТ'),
    (8, 'Радиаторы без запаха', '🟡', '—', '—', '—', '🟢', '—', '—', 'ЕСТЬ'),
    (9, 'OSB / SIP-дома', '🟡', '—', '—', '🔴', '—', '—', '—', 'НЕТ КРАСКИ OSB'),
    (10, 'Эконом-ремонт', '🟢', '🔴', '—', '🟡', '—', '—', '—', 'СЛАБО'),
    (11, 'Металл 3в1', '—', '—', '—', '—', '—', '🟢', '—', 'ЕСТЬ'),
    (12, 'ПФ-115 + ГФ-021', '🟢', '—', '—', '—', '🟢', '—', '—', 'ЕСТЬ'),
    (13, 'Полы не стираются', '🟢', '—', '—', '—', '🟢', '—', '—', 'ЕСТЬ'),
    (14, 'Яхтный лак дерево', '🟡', '—', '—', '—', '🟢', '—', '—', 'ЕСТЬ'),
    (15, 'Потолок без разводов', '🟢', '—', '—', '🟢', '—', '—', '—', 'ЕСТЬ'),
]

for idx, data in enumerate(matrix_data):
    row = 6 + idx
    for j, val in enumerate(data):
        cell = sc(ws4, row, j+2, val, ft(10), None, ac if j < 2 else ac)
        # Раскраска
        if '🟢' in str(val):
            cell.fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
        elif '🔴' in str(val):
            cell.fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
        elif '🟡' in str(val):
            cell.fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
    
    # Оценка — цвет
    assessment = data[-1]
    if 'ЕСТЬ' in assessment and 'НЕТ' not in assessment:
        ws4.cell(row=row, column=11).fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    elif 'НЕТ' in assessment:
        ws4.cell(row=row, column=11).fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    else:
        ws4.cell(row=row, column=11).fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')

# ---- ЛИСТ 5: РЕКОМЕНДАЦИИ ----
ws5 = wb.create_sheet('Рекомендации')
ws5.sheet_properties.tabColor = GREEN

ws5.column_dimensions['A'].width = 3
ws5.column_dimensions['B'].width = 5
ws5.column_dimensions['C'].width = 35
ws5.column_dimensions['D'].width = 50
ws5.column_dimensions['E'].width = 15
ws5.column_dimensions['F'].width = 20

ws5.merge_cells('B2:F2')
sc(ws5, 2, 2, 'РЕКОМЕНДАЦИИ: ЧТО ДОБАВИТЬ В АССОРТИМЕНТ ЛЕМАНА ПРО', ft(16, True, WHITE), pf(NAVY), ac)

for i, h in enumerate(['№', 'Продукт для добавления', 'Обоснование', 'Приоритет', 'Технология']):
    sc(ws5, 4, i+2, h, ft(10, True, WHITE), pf(CYAN))

recommendations = [
    (1, 'Краска ВД AVANTGARDE интерьерная моющаяся (база A/C)', 'Ключевой продукт премиум-цепочки «Идеальный интерьер». В ассортименте Леманы нет ни одной позиции AVANTGARDE по краскам ВД', 'КРИТИЧНО', 'Технология 3'),
    (2, 'Шпатлёвка для влажных помещений ОРЕОЛ с антисептиком', 'Необходима для цепочки «Влажная зона» — без неё система неполная', 'КРИТИЧНО', 'Технология 2'),
    (3, 'Шпатлёвка суперфинишная NEWTONE акриловая', 'Для премиум-интерьера — ювелирная гладкость стен', 'ВЫСОКИЙ', 'Технология 3'),
    (4, 'Шпатлёвка стартовая NEWTONE / AVANTGARDE', 'Базовое выравнивание для всех интерьерных цепочек', 'ВЫСОКИЙ', 'Технологии 2, 3'),
    (5, 'Бетоноконтакт ОРЕОЛ', 'Критичный продукт для покраски по плитке и старой краске', 'ВЫСОКИЙ', 'Технология 7'),
    (6, 'Грунтовка глубокого проникновения NEWTONE с антисептиком', 'Укрепление + антисептик — основа фасадных и интерьерных систем', 'ВЫСОКИЙ', 'Технологии 1, 3, 6'),
    (7, 'Грунт-концентрат влагоизолирующий ОРЕОЛ', 'Для влажных зон и OSB — без него система «Влажная зона» неполная', 'ВЫСОКИЙ', 'Технологии 2, 9'),
    (8, 'Краска для OSB ОРЕОЛ база А', 'Растущий сегмент SIP-домов, в ассортименте нет', 'СРЕДНИЙ', 'Технология 9'),
    (9, 'Эмаль для окон ОРЕОЛ полиуретановая акриловая', 'Anti-blocking эффект — ниша, не закрытая в ассортименте', 'СРЕДНИЙ', 'Технология 5'),
    (10, 'Лак яхтный AVANTGARDE', 'Премиум-сегмент, гибридная технология, не желтеет', 'СРЕДНИЙ', 'Технология 6'),
    (11, 'Краска фасадная силиконовая NEWTONE MODERN PRO', 'Эффект самоочищения — уникальное УТП для фасадов', 'СРЕДНИЙ', 'Технология 1'),
    (12, 'Шпатлёвка масляно-клеевая ДИСКОНТ', 'Для эконом-цепочки — минимальный бюджет', 'НИЗКИЙ', 'Технология 10'),
]

for idx, (num, product, reason, priority, tech) in enumerate(recommendations):
    row = 5 + idx
    sc(ws5, row, 2, num, ft(10, True))
    sc(ws5, row, 3, product, ft(10, True, NAVY), None, al)
    sc(ws5, row, 4, reason, ft(9), None, al)
    
    pcolor = {'КРИТИЧНО': 'C00000', 'ВЫСОКИЙ': 'ED7D31', 'СРЕДНИЙ': 'FFC000', 'НИЗКИЙ': '808080'}
    pc = pcolor.get(priority, '808080')
    sc(ws5, row, 5, priority, ft(10, True, WHITE), PatternFill(start_color=pc, end_color=pc, fill_type='solid'))
    sc(ws5, row, 6, tech, ft(9))

# Сохранение
wb.save(dst)
print(f"\n✅ Сохранено: {dst}")
