#!/usr/bin/env python3
"""
Generate Excel calendar 2027 for Empils brands.
Concept: "Цикл защиты" — Защити один раз — не перекрашивай 10 лет.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter

OUT_PATH = '/home/z/my-project/download/Календарь_2027_Цикл_защиты.xlsx'

# Colors
RED = 'C62828'
RED_DARK = '8E0000'
BLUE = '2563EB'
GREEN = '16A34A'
ORANGE = 'EA580C'
PURPLE = '7C3AED'
DARK = '1A1F36'
LIGHT = 'F5F5F5'
WHITE = 'FFFFFF'
GRAY = '9CA3AF'

wb = openpyxl.Workbook()

# ===== Styles =====
def thin_border():
    return Border(
        left=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),
        bottom=Side(style='thin', color='CCCCCC')
    )

def header_font(size=12, color=WHITE, bold=True):
    return Font(name='Calibri', size=size, bold=bold, color=color)

def body_font(size=10, color=DARK, bold=False):
    return Font(name='Calibri', size=size, bold=bold, color=color)

def fill(color):
    return PatternFill(start_color=color, end_color=color, fill_type='solid')

center = Alignment(horizontal='center', vertical='center', wrap_text=True)
left = Alignment(horizontal='left', vertical='center', wrap_text=True)
left_top = Alignment(horizontal='left', vertical='top', wrap_text=True)

# ===== SHEET 1: Обзор =====
ws1 = wb.active
ws1.title = 'Обзор 2027'

# Title
ws1.merge_cells('A1:H1')
ws1['A1'] = 'Календарь Эмпилс 2027'
ws1['A1'].font = Font(name='Calibri', size=24, bold=True, color=RED)
ws1['A1'].alignment = center
ws1.row_dimensions[1].height = 40

ws1.merge_cells('A2:H2')
ws1['A2'] = '«Цикл защиты» — Защити один раз — не перекрашивай 10 лет'
ws1['A2'].font = Font(name='Calibri', size=14, italic=True, color=GRAY)
ws1['A2'].alignment = center
ws1.row_dimensions[2].height = 25

ws1.row_dimensions[3].height = 10

# Quarter summary
headers = ['Квартал', 'Тема', 'Месяцы', 'Бренды', 'Направление', 'Слоган квартала']
for col, h in enumerate(headers, 1):
    c = ws1.cell(row=4, column=col, value=h)
    c.font = header_font(11)
    c.fill = fill(DARK)
    c.alignment = center
    c.border = thin_border()
ws1.row_dimensions[4].height = 30

quarters = [
    ['Q1', 'Металл', 'Январь · Февраль · Март', 'Ореол → TITAN', 'Органика\n(алкидные)', 'Защити металл на 10 лет'],
    ['Q2', 'Дерево', 'Апрель · Май · Июнь', 'Profiwood → Ореол (лаки)', 'Лаки', 'Защити дерево на 7 лет'],
    ['Q3', 'Интерьер', 'Июль · Август · Сентябрь', 'Newtone → Расцвет', 'Водные краски', 'Покрась без запаха и разводов'],
    ['Q4', 'Премиум и план', 'Октябрь · Ноябрь · Декабрь', 'AvantGarde → все бренды', 'Все направления', 'Заплатил больше — сэкономил на годы'],
]

colors_q = [RED, GREEN, BLUE, PURPLE]
for i, (q, theme, months, brands, direction, slogan) in enumerate(quarters):
    row = 5 + i
    ws1.row_dimensions[row].height = 50
    
    data = [q, theme, months, brands, direction, slogan]
    for col, val in enumerate(data, 1):
        c = ws1.cell(row=row, column=col, value=val)
        c.font = body_font(10, WHITE if col <= 2 else DARK, bold=(col <= 2))
        c.alignment = center if col <= 2 else left
        c.border = thin_border()
        if col <= 2:
            c.fill = fill(colors_q[i])
        elif col == 5:
            c.fill = fill('F0F4FF')
        else:
            c.fill = fill('FAFAFA')

# Column widths
ws1.column_dimensions['A'].width = 10
ws1.column_dimensions['B'].width = 15
ws1.column_dimensions['C'].width = 22
ws1.column_dimensions['D'].width = 22
ws1.column_dimensions['E'].width = 18
ws1.column_dimensions['F'].width = 32

# Brand summary
ws1.row_dimensions[10].height = 15
ws1.merge_cells('A11:F11')
ws1['A11'] = '6 фокусных брендов'
ws1['A11'].font = Font(name='Calibri', size=14, bold=True, color=DARK)
ws1['A11'].alignment = left

brand_headers = ['Бренд', 'Сайт', 'Направление', 'Квартал', 'Сезонность', 'Ключевой продукт']
for col, h in enumerate(brand_headers, 1):
    c = ws1.cell(row=12, column=col, value=h)
    c.font = header_font(10)
    c.fill = fill(DARK)
    c.alignment = center
    c.border = thin_border()
ws1.row_dimensions[12].height = 25

brands_data = [
    ['Ореол', 'oreol.pro', 'Органика (алкидные)', 'Q1', 'Весна', 'ГФ-021, ПФ-115'],
    ['TITAN', 'titanpaints.ru', 'Органика (алкидные)', 'Q1', 'Весна', 'Грунт-эмаль 3в1'],
    ['Profiwood', 'profiwood.pro', 'Защита дерева', 'Q2', 'Весна-Лето', 'Пропитки, антисептики'],
    ['Newtone', 'newtonepaints.ru', 'Водные краски', 'Q3', 'Лето-Осень', 'Интерьерные краски'],
    ['Расцвет', 'rascvet.pro', 'Водные/органика', 'Q3', 'Лето-Осень', 'Эмали для пола'],
    ['AvantGarde', 'avantgardepaints.ru', 'Премиум', 'Q4', 'Осень-Зима', 'Премиум-эмали'],
]
brand_colors = [RED, RED_DARK, GREEN, BLUE, ORANGE, PURPLE]

for i, brand in enumerate(brands_data):
    row = 13 + i
    ws1.row_dimensions[row].height = 22
    for col, val in enumerate(brand, 1):
        c = ws1.cell(row=row, column=col, value=val)
        c.font = body_font(10, WHITE if col == 1 else DARK, bold=(col == 1))
        c.alignment = center if col <= 1 else left
        c.border = thin_border()
        if col == 1:
            c.fill = fill(brand_colors[i])
        elif i % 2 == 0:
            c.fill = fill('FAFAFA')

# TSO principle
ws1.row_dimensions[20].height = 15
ws1.merge_cells('A21:F21')
ws1['A21'] = 'Принцип ТСО (Совокупная стоимость владения)'
ws1['A21'].font = Font(name='Calibri', size=14, bold=True, color=RED)
ws1['A21'].alignment = left

ws1.merge_cells('A22:F22')
ws1['A22'] = '«Защити один раз — не перекрашивай 10 лет». Каждый месяц календаря содержит мини-расчёт ТСО, показывающий, что дорогой продукт выгоднее дешёвого в долгосрочной перспективе.'
ws1['A22'].font = body_font(11, DARK)
ws1['A22'].alignment = left_top
ws1.row_dimensions[22].height = 45

# ===== SHEET 2-5: Quarterly detail =====
quarter_data = {
    'Q1 Металл': {
        'color': RED,
        'months': [
            {
                'month': 'Январь',
                'theme': 'Подготовка металла',
                'brand': 'Ореол',
                'product': 'Грунтовка ГФ-021',
                'direction': 'Органика (алкидная)',
                'focus': 'Грунтование — основа защиты. ГФ-021 блокирует ржавчину, создаёт адгезию.',
                'tso': 'Без грунта: ПФ-115 отслоится через 1 сезон.\nС грунтом ГФ-021: покрытие держится 3-5 лет.\nС TITAN 3в1 (грунт встроен): 10 лет.',
                'action': 'Промо: скидка на ГФ-021 при покупке ПФ-115. Обучение: «Зачем грунтовать?»',
                'audience': 'Дачники, владельцы гаражей',
            },
            {
                'month': 'Февраль',
                'theme': 'Антикоррозийная защита',
                'brand': 'Ореол',
                'product': 'Эмаль ПФ-115',
                'direction': 'Органика (алкидная)',
                'focus': 'ПФ-115 — классика для подготовленного металла. 8 цветов, недорого, проверен.',
                'tso': 'ПФ-115 + ГФ-021: 3-5 лет службы, 4 перекраски за 10 лет = 13 860 руб.\nTITAN 3в1: 1 покраска на 10 лет = 5 320 руб.\nЭкономия: 8 540 руб. (61%)',
                'action': 'Промо: «Готовь металл весной». Калькулятор ТСО: ПФ-115 vs TITAN.',
                'audience': 'Мастера, бригады, DIY',
            },
            {
                'month': 'Март',
                'theme': 'Премиум-защита металла',
                'brand': 'TITAN',
                'product': 'Грунт-эмаль 3в1 по ржавчине',
                'direction': 'Органика (алкидная)',
                'focus': 'TITAN 3в1 — 1 банка вместо 3 (грунт + эмаль + преобразователь). 9 цветов RAL.',
                'tso': '1 банка TITAN 3в1 = 3 продукта.\n2 слоя за 1 день вместо 3 слоёв за 2 дня.\nСрок службы: 10 лет (подтверждено).\nЭкономия: 61% за 10 лет.',
                'action': 'Главный месяц TITAN! Промо-дни, выкрасы, калькулятор ТСО в магазинах.',
                'audience': 'Все сегменты',
            },
        ],
    },
    'Q2 Дерево': {
        'color': GREEN,
        'months': [
            {
                'month': 'Апрель',
                'theme': 'Защита дерева снаружи',
                'brand': 'Profiwood',
                'product': 'Пропитки, антисептики',
                'direction': 'Защита дерева',
                'focus': 'Дерево на улице без защиты гниёт за 2-3 года. Profiwood — до 7 лет защиты.',
                'tso': 'Без защиты: замена доски через 3 года = 15 000 руб./м²\nProfiwood: защита 7 лет, замена через 15+ лет\nЭкономия: 60%+ на замене материала',
                'action': 'Промо: «Сезон террас и заборов». Набор: пропитка + кисть в подарок.',
                'audience': 'Дачники, владельцы домов',
            },
            {
                'month': 'Май',
                'theme': 'Отделка дерева',
                'brand': 'Profiwood / Ореол',
                'product': 'Декоративные пропитки, эмали по дереву',
                'direction': 'Лаки / Органика',
                'focus': 'Цвет и защита одновременно. Прозрачные и укрывистые покрытия для дерева.',
                'tso': 'Обычный лак: 2-3 года, выгорает на УФ.\nProfiwood с УФ-фильтром: 7 лет без выгорания.\nПерекраска в 2 раза реже = 50% экономии.',
                'action': 'Промо: выкрасы Profiwood vs обычный лак. Кейсы: терраса, забор, беседка.',
                'audience': 'Дачники, строители',
            },
            {
                'month': 'Июнь',
                'theme': 'Лакировка и финиш',
                'brand': 'Ореол',
                'product': 'Лаки яхтные, паркетные',
                'direction': 'Лаки',
                'focus': 'Финишное покрытие: блеск, износостойкость, УФ-защита. Яхтный лак — для улицы.',
                'tso': 'Паркетный лак эконом: 2-3 года до истирания.\nЯхтный лак Ореол: 5-7 лет.\nПерекрытие реже = меньше труда и расхода.',
                'action': 'Промо: «Защити лодку, пол, лестницу». Демонстрация износостойкости.',
                'audience': 'Лодочники, владельцы паркетных полов',
            },
        ],
    },
    'Q3 Интерьер': {
        'color': BLUE,
        'months': [
            {
                'month': 'Июль',
                'theme': 'Стены без разводов',
                'brand': 'Newtone',
                'product': 'Интерьерные водные краски',
                'direction': 'Водные краски',
                'focus': 'Newtone — инновации: 1 слой вместо 3. Без запаха. Моющиеся стены.',
                'tso': 'Обычная краска: 3 слоя = 3 дня работы + расход 3 банок.\nNewtone: 1-2 слоя = 1 день + 2 банки.\nЭкономия: 30-50% времени и материала.',
                'action': 'Промо: «Покрась стены за выходные». Демонстрация укрывистости.',
                'audience': 'Семьи, делающие ремонт',
            },
            {
                'month': 'Август',
                'theme': 'Потолки без полос',
                'brand': 'Newtone / Расцвет',
                'product': 'Потолочные краски',
                'direction': 'Водные краски',
                'focus': 'Потолочная краска, которая ложится без разводов. Белоснежный результат.',
                'tso': 'Дешёвая потолочная: 2 слоя, видны полосы, через 3 года желтеет.\nNewtone: 1 слой, без полос, 7 лет без желтизны.\nПеределка потолка = самая дорогая работа.',
                'action': 'Промо: «Белый потолок без разводов». Кейсы: до/после.',
                'audience': 'Семьи, мастера',
            },
            {
                'month': 'Сентябрь',
                'theme': 'Полы, которые не стираются',
                'brand': 'Расцвет',
                'product': 'Эмаль для бетонных полов',
                'direction': 'Водные/Органика',
                'focus': 'Расцвет — износостойкая эмаль для пола. Гараж, склад, подвал — не пылит.',
                'tso': 'Бетон без покрытия: пылит, впитывает масло, разрушается за 5 лет.\nРасцвет: не пылит, легко мыть, 7-10 лет службы.\nЗамена бетона = 10 000 руб./м². Покраска = 300 руб./м².',
                'action': 'Промо: «Пол в гараже — как на заводе». Демонстрация до/после.',
                'audience': 'Владельцы гаражей, складов',
            },
        ],
    },
    'Q4 Премиум и план': {
        'color': PURPLE,
        'months': [
            {
                'month': 'Октябрь',
                'theme': 'Премиум-покраска',
                'brand': 'AvantGarde',
                'product': 'Премиум-эмали',
                'direction': 'Все направления',
                'focus': 'AvantGarde — обновлённая формула, дизайн, премиум-позиционирование.',
                'tso': 'AvantGarde: выше укрывистость = меньше слоёв = меньше расхода.\nЦена банки выше, но расход ниже → стоимость м² сопоставима.\nДолговечность: 10+ лет.',
                'action': 'Промо: «Попробуй премиум». Выкрасы AvantGarde vs стандарт.',
                'audience': 'Профессионалы, премиум-сегмент',
            },
            {
                'month': 'Ноябрь',
                'theme': 'Декор и спецприменения',
                'brand': 'Все бренды',
                'product': 'Молотковый эффект, декоративные',
                'direction': 'Все направления',
                'focus': 'Спецэффекты: молотковая поверхность, текстуры, декоративные покрытия.',
                'tso': 'Декоративное покрытие = не только защита, но и эстетика.\nПовышает стоимость объекта (ворота, двери, металл).',
                'action': 'Промо: «Не просто краска — декор». Каталог эффектов.',
                'audience': 'Архитекторы, дизайнеры, премиум',
            },
            {
                'month': 'Декабрь',
                'theme': 'План на 2028 год',
                'brand': 'Все бренды',
                'product': 'Вся линейка Эмпилс',
                'direction': 'Все направления',
                'focus': 'Подведение итогов года. Презентация плана. Закупки на следующий сезон.',
                'tso': 'ТСО-калькулятор: планируй покраску на год вперёд.\nЗакупи правильный продукт один раз — не перекрашивай 10 лет.',
                'action': 'Промо: «Встречай 2028 год защищённым». Годовой итог, награды партнёрам.',
                'audience': 'Дистрибьюторы, партнёры',
            },
        ],
    },
}

for sheet_name, q_data in quarter_data.items():
    ws = wb.create_sheet(title=sheet_name)
    color = q_data['color']
    
    # Title
    ws.merge_cells('A1:H1')
    ws['A1'] = sheet_name
    ws['A1'].font = Font(name='Calibri', size=20, bold=True, color=WHITE)
    ws['A1'].fill = fill(color)
    ws['A1'].alignment = center
    ws.row_dimensions[1].height = 35
    
    ws.merge_cells('A2:H2')
    ws['A2'] = '«Защити один раз — не перекрашивай 10 лет»'
    ws['A2'].font = Font(name='Calibri', size=12, italic=True, color=color)
    ws['A2'].alignment = center
    ws.row_dimensions[2].height = 22
    
    ws.row_dimensions[3].height = 8
    
    # Headers
    headers = ['Месяц', 'Тема', 'Бренд', 'Продукт', 'Направление', 'Фокус месяца', 'ТСО-расчёт', 'Действия и аудитория']
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=4, column=col, value=h)
        c.font = header_font(10, WHITE)
        c.fill = fill(color)
        c.alignment = center
        c.border = thin_border()
    ws.row_dimensions[4].height = 28
    
    # Data
    for i, m in enumerate(q_data['months']):
        row = 5 + i
        ws.row_dimensions[row].height = 90
        
        data = [m['month'], m['theme'], m['brand'], m['product'], m['direction'], 
                m['focus'], m['tso'], m['action'] + '\n\nАудитория: ' + m['audience']]
        
        for col, val in enumerate(data, 1):
            c = ws.cell(row=row, column=col, value=val)
            c.font = body_font(9, WHITE if col == 1 else DARK, bold=(col <= 2))
            c.alignment = center if col == 1 else left_top
            c.border = thin_border()
            if col == 1:
                c.fill = fill(color)
            elif col == 7:
                c.fill = fill('F0FDF4')  # Light green for TSO
            elif i % 2 == 0:
                c.fill = fill('FAFAFA')
    
    # Column widths
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 14
    ws.column_dimensions['D'].width = 20
    ws.column_dimensions['E'].width = 16
    ws.column_dimensions['F'].width = 35
    ws.column_dimensions['G'].width = 40
    ws.column_dimensions['H'].width = 35

# ===== SHEET 6: Сводная матрица =====
ws6 = wb.create_sheet(title='Сводная матрица')

ws6.merge_cells('A1:G1')
ws6['A1'] = 'Сводная матрица: Месяцы × Бренды × Направления'
ws6['A1'].font = Font(name='Calibri', size=16, bold=True, color=DARK)
ws6['A1'].alignment = center
ws6.row_dimensions[1].height = 30

headers6 = ['Месяц', 'Квартал', 'Тема', 'Бренд', 'Направление', 'Ключевой продукт', 'ТСО-вывод']
for col, h in enumerate(headers6, 1):
    c = ws6.cell(row=3, column=col, value=h)
    c.font = header_font(10, WHITE)
    c.fill = fill(DARK)
    c.alignment = center
    c.border = thin_border()
ws6.row_dimensions[3].height = 25

all_months = []
for q_name, q_data in quarter_data.items():
    for m in q_data['months']:
        # Extract TSO one-liner
        tso_lines = m['tso'].split('\n')
        tso_short = tso_lines[-1] if len(tso_lines) > 1 else tso_lines[0]
        all_months.append([
            m['month'],
            q_name.split()[0],
            m['theme'],
            m['brand'],
            m['direction'],
            m['product'],
            tso_short,
        ])

q_colors = {'Q1': RED, 'Q2': GREEN, 'Q3': BLUE, 'Q4': PURPLE}

for i, row_data in enumerate(all_months):
    row = 4 + i
    ws6.row_dimensions[row].height = 35
    q = row_data[1]
    for col, val in enumerate(row_data, 1):
        c = ws6.cell(row=row, column=col, value=val)
        c.font = body_font(9, WHITE if col == 2 else DARK, bold=(col <= 2))
        c.alignment = center if col <= 2 else left
        c.border = thin_border()
        if col == 2:
            c.fill = fill(q_colors.get(q, GRAY))
        elif col == 7:
            c.fill = fill('F0FDF4')
        elif i % 2 == 0:
            c.fill = fill('FAFAFA')

ws6.column_dimensions['A'].width = 12
ws6.column_dimensions['B'].width = 10
ws6.column_dimensions['C'].width = 22
ws6.column_dimensions['D'].width = 16
ws6.column_dimensions['E'].width = 18
ws6.column_dimensions['F'].width = 24
ws6.column_dimensions['G'].width = 45

# Freeze panes
ws6.freeze_panes = 'A4'

# Save
wb.save(OUT_PATH)
print(f'Excel saved: {OUT_PATH}')
print(f'Size: {os.path.getsize(OUT_PATH)} bytes')
print(f'Sheets: {wb.sheetnames}')
