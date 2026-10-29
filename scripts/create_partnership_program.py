#!/usr/bin/env python3
"""
Generate partnership program DOCX for "Центр Красок" (Krasnodar, ЮФО).
Based on meeting protocol from 01.07.2026.
Period: 06.07.2026 — 31.12.2026.
"""

import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT_PATH = '/home/z/my-project/download/Программа_сотрудничества_Центр_Красок_Краснодар_2026.docx'

P = {
    'primary': RGBColor(0xC6, 0x28, 0x28),
    'dark': RGBColor(0x1A, 0x1F, 0x36),
    'body': RGBColor(0x2D, 0x31, 0x42),
    'secondary': RGBColor(0x5A, 0x60, 0x80),
    'accent': RGBColor(0x16, 0xA3, 0x4A),
    'warn': RGBColor(0xEA, 0x58, 0x0C),
}


def set_cell_bg(cell, color_hex):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tc_pr.append(shd)


def set_cell_borders(cell, color='CCCCCC', size='4'):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = OxmlElement('w:tcBorders')
    for side in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{side}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), size)
        border.set(qn('w:color'), color)
        tc_borders.append(border)
    tc_pr.append(tc_borders)


def add_heading(doc, text, level=1, color=None):
    if color is None:
        color = P['primary'] if level == 1 else P['dark']
    h = doc.add_heading('', level=level)
    run = h.add_run(text)
    run.font.color.rgb = color
    run.font.name = 'Calibri'
    return h


def add_para(doc, text, size=11, bold=False, italic=False, color=None, indent=False):
    p = doc.add_paragraph()
    if indent:
        p.paragraph_format.first_line_indent = Cm(0.75)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.3
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = 'Calibri'
    if color:
        run.font.color.rgb = color
    return p


def add_bullet(doc, text):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.3
    run = p.add_run(text)
    run.font.size = Pt(11)
    run.font.name = 'Calibri'


def add_numbered(doc, text):
    p = doc.add_paragraph(style='List Number')
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.3
    run = p.add_run(text)
    run.font.size = Pt(11)
    run.font.name = 'Calibri'


def make_table(doc, headers, rows, col_widths=None, header_color='C62828'):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    if col_widths:
        for i, width in enumerate(col_widths):
            for cell in table.columns[i].cells:
                cell.width = Cm(width)
    header_cells = table.rows[0].cells
    for i, header in enumerate(headers):
        cell = header_cells[i]
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(header)
        run.font.bold = True
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        run.font.name = 'Calibri'
        set_cell_bg(cell, header_color)
        set_cell_borders(cell)
    for r_idx, row_data in enumerate(rows):
        row_cells = table.rows[r_idx + 1].cells
        for c_idx, cell_text in enumerate(row_data):
            cell = row_cells[c_idx]
            cell.text = ''
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.25
            run = p.add_run(str(cell_text))
            run.font.size = Pt(10)
            run.font.name = 'Calibri'
            set_cell_borders(cell)
            if r_idx % 2 == 0:
                set_cell_bg(cell, 'FAF5F5')
    return table


# ====== Build Document ======
doc = Document()
style = doc.styles['Normal']
style.font.name = 'Calibri'
style.font.size = Pt(11)
style.paragraph_format.line_spacing = 1.3

section = doc.sections[0]
section.left_margin = Cm(2.5)
section.right_margin = Cm(2.0)
section.top_margin = Cm(2.0)
section.bottom_margin = Cm(2.0)

# ===== COVER =====
cover_table = doc.add_table(rows=1, cols=1)
cover_cell = cover_table.rows[0].cells[0]
set_cell_bg(cover_cell, 'C62828')
cover_cell.text = ''
p = cover_cell.paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(40)
p.paragraph_format.space_after = Pt(20)
run = p.add_run('Empils')
run.font.size = Pt(36)
run.font.bold = True
run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

p2 = cover_cell.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
p2.paragraph_format.space_after = Pt(40)
run = p2.add_run('Программа партнёрского сотрудничества')
run.font.size = Pt(20)
run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

for _ in range(2):
    doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Дистрибьютор: «Центр Красок»')
run.font.size = Pt(26)
run.font.bold = True
run.font.color.rgb = P['dark']

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Краснодар • Южный федеральный округ (ЮФО)')
run.font.size = Pt(16)
run.font.color.rgb = P['secondary']

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(20)
run = p.add_run('Период действия программы')
run.font.size = Pt(14)
run.font.italic = True
run.font.color.rgb = P['secondary']

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('06 июля 2026 г. — 31 декабря 2026 г.')
run.font.size = Pt(18)
run.font.bold = True
run.font.color.rgb = P['primary']

for _ in range(2):
    doc.add_paragraph()
cond_table = doc.add_table(rows=1, cols=1)
cond_cell = cond_table.rows[0].cells[0]
set_cell_bg(cond_cell, 'F5F5F5')
set_cell_borders(cond_cell, 'C62828', '8')
cond_cell.text = ''
p = cond_cell.paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(10)
run = p.add_run('Ключевые условия программы')
run.font.size = Pt(14)
run.font.bold = True
run.font.color.rgb = P['primary']

for line in [
    '✓ Ежемесячный бонус 8% + полугодовой бонус 3%',
    '✓ Отсрочка оплаты до 60 дней',
    '✓ Выездное обучение торговой команды',
    '✓ Экскурсия на завод Empils',
    '✓ Обеспечение выкрасами по TITAN',
]:
    p = cond_cell.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(line)
    run.font.size = Pt(11)
    run.font.name = 'Calibri'
    run.font.color.rgb = P['body']

for _ in range(1):
    doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Empils • TITAN • 2026')
run.font.size = Pt(11)
run.font.color.rgb = P['secondary']

doc.add_page_break()

# ===== 1. ОБЩИЕ ПОЛОЖЕНИЯ =====
add_heading(doc, '1. Общие положения', level=1)

add_para(doc,
    'Настоящая программа определяет условия партнёрского сотрудничества между компанией '
    '«Эмпилс» (производитель лакокрасочных материалов, бренд TITAN) и дистрибьютором '
    '«Центр Красок» (Краснодар) на период с 6 июля по 31 декабря 2026 года.',
    indent=True)

info_data = [
    ['Производитель', 'АО «Эмпилс» (бренд TITAN)'],
    ['Дистрибьютор', '«Центр Красок», г. Краснодар'],
    ['Территория распространения', 'Южный федеральный округ (ЮФО)'],
    ['Период действия', '06.07.2026 — 31.12.2026 (6 месяцев)'],
    ['Основание', 'Протокол встречи от 01.07.2026'],
    ['Цель программы', 'Усиление партнёрства, рост продаж TITAN, обучение торговой команды'],
]
make_table(doc, ['Параметр', 'Значение'], info_data, col_widths=[5, 12])

add_para(doc, 'Программа разработана на основе обратной связи, полученной в ходе встречи '
              'с руководителем подразделения «Центр Красок» 1 июля 2026 года, и учитывает '
              'ключевые потребности партнёра в маркетинговой поддержке, обучении и '
              'финансовых условиях.', indent=True, italic=True, color=P['secondary'])

doc.add_page_break()

# ===== 2. ФИНАНСОВЫЕ УСЛОВИЯ =====
add_heading(doc, '2. Финансовые условия', level=1)

add_heading(doc, '2.1. Бонусная система', level=2, color=P['dark'])

add_para(doc,
    'Дистрибьютор получает двухуровневую бонусную систему, включающую ежемесячный и '
    'полугодовой бонусы. Бонусы являются безоговорочными и не зависят от выполнения '
    'дополнительных условий.', indent=True)

bonus_data = [
    ['Ежемесячный бонус', '8% от объёма закупок', 'Выплачивается ежемесячно, безоговорочно'],
    ['Полугодовой бонус', '3% от объёма закупок за полугодие', 'Выплачивается по итогам полугодия (июль—декабрь 2026)'],
    ['Итого совокупный бонус', 'До 11% от объёма закупок', 'При выполнении всех условий программы'],
]
make_table(doc, ['Бонус', 'Размер', 'Условия выплаты'], bonus_data, col_widths=[4, 5, 8])

add_para(doc, 'Пример расчёта бонуса:', bold=True, color=P['primary'])
add_para(doc,
    'При ежемесячном объёме закупок 1 000 000 руб.:\n'
    '• Ежемесячный бонус: 1 000 000 × 8% = 80 000 руб./мес.\n'
    '• Полугодовой бонус (за 6 мес.): 6 000 000 × 3% = 180 000 руб.\n'
    '• Итого за полугодие: 80 000 × 6 + 180 000 = 660 000 руб.',
    indent=True)

add_para(doc, 'Порядок выплаты:', bold=True, color=P['primary'])
add_bullet(doc, 'Ежемесячный бонус начисляется в течение 10 рабочих дней после окончания расчётного месяца')
add_bullet(doc, 'Полугодовой бонус выплачивается в течение 20 рабочих дней после окончания полугодия')
add_bullet(doc, 'Выплата производится путём зачёта в счёт будущих поставок или денежным переводом по согласованию')
add_bullet(doc, 'Бонусы не облагаются дополнительными условиями по ассортименту или объёму')

add_heading(doc, '2.2. Отсрочка платежа', level=2, color=P['dark'])

add_para(doc,
    'Отсрочка оплаты за поставленную продукцию увеличивается до 60 календарных дней. '
    'Это улучшает оборачиваемость средств дистрибьютора и снижает финансовую нагрузку.',
    indent=True)

payment_data = [
    ['Стандартная отсрочка', '30 дней'],
    ['Программная отсрочка (по данной программе)', '60 дней'],
    ['Условие предоставления', 'Соблюдение графика платежей по предыдущим поставкам'],
    ['Ответственность за просрочку', 'Приостановка отгрузок до погашения задолженности'],
]
make_table(doc, ['Параметр', 'Значение'], payment_data, col_widths=[6, 11])

add_para(doc, '⚠️ Важно: Увеличенная отсрочка 60 дней предоставляется в рамках данной программы '
              'и действует до 31 декабря 2026 года. С 1 января 2027 года условия отсрочки '
              'пересматриваются по итогам программы.', italic=True, color=P['warn'])

doc.add_page_break()

# ===== 3. ОБУЧЕНИЕ =====
add_heading(doc, '3. Выездное обучение', level=1)

add_para(doc,
    'Компания «Эмпилс» организует выездное обучение для торговой и обучающей команды '
    'дистрибьютора «Центр Красок». Обучение проводится на территории партнёра или '
    'в удобном для него формате (офлайн + онлайн).', indent=True)

add_heading(doc, '3.1. Программа обучения', level=2, color=P['dark'])

add_para(doc, 'Обучение включает следующие модули:', indent=True)

training_data = [
    ['1', 'О компании «Эмпилс»', 'История, производство, качество, сертификаты, гарантии', '45 мин'],
    ['2', 'О продуктах TITAN', 'Ассортимент, технические характеристики, области применения, новинки', '60 мин'],
    ['3', 'Упор на TITAN 3в1', 'Грунт-эмаль по ржавчине: преимущества, сравнение с ПФ-115, применение', '60 мин'],
    ['4', 'Методика ТСО', 'Совокупная стоимость владения: расчёт, аргументы для покупателя, кейсы', '60 мин'],
    ['5', 'Примеры ТСО для розницы', 'Готовые расчёты для розничных точек: ворота, гараж, забор, трубы', '45 мин'],
    ['6', 'Примеры ТСО для клиентов', 'B2B-кейсы: заводы, склады, фермы — расчёт экономии за 5–10 лет', '45 мин'],
    ['7', 'Практикум: выкрасы', 'Демонстрация TITAN vs Hammerite vs ПФ-115 — сравнение в реальности', '60 мин'],
    ['8', 'Работа с возражениями', 'Топ-10 возражений покупателей и отработка по методике ТСО', '45 мин'],
    ['9', 'Цифровые инструменты', 'Калькулятор ТСО, QR-код на банке, онлайн-платформа обучения', '30 мин'],
    ['10', 'Тестирование', 'Экзамен: 20 вопросов по продукту и ТСО. Сертификат участника', '30 мин'],
]
make_table(doc, ['№', 'Модуль', 'Содержание', 'Длительность'], training_data,
           col_widths=[0.8, 4, 8.5, 2.5])

add_para(doc, 'Общая длительность обучения: ~8 часов (1 полный день или 2 дня по 4 часа).',
         bold=True, color=P['primary'])

add_heading(doc, '3.2. Формат и условия проведения', level=2, color=P['dark'])

add_para(doc, 'Дата и время проведения:', bold=True)
add_bullet(doc, 'Согласовываются сторонами не позднее чем за 14 дней до даты обучения')
add_bullet(doc, 'Предпочтительный период: июль—август 2026 г.')
add_bullet(doc, 'Место: на территории дистрибьютора (г. Краснодар) или в удобном помещении')

add_para(doc, 'Аудитория:', bold=True)
add_bullet(doc, 'Торговая команда дистрибьютора (торговые представители, супервайзеры)')
add_bullet(doc, 'Обучающая команда (тренеры, методисты)')
add_bullet(doc, 'Рекомендуемое количество участников: 15–25 человек')
add_bullet(doc, 'Возможно разделение на 2 потока: торговая команда + продавцы магазинов')

add_para(doc, 'Материалы для участников:', bold=True)
add_bullet(doc, 'Печатная шпаргалка продавца TITAN 3в1')
add_bullet(doc, 'Сравнительная таблица TITAN 3в1 vs ПФ-115 vs Hammerite')
add_bullet(doc, 'Калькулятор ТСО (PDF + онлайн-доступ)')
add_bullet(doc, 'Образцы покрытий (до/после)')
add_bullet(doc, 'Сертификат участника обучения')

add_para(doc, 'Расходы:', bold=True)
add_bullet(doc, 'Командировочные расходы тренера «Эмпилс» — за счёт «Эмпилс»')
add_bullet(doc, 'Аренда помещения (при необходимости) — за счёт дистрибьютора')
add_bullet(doc, 'Материалы участников — за счёт «Эмпилс»')

doc.add_page_break()

# ===== 4. ЭКСКУРСИЯ НА ЗАВОД =====
add_heading(doc, '4. Экскурсия на завод «Эмпилс»', level=1)

add_para(doc,
    'Компания «Эмпилс» приглашает торговую команду и руководство дистрибьютора «Центр Красок» '
    'на экскурсию на производственный завод в г. Ростов-на-Дону. Экскурсия направлена на '
    'укрепление лояльности и глубокое понимание продукта.', indent=True)

excursion_data = [
    ['Кого приглашаем', 'Торговая команда + руководство (до 10 человек)'],
    ['Где', 'Завод «Эмпилс», г. Ростов-на-Дону'],
    ['Когда', 'Согласовывается сторонами (предпочтительно сентябрь—октябрь 2026)'],
    ['Длительность', '1 день (с программой 8:00 — 17:00)'],
    ['Программа', 'Экскурсия по производству + дегустация продуктов + встреча с руководством'],
]
make_table(doc, ['Параметр', 'Значение'], excursion_data, col_widths=[4, 13])

add_heading(doc, '4.1. Программа экскурсии', level=2, color=P['dark'])

add_numbered(doc, '08:00 — Встреча делегации, приветственный кофе')
add_numbered(doc, '08:30 — Презентация компании «Эмпилс»: история, миссия, производственные мощности')
add_numbered(doc, '09:30 — Экскурсия по заводу: производственные линии, лаборатория контроля качества, склад')
add_numbered(doc, '11:30 — Дегустация продуктов: TITAN 3в1, Avangarde, Newton — выкрасы и сравнение')
add_numbered(doc, '13:00 — Обед с руководством компании')
add_numbered(doc, '14:00 — Презентация новинок и планы развития ассортимента')
add_numbered(doc, '15:00 — Круглый стол: обсуждение программы сотрудничества, обратная связь')
add_numbered(doc, '16:30 — Вручение сувениров и сертификатов участников экскурсии')
add_numbered(doc, '17:00 — Завершение программы')

add_para(doc, 'Расходы:', bold=True, color=P['primary'])
add_bullet(doc, 'Программа экскурсии, обед, сувениры — за счёт «Эмпилс»')
add_bullet(doc, 'Проезд до Ростова-на-Дону и обратно — по согласованию (50/50)')
add_bullet(doc, 'Проживание (при необходимости) — по согласованию')

doc.add_page_break()

# ===== 5. ВЫКРАСЫ TITAN =====
add_heading(doc, '5. Обеспечение выкрасами по TITAN', level=1)

add_para(doc,
    'Компания «Эмпилс» обеспечивает дистрибьютора «Центр Красок» образцами покрытий (выкрасами) '
    'по продукту TITAN для демонстрации покупателям в розничных точках и на обучающих мероприятиях.',
    indent=True)

add_heading(doc, '5.1. Состав набора выкрасов', level=2, color=P['dark'])

samples_data = [
    ['1', 'TITAN 3в1 — чёрный (RAL 9005)', 'Образец покрытия на металле, размер 10×15 см'],
    ['2', 'TITAN 3в1 — белый (RAL 9010)', 'Образец покрытия на металле, размер 10×15 см'],
    ['3', 'TITAN 3в1 — красно-коричневый (RAL 3009)', 'Образец покрытия на металле, размер 10×15 см'],
    ['4', 'TITAN 3в1 — серый (RAL 7004)', 'Образец покрытия на металле, размер 10×15 см'],
    ['5', 'TITAN 3в1 — зелёный (RAL 6005)', 'Образец покрытия на металле, размер 10×15 см'],
    ['6', 'Сравнительный образец: ПФ-115 vs TITAN 3в1', 'Два образца рядом для наглядного сравнения'],
    ['7', 'Сравнительный образец: Hammerite vs TITAN 3в1', 'Два образца рядом для наглядного сравнения'],
    ['8', 'Образец «до/после через 1 год»', 'Демонстрация долговечности покрытия'],
]
make_table(doc, ['№', 'Наименование', 'Описание'], samples_data, col_widths=[0.8, 6, 9])

add_heading(doc, '5.2. Условия предоставления', level=2, color=P['dark'])

add_para(doc, 'Количество и формат:', bold=True)
add_bullet(doc, 'Базовый набор — 1 комплект на каждый филиал дистрибьютора')
add_bullet(doc, 'Дополнительные наборы — по запросу, не более 5 комплектов на филиал')
add_bullet(doc, 'Возможно изготовление индивидуальных выкрасов по запросу партнёра')

add_para(doc, 'Сроки предоставления:', bold=True)
add_bullet(doc, 'Первая поставка — в течение 14 дней после подписания программы (до 20 июля 2026)')
add_bullet(doc, 'Дополнительные поставки — по запросу, в течение 21 дня')

add_para(doc, 'Стоимость:', bold=True)
add_bullet(doc, 'Выкрасы предоставляются БЕСПЛАТНО в рамках программы партнёрства')
add_bullet(doc, 'Доставка — за счёт «Эмпилс»')

add_para(doc, 'Использование:', bold=True)
add_bullet(doc, 'Демонстрация покупателям в розничных точках')
add_bullet(doc, 'Использование на обучающих мероприятиях для продавцов')
add_bullet(doc, 'Выставки, презентации, мероприятия для клиентов')

doc.add_page_break()

# ===== 6. КАЛЕНДАРНЫЙ ПЛАН =====
add_heading(doc, '6. Календарный план реализации программы', level=1)

add_para(doc, 'Программа реализуется в период с 6 июля по 31 декабря 2026 года по следующему графику:',
         indent=True)

calendar_data = [
    ['06.07.2026', 'Старт программы', 'Подписание программы, начало действия условий'],
    ['До 20.07.2026', 'Первая поставка выкрасов', 'Комплект выкрасов по TITAN для всех филиалов'],
    ['Июль—Авг 2026', 'Выездное обучение', 'Обучение торговой команды (10 модулей, 8 часов)'],
    ['Ежемесячно', 'Выплата ежемесячного бонуса', '8% от объёма закупок, в течение 10 раб. дней'],
    ['Сент—Окт 2026', 'Экскурсия на завод', 'Поездка делегации (до 10 чел.) в Ростов-на-Дону'],
    ['Окт—Ноя 2026', 'Дополнительное обучение', 'При необходимости — повторное обучение или онлайн-сессия'],
    ['31.12.2026', 'Окончание программы', 'Подведение итогов, расчёт полугодового бонуса'],
    ['Янв 2027', 'Выплата полугодового бонуса', '3% от объёма закупок за полугодие'],
    ['Янв 2027', 'Анализ результатов', 'Оценка эффективности, решение о продлении программы'],
]
make_table(doc, ['Срок', 'Мероприятие', 'Описание'], calendar_data, col_widths=[3, 4, 9])

doc.add_page_break()

# ===== 7. ОЖИДАЕМЫЕ РЕЗУЛЬТАТЫ =====
add_heading(doc, '7. Ожидаемые результаты', level=1)

add_para(doc, 'По итогам реализации программы (6 месяцев) ожидаются следующие результаты:',
         indent=True)

results_data = [
    ['Рост объёмов продаж TITAN', '+20–30% к уровню первого полугодия 2026'],
    ['Рост доли TITAN в ассортименте', 'Восстановление сокращённого ассортимента + расширение'],
    ['Обучение торговой команды', 'Не менее 20 сотрудников прошли обучение и тестирование'],
    ['Готовность продавать через ТСО', '≥80% торговой команды применяют методику ТСО'],
    ['Экскурсия на завод', 'Делегация 8–10 человек посетила производство'],
    ['Удовлетворённость партнёра', 'Высокая — партнёр готов к долгосрочному сотрудничеству'],
    ['Лояльность амбассадора', 'Поддержка и усиление ключевого сотрудника-амбассадора'],
]
make_table(doc, ['Показатель', 'Целевое значение'], results_data, col_widths=[7, 9])

add_para(doc, 'Дополнительные качественные результаты:', bold=True, color=P['primary'])
add_bullet(doc, 'Усиление позиций TITAN на полке относительно конкурентов (AkzoNobel, Hammerite)')
add_bullet(doc, 'Повышение узнаваемости бренда TITAN среди розничных покупателей ЮФО')
add_bullet(doc, 'Формирование базы кейсов ТСО для розничных точек и B2B-клиентов')
add_bullet(doc, 'Создание платформы для долгосрочного партнёрства на 2027+ годы')

doc.add_page_break()

# ===== 8. КООРДИНАЦИЯ =====
add_heading(doc, '8. Координация и контакты', level=1)

add_para(doc, 'Для реализации программы назначаются ответные стороны:', indent=True)

contacts_data = [
    ['Со стороны «Эмпилс»', 'Представитель по работе с партнёрами ЮФО', 'Согласование обучения, экскурсии, выкрасов'],
    ['Со стороны «Центр Красок»', 'Руководитель подразделения', 'Координация с внутренней командой'],
    ['По финансовым вопросам', 'Менеджер по работе с ключевыми клиентами «Эмпилс»', 'Бонусы, отсрочка, расчёты'],
    ['По обучению', 'Тренер «Эмпилс»', 'Проведение обучения, тестирование'],
]
make_table(doc, ['Сторона', 'Ответственный', 'Зона ответственности'], contacts_data,
           col_widths=[4, 5, 7])

add_para(doc, 'Порядок взаимодействия:', bold=True, color=P['primary'])
add_bullet(doc, 'Ежемесячные созвоны (онлайн) для контроля выполнения программы')
add_bullet(doc, 'Ежеквартальные очные встречи для обсуждения результатов')
add_bullet(doc, 'Оперативные вопросы решаются по телефону/email в течение 1–2 рабочих дней')
add_bullet(doc, 'Эскалация вопросов — на уровне руководителей подразделений')

doc.add_page_break()

# ===== 9. ЗАКЛЮЧИТЕЛЬНЫЕ ПОЛОЖЕНИЯ =====
add_heading(doc, '9. Заключительные положения', level=1)

add_para(doc, '9.1. Программа вступает в силу с 6 июля 2026 года и действует до 31 декабря 2026 года.',
         indent=True)

add_para(doc, '9.2. По истечении срока действия программы стороны проводят совместный анализ '
              'результатов и принимают решение о продлении, изменении или прекращении программы.',
         indent=True)

add_para(doc, '9.3. Все изменения и дополнения к программе оформляются в письменном виде и '
              'подписываются обеими сторонами.', indent=True)

add_para(doc, '9.4. Программа не является публичной офертой. Условия могут быть скорректированы '
              'по согласованию сторон.', indent=True)

add_para(doc, '9.5. Все споры и разногласия разрешаются путём переговоров. При недостижении '
              'согласия — в соответствии с действующим законодательством РФ.', indent=True)

# Signatures section
doc.add_paragraph()
add_heading(doc, 'Подписи сторон', level=2, color=P['dark'])

sig_table = doc.add_table(rows=4, cols=2)
sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER

# Headers
for i, header in enumerate(['Производитель: АО «Эмпилс»', 'Дистрибьютор: «Центр Красок»']):
    cell = sig_table.rows[0].cells[i]
    cell.text = ''
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(header)
    run.font.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = P['primary']
    set_cell_bg(cell, 'F5F5F5')
    set_cell_borders(cell)

# Empty rows for signatures
for row_idx in [1, 2]:
    for col_idx in [0, 1]:
        cell = sig_table.rows[row_idx].cells[col_idx]
        cell.text = ''
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(30)
        p.paragraph_format.space_after = Pt(30)
        run = p.add_run('')
        set_cell_borders(cell)

# Signature lines
for i, label in enumerate(['_______________________ / ____________', '_______________________ / ____________']):
    cell = sig_table.rows[3].cells[i]
    cell.text = ''
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(label)
    run.font.size = Pt(10)
    run.font.color.rgb = P['secondary']
    set_cell_borders(cell)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(5)
run = p.add_run('М.П.                                                        М.П.')
run.font.size = Pt(10)
run.font.color.rgb = P['secondary']

# Date
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(20)
run = p.add_run('Дата подписания: «____» ____________ 2026 г.')
run.font.size = Pt(11)
run.font.color.rgb = P['body']

# Final note
doc.add_paragraph()
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(20)
run = p.add_run('Empils • TITAN • Программа партнёрского сотрудничества • 2026')
run.font.size = Pt(10)
run.font.italic = True
run.font.color.rgb = P['secondary']

# Save
doc.save(OUT_PATH)
print(f'Document saved: {OUT_PATH}')
print(f'Size: {os.path.getsize(OUT_PATH)} bytes')
