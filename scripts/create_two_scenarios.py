#!/usr/bin/env python3
"""
Generate a document with TWO matrices:
A) For NEW distributors (don't sell TITAN yet) - convince to TRY
B) For EXISTING distributors (already sell TITAN) - INCREASE sales
Based on the matrix format from the original program.
"""

import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT_PATH = '/home/z/my-project/download/Матрицы_инструментов_ТITAN_два_сценария.docx'

P = {
    'primary': RGBColor(0xC6, 0x28, 0x28),
    'dark': RGBColor(0x1A, 0x1F, 0x36),
    'body': RGBColor(0x2D, 0x31, 0x42),
    'secondary': RGBColor(0x5A, 0x60, 0x80),
    'blue': RGBColor(0x25, 0x63, 0xEB),
    'green': RGBColor(0x16, 0xA3, 0x4A),
    'orange': RGBColor(0xEA, 0x58, 0x0C),
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


def add_h(doc, text, level=1, color=None):
    if color is None:
        color = P['primary'] if level == 1 else P['dark']
    h = doc.add_heading('', level=level)
    run = h.add_run(text)
    run.font.color.rgb = color
    run.font.name = 'Calibri'


def add_p(doc, text, size=11, bold=False, italic=False, color=None, indent=False):
    p = doc.add_paragraph()
    if indent:
        p.paragraph_format.first_line_indent = Cm(0.5)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.25
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = 'Calibri'
    if color:
        run.font.color.rgb = color
    return p


def add_b(doc, text):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.2
    run = p.add_run(text)
    run.font.size = Pt(10)
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
        run.font.size = Pt(10)
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
            p.paragraph_format.line_spacing = 1.15
            run = p.add_run(str(cell_text))
            run.font.size = Pt(9)
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
style.paragraph_format.line_spacing = 1.25

section = doc.sections[0]
section.orientation = WD_ORIENT.LANDSCAPE
section.page_width = Cm(29.7)
section.page_height = Cm(21.0)
section.left_margin = Cm(2.0)
section.right_margin = Cm(2.0)
section.top_margin = Cm(1.5)
section.bottom_margin = Cm(1.5)

# ===== COVER =====
cover = doc.add_table(rows=1, cols=1)
cell = cover.rows[0].cells[0]
set_cell_bg(cell, 'C62828')
cell.text = ''
p = cell.paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(35)
p.paragraph_format.space_after = Pt(10)
run = p.add_run('TITAN | Эмпилс')
run.font.size = Pt(36)
run.font.bold = True
run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
p2 = cell.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
p2.paragraph_format.space_after = Pt(25)
run = p2.add_run('Матрицы инструментов TITAN 3в1\nДва сценария работы с дистрибьюторами')
run.font.size = Pt(22)
run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
p3 = cell.add_paragraph()
p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
p3.paragraph_format.space_after = Pt(35)
run = p3.add_run('Сценарий A: Новый партнёр (убедить попробовать)\nСценарий Б: Действующий партнёр (увеличить продажи)')
run.font.size = Pt(14)
run.font.italic = True
run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

doc.add_paragraph()
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Empils • TITAN • 2026')
run.font.size = Pt(13)
run.font.color.rgb = P['secondary']

doc.add_page_break()

# ===== 1. ВВЕДЕНИЕ =====
add_h(doc, '1. Логика двух сценариев', level=1)

add_p(doc,
    'Не все инструменты применимы ко всем дистрибьюторам. Дистрибьютор, который '
    'впервые слышит о TITAN 3в1, не оценит «квартальный конкурс ТП» — у него '
    'ещё нет ТП, работающих с TITAN. И наоборот, дистрибьютор, который уже продаёт '
    'TITAN 5 лет, не нуждается в «бонусе за первичную закупку» — он уже закупает.',
    indent=True)

add_p(doc,
    'Поэтому инструменты разделены на два сценария. В каждом сценарии выбраны '
    'только те практики, которые наиболее эффективны для конкретной цели — '
    'убедить попробовать (Сценарий А) или увеличить продажи (Сценарий Б).',
    indent=True)

# Two scenarios summary table
scen_data = [
    ['A', 'Новый партнёр', 'Дистрибьютор НЕ продаёт TITAN 3в1. У него уже есть ПФ-115, Hammerite или другой аналог.',
     'Убедить попробовать: снизить риск, снять страх, показать выгоду', 'Оранжевый'],
    ['Б', 'Действующий партнёр', 'Дистрибьютор УЖЕ продаёт TITAN 3в1, но продажи нестабильны или малы.',
     'Увеличить продажи: масштабировать, мотивировать, углублять', 'Зелёный'],
]
make_table(doc, ['Сценарий', 'Тип', 'Описание', 'Цель', 'Цвет'],
           scen_data, col_widths=[1.5, 3, 8, 8, 2])

add_p(doc, '')
add_p(doc,
    'Принцип отбора: из 130+ инструментов (собранных в трёх отдельных файлах по '
    'первичным, вторичным и третичным продажам) выбраны наиболее эффективные '
    'для каждой цели. Не все — только то, что даёт максимальный эффект при '
    'разумных затратах.', italic=True, color=P['secondary'])

doc.add_page_break()

# ===== 2. СЦЕНАРИЙ A =====
add_h(doc, '2. Сценарий А. Новый партнёр: «Убедить попробовать»', level=1, color=P['orange'])

add_p(doc,
    'Дистрибьютор не хочет заводить TITAN, потому что у него уже есть аналог '
    '(ПФ-115, Hammerite, СТМ сети). Главный барьер — страх: «Закуплю — не продам, '
    'заморожу деньги, потеряю время». Инструменты этого сценария снимают страх '
    'и делают первую закупку максимально безболезненной.', indent=True)

add_p(doc,
    'Принцип: «Попробуй без риска». Дистрибьютор не платит за ошибку. Если TITAN '
    'не пойдёт — он ничего не теряет. Если пойдёт — зарабатывает больше, чем на ПФ-115.',
    bold=True, color=P['orange'])

# Matrix A
matrix_a = [
    ['1', 'Бонус за первичную закупку 5–10% товаром', '—', '—', 'Снизить порог входа: первая закупка без финансовых потерь'],
    ['1', 'Возврат неликвида 90 дней', 'Минимальный заказ от 1 банки для розницы', 'Гарантия возврата 30 дней для покупателя', 'Снять страх «куплю и не продам» на всех уровнях'],
    ['1', 'Отсрочка 60 дней', 'Отсрочка 14–30 дней для розницы', '—', 'Снизить финансовую нагрузку на старте'],
    ['1', 'Бесплатные POS-материалы + стартовый набор', 'Стартовый набор TITAN (выкрасы + шпаргалка + QR-стенд) бесплатно', 'Выкрасы «до/после» на полку бесплатно', 'Полка готова к продажам без затрат партнёра'],
    ['1', 'Выездное обучение команды (1 день)', 'Выездное обучение продавцов в магазине (45 мин)', 'Скрипт «2 минуты ТСО» + шпаргалка продавца', 'Команда умеет продавать с первого дня'],
    ['1', 'Экскурсия на завод (делегация 5–7 чел.)', '—', '—', 'Показать производство → лояльность руководства'],
    ['1', 'Аргумент: сокращение SKU на 5–7 позиций', 'Аргумент: выше маржа 30–50% на банку', 'Аргумент: «1 банка вместо 3»', 'Показать прямую выгоду без бонусов'],
    ['—', '—', 'ТП проводит презентацию закупщику розницы (15 мин)', 'Воблер «1 банка вместо 2» на полку', 'Розница видит выгоду и берёт TITAN на пробу'],
    ['—', '—', 'Консигнация для топ-5 точек дистрибьютора', '—', 'Товар стоит на полке без оплаты — риск = 0'],
    ['—', '—', '—', 'QR-код на банке → видео 2 мин', 'Покупатель сам узнаёт о TITAN, без продавца'],
]
make_table(doc, ['Этап', 'Первичные (Завод→Дистр.)', 'Вторичные (Дистр.→Розница)', 'Третичные (Розница→Покупатель)', 'Цель'],
           matrix_a, col_widths=[1.2, 6, 6, 5.5, 5], header_color='ea580c')

add_p(doc, '')
add_p(doc, 'Ключевые принципы Сценария А:', bold=True, color=P['orange'])
add_b(doc, 'Минимум финансовых обязательств дистрибьютора — всё на риске Эмпилс')
add_b(doc, 'Обучение «в упаковке» — команда умеет продавать с первого дня')
add_b(doc, 'Аргумент «сокращение SKU + рост маржи» — главный козырь для убеждения')
add_b(doc, 'Возврат неликвида + гарантия для покупателя — снимают страх на всех уровнях')
add_b(doc, 'Цель первого этапа: не объём, а ПРОБА. Пусть купят 5 банок — но продадут и вернутся')

doc.add_page_break()

# ===== 3. СЦЕНАРИЙ Б =====
add_h(doc, '3. Сценарий Б. Действующий партнёр: «Увеличить продажи»', level=1, color=P['green'])

add_p(doc,
    'Дистрибьютор уже продаёт TITAN 3в1, но продажи не растут или нестабильны. '
    'Он знает продукт, но не вкладывается в его продвижение. Причины: нет мотивации '
    'ТП, нет инструментов для розницы, нет обучения продавцов, нет системного подхода.',
    indent=True)

add_p(doc,
    'Принцип: «Заработай больше, инвестируя в рост». Дистрибьютор уже доверяет '
    'продукту. Теперь нужно дать ему инструменты, чтобы продавать больше — и '
    'вознаградить за рост.', bold=True, color=P['green'])

# Matrix B
matrix_b = [
    ['1', 'Ретробонус 8% ежемесячно + 3% полугодовой', '—', '—', 'Регулярная мотивация закупать больше'],
    ['1', 'Прогрессивная объёмная скидка 3/5/7%', '—', '—', 'Чем больше закупает — тем выгоднее'],
    ['1', 'Co-op маркетинг 50% (совместное промо)', 'Субсидирование промо-акций для розницы', 'Сезонная акция «Весна/Осень» (скидка 5–7%)', 'Совместные инвестиции в рост продаж'],
    ['1', 'Ценовая защита 3–6 месяцев', '—', '—', 'Гарантия маржи — дистрибьютор планирует рост'],
    ['1', 'Регулярное обучение + онлайн-платформа', 'Сертификация ТП по ТСО (бонус 3000 руб.)', 'Сертификация «Продавец ТСО» (+3000/−3000 руб.)', 'Превратить знание в навык на всех уровнях'],
    ['—', '—', 'Бонус за новую точку 500–1000 руб.', '—', 'ТП активно расширяют сеть точек с TITAN'],
    ['—', '—', 'Ретробонус 2–3% для розницы за оборот', 'Бонус продавцу 50–100 руб. за банку', 'Финансовая мотивация розницы и продавцов'],
    ['—', '—', 'Квартальный конкурс ТП (топ-3: 10/7/5 тыс.)', 'Конкурс «Продавец месяца» (5/3/2 тыс. + бейдж)', 'Соревнование → эмоции → рост продаж'],
    ['—', '—', 'Калькулятор ТСО для ТП (мобильный)', 'Калькулятор ТСО для покупателя (онлайн + QR)', 'ТСО-инструмент на двух уровнях'],
    ['—', '—', 'Набор выкрасов для ТП (портативный)', 'Выкрасы «до/после» + «через год» на полке', 'Наглядность = конверсия'],
    ['—', '—', 'Сравнительная таблица TITAN vs конкуренты (А4)', 'Шпаргалка продавца А6 (ламинированная)', 'Стандарт качества работы с полкой'],
    ['—', '—', 'Программа для розницы: скидка за 3+ цвета RAL', 'Подарок при покупке (кисть/валик)', 'Стимулирование ассортимента и покупки'],
    ['—', '—', 'CRM-интеграция + аналитика продаж', 'Отчётность по сканированиям QR', 'Данные для принятия решений'],
    ['—', '—', '—', 'Гарантия 10 лет + сертификаты на банке', 'Доверие покупателя → конверсия → повторные покупки'],
    ['—', '—', '—', 'Реферальная программа «Приведи соседа»', 'Сарафанное радио → новые покупатели бесплатно'],
]
make_table(doc, ['Этап', 'Первичные (Завод→Дистр.)', 'Вторичные (Дистр.→Розница)', 'Третичные (Розница→Покупатель)', 'Цель'],
           matrix_b, col_widths=[1.2, 6, 6, 5.5, 5], header_color='16a34a')

add_p(doc, '')
add_p(doc, 'Ключевые принципы Сценария Б:', bold=True, color=P['green'])
add_b(doc, 'Финансовая мотивация на всех уровнях: ретробонус → бонус ТП → бонус продавца')
add_b(doc, 'Конкурсы и соревнования — эмоции работают лучше, чем просто деньги')
add_b(doc, 'Калькулятор ТСО — центральный инструмент, работает на двух уровнях (ТП + покупатель)')
add_b(doc, 'Сертификация — стандарт качества, который отличает TITAN от конкурентов')
add_b(doc, 'Аналитика — данные показывают, что работает, а что нет. Без данных — слепота')

doc.add_page_break()

# ===== 4. СРАВНЕНИЕ =====
add_h(doc, '4. Сравнение сценариев: что общего и чем отличаются', level=1)

compare_data = [
    ['Цель', 'Попробовать без риска', 'Увеличить продажи системно'],
    ['Главный принцип', '«Сними страх»', '«Заработай больше»'],
    ['Финансовые инструменты', 'Бонус за первую закупку, возврат неликвида, отсрочка', 'Ретробонус 8%+3%, объёмные скидки, co-op маркетинг'],
    ['Риск для Эмпилс', 'Высокий (бесплатные материалы, возвраты)', 'Средний (бонусы выплачиваются из маржи)'],
    ['Риск для дистрибьютора', 'Минимальный (всё за счёт Эмпилс)', 'Средний (инвестирует своё время и полку)'],
    ['Обучение', 'Разовое выездное (1 день)', 'Системное: онлайн-платформа + сертификация'],
    ['Мотивация ТП', 'Базовая (зарплата)', 'Бонусы + конкурсы + сертификация'],
    ['Мотивация продавцов', 'Шпаргалка + скрипт', 'Бонус за банку + конкурс + сертификация'],
    ['POS-материалы', 'Стартовый набор (бесплатно)', 'Полный пакет + обновления'],
    ['Калькулятор ТСО', 'Не нужен на старте', 'Ключевой инструмент роста'],
    ['Срок до результата', '1–2 месяца (первая продажа)', '3–6 месяцев (рост +20–30%)'],
    ['Переход в Сценарий Б', 'После 3 месяцев стабильных продаж', '—'],
]
make_table(doc, ['Параметр', 'Сценарий А (Новый)', 'Сценарий Б (Действующий)'],
           compare_data, col_widths=[5, 8, 8], header_color='1A1F36')

add_p(doc, '')
add_p(doc,
    'Логика перехода: дистрибьютор начинает в Сценарии А (проба). После 2–3 месяцев '
    'стабильных продаж переходит в Сценарий Б (рост). Сценарий А — это «турбостарт», '
    'Сценарий Б — «крейсерский режим». Нельзя начинать с Б — дистрибьютор, который '
    'не пробовал TITAN, не оценит бонусы за рост. Нельзя оставаться в А forever — '
    'бесплатные материалы и возвраты не вечны.', indent=True, bold=True)

doc.add_page_break()

# ===== 5. КРИТЕРИИ ВЫБОРА =====
add_h(doc, '5. Как определить, какой сценарий применять', level=1)

add_p(doc, 'Решающее дерево для каждого дистрибьютора:', indent=True)

tree_data = [
    ['1', 'Закупал ли дистрибьютор TITAN 3в1 за последние 6 месяцев?', 'Да → Сценарий Б. Нет → шаг 2', '—'],
    ['2', 'Есть ли TITAN 3в1 на складе дистрибьютора прямо сейчас?', 'Да → Сценарий Б (поддержка). Нет → шаг 3', '—'],
    ['3', 'Знает ли дистрибьютор о TITAN 3в1?', 'Нет → Сценарий А (с нуля). Да → шаг 4', '—'],
    ['4', 'Отказался ли ранее от TITAN?', 'Да → Сценарий А (повторная проба с новыми условиями). Нет → шаг 5', '—'],
    ['5', 'Есть ли у дистрибьютора аналог (Hammerite, Dulux)?', 'Да → Сценарий А (аргумент сокращения SKU). Нет → Сценарий А (чистый вход)', '—'],
]
make_table(doc, ['Шаг', 'Вопрос', 'Решение', 'Примечание'],
           tree_data, col_widths=[1.5, 8, 7, 5.5], header_color='5A6080')

add_p(doc, '')
add_h(doc, '5.1. Дополнительно: размер дистрибьютора', level=2, color=P['dark'])

add_p(doc,
    'Размер дистрибьютора определяет не сценарий, а набор инструментов внутри сценария. '
    'Малому дистрибьютору не нужен «территориальный эксклюзив», а крупному — «минимальный '
    'заказ от 1 банки».', indent=True)

size_data = [
    ['Малый (< 10 млн/год)', 'Стартовый набор, возврат неликвида, обучение 1 день', 'Бонус за оборот, конкурс продавцов, QR-стенд'],
    ['Средний (10–50 млн/год)', '+ отсрочка 60 дней, экскурсия на завод, co-op маркетинг', '+ ретробонус 8%+3%, объёмные скидки, онлайн-платформа'],
    ['Крупный (> 50 млн/год)', '+ территориальный эксклюзив, ценовая защита, консигнация', '+ CRM-интеграция, гарантия маржи, стратегическое партнёрство'],
]
make_table(doc, ['Размер', 'Дополнительно для Сценария А', 'Дополнительно для Сценария Б'],
           size_data, col_widths=[5, 8, 8], header_color='1A1F36')

doc.add_page_break()

# ===== 6. ПИСЬМО КОЛЛЕГАМ =====
add_h(doc, '6. Письмо коллегам для согласования', level=1)

add_p(doc, 'Ниже — проект письма региональным менеджерам для согласования программы:',
      italic=True, color=P['secondary'])

# Letter in a styled box
letter_table = doc.add_table(rows=1, cols=1)
letter_cell = letter_table.rows[0].cells[0]
set_cell_bg(letter_cell, 'F8F9FA')
set_cell_borders(letter_cell, 'CCCCCC', '6')
letter_cell.text = ''

p = letter_cell.paragraphs[0]
p.paragraph_format.space_after = Pt(8)
run = p.add_run('Коллеги, добрый день!')
run.font.size = Pt(12)
run.font.bold = True

letter_text = (
    'Разрабатываю комплексную программу продвижения TITAN 3в1 через методологию '
    'ТСО (совокупная стоимость владения). Программа охватывает три этапа: '
    'первичные (завод → дистрибьютор), вторичные (дистрибьютор → розница) и '
    'третичные (розница → покупатель) продажи.\n\n'
    'Подготовил файлы с инструментами мотивации по каждому этапу. Внутри — '
    'практики, применимость по типам партнёров, матрицы инструментов.\n\n'
    'Задача: на основе этих инструментов собрать пакет мер для каждого '
    'дистрибьютора из проекта «Белые пятна» — как новых (кому предлагаем TITAN '
    'впервые), так и текущих (у кого хотим увеличить продажи).\n\n'
    'Подход — два сценария:\n'
    '• Сценарий А «Новый партнёр» — убедить попробовать без риска\n'
    '• Сценарий Б «Действующий партнёр» — увеличить продажи системно\n\n'
    'Для каждого сценария — свой набор инструментов по трём этапам.\n\n'
    'Прошу:\n'
    '1. Ознакомиться с файлами (вложения)\n'
    '2. Подсказать, какие инструменты реально применимы в ваших регионах\n'
    '3. Указать дистрибьюторов, с которых стоит начать\n\n'
    'Жду обратную связь до [дата].'
)

for para_text in letter_text.split('\n\n'):
    p = letter_cell.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.3
    run = p.add_run(para_text)
    run.font.size = Pt(11)
    run.font.name = 'Calibri'

# Final note
doc.add_paragraph()
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(15)
run = p.add_run('Empils • TITAN • Матрицы инструментов • 2026')
run.font.size = Pt(10)
run.font.italic = True
run.font.color.rgb = P['secondary']

# Save
doc.save(OUT_PATH)
print(f'Document saved: {OUT_PATH}')
print(f'Size: {os.path.getsize(OUT_PATH)} bytes')
