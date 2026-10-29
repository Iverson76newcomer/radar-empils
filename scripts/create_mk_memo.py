#!/usr/bin/env python3
"""
Памятка для проведения мастер-класса TITAN — 1 лист A4.
Последовательность действий + скрипт + главные тезисы.
Компактно, для печати и использования на месте.
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = '/home/z/my-project/download/Памятка_мастер-класс_TITAN_A4.docx'

RED = RGBColor(0xC6, 0x28, 0x28)
DARK = RGBColor(0x1A, 0x1F, 0x36)
GRAY = RGBColor(0x60, 0x60, 0x60)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x00, 0x00, 0x00)


def set_para(p, size=10, bold=False, color=DARK, align=WD_ALIGN_PARAGRAPH.LEFT,
             space_after=2, space_before=0, italic=False, font='Arial'):
    pf = p.paragraph_format
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(space_before)
    pf.line_spacing = 1.05
    p.alignment = align
    for run in p.runs:
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = color
        run.font.name = font


def add_para(doc, text='', size=10, bold=False, color=DARK,
             align=WD_ALIGN_PARAGRAPH.LEFT, space_after=2, space_before=0,
             italic=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(space_before)
    pf.line_spacing = 1.05
    p.alignment = align
    if text:
        run = p.add_run(text)
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = color
        run.font.name = 'Arial'
    return p


def add_cell_text(cell, text, size=9, bold=False, color=DARK, align=WD_ALIGN_PARAGRAPH.LEFT):
    cell.text = ''
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.line_spacing = 1.1
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = 'Arial'


def shade_cell(cell, hex_color):
    shading = OxmlElement('w:shd')
    shading.set(qn('w:val'), 'clear')
    shading.set(qn('w:fill'), hex_color)
    cell._element.get_or_add_tcPr().append(shading)


def main():
    doc = Document()

    # Минимальные поля для размещения на 1 листе A4
    for section in doc.sections:
        section.top_margin = Cm(0.8)
        section.bottom_margin = Cm(0.6)
        section.left_margin = Cm(1.0)
        section.right_margin = Cm(1.0)

    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(9)

    # === ЗАГОЛОВОК ===
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run('МАСТЕР-КЛАСС TITAN 3в1 — ПАМЯТКА')
    run.font.size = Pt(13)
    run.font.bold = True
    run.font.color.rgb = RED
    run.font.name = 'Arial'

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run('«Красишь 1 раз — служит 10 лет»  •  Empils  •  Лемана ПРО  •  17-25 июля 2026')
    run.font.size = Pt(8)
    run.font.italic = True
    run.font.color.rgb = GRAY
    run.font.name = 'Arial'

    # === ТАБЛИЦА: ПОСЛЕДОВАТЕЛЬНОСТЬ ДЕЙСТВИЙ + СКРИПТ ===
    # Создаём таблицу: 2 колонки (Шаг + Что делать/говорить)
    steps = [
        ('1', 'Привлечение', 'Подойти к покупателю у полки или дождаться у стола МК.',
         '«Добрый день! Сегодня бренд TITAN проводит мастер-класс — можете попробовать краску в действии и задать вопросы производителю»'),
        ('2', 'Знакомство', 'Представиться, пригласить к столу.',
         '«Меня зовут [Имя], представляю завод Эмпилс. Подойдём к столу — покажу, как работает TITAN 3в1»'),
        ('3', 'Цель покупки', 'Уточнить, что будет красить.',
         '«С какой целью планируете покупку? Что красите — забор, ворота, гараж?»'),
        ('4', 'Подготовка', 'Предложить фартук, перчатки, кисть, выбрать цвет (8017, 7024, 5005).',
         '«Попробуйте материал на пластине. Выберите цвет — графитовый, шоколадный или синий»'),
        ('5', 'Выкрашивание', 'Покупатель красит пластину. Параллельно — тезисы (см. ниже).',
         '*(презентация во время выкраса — шаги 6-9)*'),
        ('6', 'УТП: 10 лет', 'Пока красит — про срок службы.',
         '«TITAN 3в1 — грунт-эмаль по металлу. Гарантия защиты 10 лет. Покрасили 1 раз — забыли»'),
        ('7', 'УТП: 3в1', 'Не нужен отдельный грунт.',
         '«В составе: грунт + эмаль + преобразователь ржавчины. 3 продукта в одной банке. Экономия 61% vs ПФ-115 с грунтом»'),
        ('8', 'УТП: сушка', 'Акцент на времени.',
         '«Сохнет 1,5 часа на отлип. Второй слой через 5 часов. За 1 выходной — всё готово на 10 лет»'),
        ('9', 'Акция', 'Рассказать про подарок + кэшбек.',
         '«Сегодня акция: от 1 900 ₽ — кисть + перчатки в подарок, от 3 900 ₽ — сертификат 500 ₽. Плюс 5% кэшбек»'),
        ('10', 'Результат', 'Показать готовые пластины (7024, 8017).',
         '«Посмотрите — глянцевая, равномерная, не стекает, без дефектов»'),
        ('11', 'Возражения', 'Если сравнивают с Dali.',
         '«TITAN дешевле на 5% (кэшбек) + подарок 250-500 ₽. Dali выводят из Лемана ПРО. У TITAN укрывистость до 18 м²/л. У Dali 2-й слой через 2 ч или 2 суток — у TITAN через 5 ч, покрытие идеальное»'),
        ('12', 'Опрос', 'Собрать обратную связь.',
         '«Что думаете? Рассматриваете к покупке?» → записать: критерии выбора ЛКМ, барьеры к покупке → передать Зингеровской Е.Г.'),
        ('13', 'Прощание', 'Поблагодарить.',
         '«Спасибо за участие! Акция действует сегодня. Удачи!»'),
    ]

    table = doc.add_table(rows=len(steps) + 1, cols=3)
    table.style = 'Table Grid'
    table.autofit = False

    # Заголовки
    headers = ['№', 'Действие', 'Что говорить / делать']
    for ci, h in enumerate(headers):
        cell = table.cell(0, ci)
        add_cell_text(cell, h, size=7, bold=True, color=WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)
        shade_cell(cell, 'C62828')

    # Ширина колонок
    for row in table.rows:
        row.cells[0].width = Cm(0.7)
        row.cells[1].width = Cm(3.2)
        row.cells[2].width = Cm(14.0)

    # Данные
    for ri, (num, action, what_do, what_say) in enumerate(steps):
        row = table.rows[ri + 1]
        # №
        add_cell_text(row.cells[0], num, size=7, bold=True, color=RED, align=WD_ALIGN_PARAGRAPH.CENTER)
        # Действие
        add_cell_text(row.cells[1], action, size=7, bold=True, color=DARK)
        # Что говорить
        # Объединяем что делать + что говорить
        p = row.cells[2].paragraphs[0]
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.1
        run1 = p.add_run(what_do + '\n')
        run1.font.size = Pt(7)
        run1.font.color.rgb = GRAY
        run1.font.name = 'Arial'
        run2 = p.add_run(what_say)
        run2.font.size = Pt(7)
        run2.font.color.rgb = DARK
        run2.font.bold = False
        run2.font.name = 'Arial'
        run2.font.italic = True

    # === ГЛАВНЫЕ ТЕЗИСЫ ===
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run('ГЛАВНЫЕ ТЕЗИСЫ (для запоминания)')
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = RED
    run.font.name = 'Arial'

    theses = [
        ('1.', '«Красишь 1 раз — служит 10 лет»'),
        ('2.', '3в1 — грунт + эмаль + преобразователь ржавчины (не нужен отдельный грунт)'),
        ('3.', 'Экономия 61% vs ПФ-115 с грунтом ГФ-021'),
        ('4.', 'Быстро сохнет — 1,5 часа на отлип, 2-й слой через 5 часов'),
        ('5.', 'Акция — от 1 900 ₽ подарок (кисть + перчатки), от 3 900 ₽ сертификат 500 ₽'),
        ('6.', '5% кэшбек'),
        ('7.', 'Dali выводят из Лемана ПРО — TITAN останется'),
    ]

    for num, text in theses:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.1
        p.paragraph_format.left_indent = Cm(0.5)
        run1 = p.add_run(num + '  ')
        run1.font.size = Pt(8)
        run1.font.bold = True
        run1.font.color.rgb = RED
        run1.font.name = 'Arial'
        run2 = p.add_run(text)
        run2.font.size = Pt(8)
        run2.font.color.rgb = DARK
        run2.font.name = 'Arial'

    # === ЧЕК-ЛИСТ МАТЕРИАЛОВ ===
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run('ЧЕК-ЛИСТ МАТЕРИАЛОВ НА СТОЛЕ')
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = RED
    run.font.name = 'Arial'

    materials = [
        'Банки 0,5 л: RAL 8017 (шоколад), RAL 7024 (графит), RAL 5005 (синий) — 3 шт',
        'Пластинки неокрашенные 15×7 см — 25 шт',
        'Готовые пластины RAL 7024 и RAL 8017 — по 2 шт',
        'Кисти, фартуки одноразовые, перчатки',
        'Сухие + влажные салфетки — по 1 упаковке',
        'Скатерть на стол, мешки для мусора, шпатель для банок',
        'Внешний вид: футболка Эмпилс + тёмные джинсы + закрытая обувь',
    ]

    for m in materials:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.1
        p.paragraph_format.left_indent = Cm(0.5)
        run1 = p.add_run('☐  ')
        run1.font.size = Pt(8)
        run1.font.color.rgb = RED
        run1.font.name = 'Arial'
        run2 = p.add_run(m)
        run2.font.size = Pt(7)
        run2.font.color.rgb = DARK
        run2.font.name = 'Arial'

    # === НИЖНЯЯ СТРОКА ===
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('© ЗАО «Эмпилс», 2026  •  www.titanpaints.ru  •  Ростов-на-Дону')
    run.font.size = Pt(7)
    run.font.italic = True
    run.font.color.rgb = GRAY
    run.font.name = 'Arial'

    doc.save(OUT)
    print(f'OK: {OUT}')


if __name__ == '__main__':
    main()
