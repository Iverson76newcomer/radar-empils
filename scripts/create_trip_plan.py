#!/usr/bin/env python3
"""
Word — план командировки по белым пятнам TITAN (1 страница А4).
7 партнёров: Олимп, НИССА, Стройторг, Трианон, Гектор, Alterra, + (все в Сибири).
Цель визита: первичный контакт. Дальнейшие шаги.
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = '/home/z/my-project/download/План_командировки_белые_пятна_TITAN.docx'

COL_PRIMARY = RGBColor(0x1E, 0x3A, 0x5F)
COL_ACCENT = RGBColor(0xF4, 0x79, 0x20)
COL_BODY = RGBColor(0x1A, 0x1F, 0x36)
COL_GRAY = RGBColor(0x6B, 0x72, 0x80)
COL_RED = RGBColor(0xC6, 0x28, 0x28)
COL_GREEN = RGBColor(0x16, 0xA3, 0x4A)


def add_para(doc, text='', size=10, bold=False, color=COL_BODY, align=None,
             font='Arial', space_before=0, space_after=2, line_spacing=1.15,
             indent_left=0):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    if align is not None:
        p.alignment = align
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = line_spacing
    if indent_left:
        pf.left_indent = Cm(indent_left)
    if text:
        run = p.add_run(text)
        run.font.name = font
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.append(rFonts)
        rFonts.set(qn('w:ascii'), font)
        rFonts.set(qn('w:hAnsi'), font)
        rFonts.set(qn('w:cs'), font)
    return p


def set_cell_shading(cell, hex_color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tc_pr.append(shd)


def set_cell_borders(cell, color='D1D5DB', sz='4'):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        b = OxmlElement(f'w:{edge}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), sz)
        b.set(qn('w:color'), color)
        tc_borders.append(b)
    tc_pr.append(tc_borders)


def style_cell(cell, value, font_size=9, bold=False, color=COL_BODY, fill=None, align='left', valign='top'):
    if fill:
        set_cell_shading(cell, fill)
    set_cell_borders(cell)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP if valign == 'top' else WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.line_spacing = 1.05
    if align == 'center':
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == 'right':
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    # Очищаем и добавляем run
    run = p.add_run(value)
    run.font.name = 'Arial'
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), 'Arial')
    rFonts.set(qn('w:hAnsi'), 'Arial')


# =========================================================
# СОЗДАНИЕ ДОКУМЕНТА
# =========================================================
doc = Document()

# Поля А4 — компактные, чтобы всё вошло на 1 страницу
for section in doc.sections:
    section.top_margin = Cm(1.0)
    section.bottom_margin = Cm(1.0)
    section.left_margin = Cm(1.3)
    section.right_margin = Cm(1.3)

style = doc.styles['Normal']
style.font.name = 'Arial'
style.font.size = Pt(10)
style.font.color.rgb = COL_BODY

# === ЗАГОЛОВОК ===
add_para(doc, 'ПЛАН КОМАНДИРОВКИ — БЕЛЫЕ ПЯТНА TITAN', size=9, bold=True, color=COL_ACCENT,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=0)
add_para(doc, 'Сибирский ФО: Барнаул + Бийск + Новосибирск', size=14, bold=True, color=COL_PRIMARY,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=1)
add_para(doc, 'Цель визита: первичный контакт с 7 потенциальными партнёрами (Барнаул — 3, Бийск — 1, Новосибирск — 3). Дальнейшие шаги по каждому — в таблице ниже.',
         size=9, color=COL_GRAY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=3)

# === ТАБЛИЦА С ПАРТНЁРАМИ ===
# 6 колонок: № / Партнёр / Город / Текущий статус / Цель визита / Дальнейшие шаги
headers = ['№', 'Партнёр / ЛПР', 'Город', 'Текущий статус', 'Цель визита', 'Дальнейшие шаги']
partners = [
    {
        'n': '1',
        'name': 'Олимп\nАлексей Анатольевич\n+7 913 213-95-40\nolimp22.ru',
        'city': 'Барнаул',
        'status': 'Взял паузу. Отходит от ЛКМ — развивает поставки электроинструмента из Китая. Выручка 495 → 442 млн (−11%).',
        'goal': 'Понять, окончательно ли ушёл из ЛКМ. Если нет — предложить TITAN 3в1 как маржинальный продукт без конкуренции с их текущим ассортиментом (Текс, Новоколор).',
        'next': 'Если вернётся в ЛКМ — КП на 3в1+молотковая (5 SKU). Если ушёл — фиксируем «на удержание», визит раз в полгода.',
    },
    {
        'n': '2',
        'name': 'НИССА\nАлександр\n+7 983 106-85-65\nnissa-biysk.net',
        'city': 'Бийск',
        'status': 'Договорённость на 2 контейнера Простокрашено+Profiwood, но переносит сроки. Демпингуют Лакру. Выручка 4 млрд.',
        'goal': 'Зафиксировать конкретные сроки поставки 2 контейнеров. Понять причину переносов — логистика, цена, или ЛПР не принимает решение.',
        'next': 'Подписать договор на 2 контейнера до 31.08.2026. Если опять перенос — эскалация на гендиректора НИССА.',
    },
    {
        'n': '3',
        'name': 'Стройторг\nХалилов Станислав Анатольевич\n+7 913 780-53-53\nstroytorgmoda.ru',
        'city': 'Новосибирск',
        'status': 'КП рассмотрел, но «не готовы заводить нового поставщика». С 16.07.2025 не выходит на связь. Выручка 621 млн.',
        'goal': 'Восстановить контакт. Понять реальные возражения — возможно, дело в демпинге конкурентов или в логистике.',
        'next': 'Назначить повторную встречу с презентацией TITAN через выкрас. Если откажется — передать в отдел трейд-маркетинга на акцию возврата.',
    },
    {
        'n': '4',
        'name': 'ТД Трианон\nГурин Константин Викторович\n+7 383 255-11-53\ntrianon-nsk.ru',
        'city': 'Новосибирск',
        'status': 'Встреча 28.07 в 11:00 подтверждена. КП рассмотрел. Возражение: «не узнаваемый бренд в РТТ, риск нового контракта на падающем рынке».',
        'goal': 'Закрыть возражение «не узнаваемый бренд»: показать матрицу TCO, выкрас TITAN vs ПФ-115, кейсы реальных объектов.',
        'next': 'Подписать тестовый контракт на 5 SKU TITAN (3в1 + молотковая) с маркетинговым бюджетом на 3 месяца. Повторная встреча через 6 недель.',
    },
    {
        'n': '5',
        'name': 'Гектор (ООО Арес групп)\nЛПР — поиск (сменился управл.)\n+7 923 225-51-15\ngektor-nsk.ru',
        'city': 'Новосибирск',
        'status': 'Предварительно 28.07. Сменился управляющий департамента, ищем нового ЛПР. Возражение: нужна логистика — склад в НСК с подгрузкой, акции демпинга.',
        'goal': 'Найти нового ЛПР. Предложить региональный склад Эмпилс в НСК как решение логистического возражения.',
        'next': 'После встречи — в течение 2 недель подтвердить складскую схему с логистикой Эмпилс. Контракт на 7 SKU TITAN + Profiwood.',
    },
    {
        'n': '6',
        'name': 'Alterra (ООО Альтерра)\nДиплинская Валентина\n+7 906 964-71-46\ntg-alterra.ru',
        'city': 'Барнаул',
        'status': 'КП рассмотрел, новых поставщиков на 2026 не рассматривают. Сокращение ассортимента, работают с Лакра+Рогнеда. Сеть 16 магазинов. Выручка 4.9 млрд.',
        'goal': 'Договориться о тесте в 1-2 магазинах из 16. Не просить всю сеть — только пилот, чтобы минимизировать их риск.',
        'next': 'Пилот в 2 магазинах: 3в1 + молотковая (5 SKU). Через 8 недель — review продаж, расширение на 4-8 магазинов.',
    },
    {
        'n': '7',
        'name': 'Хелми\nЛПР — подтверждается\n+7 3852 50-50-70\nhelmi-opt.ru',
        'city': 'Барнаул',
        'status': 'Отказ от встречи и рассмотрения КП на 2026 — «нас знают». Не рассматривали в июне новых поставщиков. Выручка 589 млн. Ассортимент: Колорит, Текс.',
        'goal': 'Изменить позицию «нас знают» на «видите выгоду». Показать TITAN через TCO — маржа выше, возвраты ниже, чем по Колорит/Текс. Предварительно 29.07.',
        'next': 'Если согласятся на пилот — 5 SKU TITAN в 1-2 точках. Если категорический отказ — фиксируем «отказ 2026», возврат через 6 месяцев с новым КП и кейсами.',
    },
]

# Создаём таблицу
table = doc.add_table(rows=1 + len(partners), cols=6)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
table.autofit = False

# Ширины колонок (сумма ~ 18 см при полях 1.5+1.5 на А4 21 см)
col_widths = [Cm(0.7), Cm(3.5), Cm(1.5), Cm(4.3), Cm(4.0), Cm(4.0)]
for i, w in enumerate(col_widths):
    table.columns[i].width = w

# Шапка
header_fill = '1E3A5F'
for i, h in enumerate(headers):
    cell = table.cell(0, i)
    cell.width = col_widths[i]
    style_cell(cell, h, font_size=8, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF), fill=header_fill, align='center', valign='center')

# Данные
for ri, p in enumerate(partners, 1):
    fill = 'FFFFFF' if ri % 2 == 1 else 'F9FAFB'
    # №
    style_cell(table.cell(ri, 0), p['n'], font_size=9, bold=True, color=COL_PRIMARY, fill=fill, align='center', valign='center')
    # Партнёр
    style_cell(table.cell(ri, 1), p['name'], font_size=7, bold=False, color=COL_BODY, fill=fill, align='left')
    # Город
    style_cell(table.cell(ri, 2), p['city'], font_size=8, bold=True, color=COL_PRIMARY, fill=fill, align='center', valign='center')
    # Статус
    style_cell(table.cell(ri, 3), p['status'], font_size=7, color=COL_BODY, fill=fill, align='left')
    # Цель
    style_cell(table.cell(ri, 4), p['goal'], font_size=7, color=COL_BODY, fill=fill, align='left')
    # Дальнейшие шаги
    style_cell(table.cell(ri, 5), p['next'], font_size=7, bold=True, color=COL_PRIMARY, fill=fill, align='left')

# === ИТОГОВАЯ СВОДКА ВИЗИТА ===
add_para(doc, '', size=3, space_after=1)
add_para(doc, 'СВОДКА КОМАНДИРОВКИ', size=9, bold=True, color=COL_ACCENT, space_after=1)

# Сводная таблица: 3 колонки
summary_table = doc.add_table(rows=2, cols=3)
summary_table.alignment = WD_TABLE_ALIGNMENT.CENTER
summary_table.autofit = False
sum_widths = [Cm(6.1), Cm(6.1), Cm(6.1)]
for i, w in enumerate(sum_widths):
    summary_table.columns[i].width = w

# Шапка
sum_headers = ['География и сроки', 'Приоритеты', 'Ожидаемый результат']
for i, h in enumerate(sum_headers):
    cell = summary_table.cell(0, i)
    cell.width = sum_widths[i]
    style_cell(cell, h, font_size=8, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF), fill='F47920', align='center', valign='center')

# Данные
sum_data = [
    '3 города, 7 партнёров: Барнаул (Олимп, Alterra, Хелми) → Бийск (НИССА) → Новосибирск (Стройторг, Трианон, Гектор). Период: 28.07–01.08.2026.',
    '1) НИССА — закрыть 2 контейнера. 2) Трианон — закрыть тест. 3) Гектор — найти ЛПР. 4) Стройторг — реанимировать. 5) Alterra — пилот в 2 магазинах. 6) Хелми — сломать отказ. 7) Олимп — на удержание.',
    'Подписать 3 контракта (НИССА, Трианон, Alterra-пилот) = +2 контейнера Простокрашено+Profiwood и +12 SKU TITAN в тесте. Плюс 4 тёплых контакта на Q4 2026.',
]
for i, val in enumerate(sum_data):
    cell = summary_table.cell(1, i)
    cell.width = sum_widths[i]
    style_cell(cell, val, font_size=7, color=COL_BODY, fill='FEF3C7', align='left')

# === ЧТО ВЗЯТЬ С СОБОЙ ===
add_para(doc, '', size=3, space_after=1)
add_para(doc, 'ВЗЯТЬ НА ВСТРЕЧИ:', size=8, bold=True, color=COL_PRIMARY, space_after=1)

kit_text = (
    'КП TITAN (3 линейки: 3в1 / быстросохнущая / молотковая) • выкрасы TITAN vs ПФ-115 (2 пластины) • '
    'матрица TCO и расчёт маржинальности для дистрибьютора • КП Profiwood (для НИССА) • '
    'договоры на 2 контейнера (для НИССА — готовить к подписи) • маркетинговый бюджет на 3 месяца (для Трианон) • '
    'QR-код на видеоинструкцию покупателя (для всех)'
)
add_para(doc, kit_text, size=7, color=COL_BODY, space_after=1, line_spacing=1.1)

# Подпись
add_para(doc, 'План составлен на основе «Белые пятна_общая_таблица_03 07 2026.xlsx». РД: Захарова О.Г. (Сибирский ФО).',
         size=6, color=COL_GRAY, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=1)

doc.save(OUT)
print(f'OK: {OUT}')
import os
print(f'Размер: {os.path.getsize(OUT) / 1024:.1f} KB')

# Проверка количества страниц
print()
print("=== Проверка структуры ===")
print(f"Параграфов: {len(doc.paragraphs)}")
print(f"Таблиц: {len(doc.tables)}")
EOF