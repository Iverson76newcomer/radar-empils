#!/usr/bin/env python3
"""
Word-документ: текст письма для ручной рассылки по региональным директорам TITAN.
Простой текстовый формат, готов к копированию в почтовый клиент.
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = '/home/z/my-project/download/Письмо_РД_TITAN_для_почты.docx'

COL_PRIMARY = RGBColor(0x1E, 0x3A, 0x5F)
COL_ACCENT = RGBColor(0xF4, 0x79, 0x20)
COL_BODY = RGBColor(0x1A, 0x1F, 0x36)
COL_GRAY = RGBColor(0x6B, 0x72, 0x80)
COL_RED = RGBColor(0xC6, 0x28, 0x28)
COL_GREEN = RGBColor(0x16, 0xA3, 0x4A)


def add_para(doc, text='', size=11, bold=False, color=COL_BODY, align=None,
             font='Arial', space_before=0, space_after=6, line_spacing=1.35):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    if align is not None:
        p.alignment = align
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = line_spacing
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


def add_runs_para(doc, runs, size=11, align=None, space_before=0, space_after=6, line_spacing=1.35):
    """Параграф с разными форматированиями в одном тексте.
    runs = [(text, bold, color), ...]"""
    p = doc.add_paragraph()
    pf = p.paragraph_format
    if align is not None:
        p.alignment = align
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = line_spacing
    for text, bold, color in runs:
        run = p.add_run(text)
        run.font.name = 'Arial'
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.append(rFonts)
        rFonts.set(qn('w:ascii'), 'Arial')
        rFonts.set(qn('w:hAnsi'), 'Arial')
    return p


def add_bullet(doc, text, size=11, bold=False, color=COL_BODY, level=0):
    p = doc.add_paragraph(style='List Bullet')
    pf = p.paragraph_format
    pf.left_indent = Cm(0.8 + level * 0.6)
    pf.space_before = Pt(0)
    pf.space_after = Pt(3)
    pf.line_spacing = 1.3
    run = p.add_run(text)
    run.font.name = 'Arial'
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), 'Arial')
    rFonts.set(qn('w:hAnsi'), 'Arial')


def add_heading(doc, text, size=12, color=COL_PRIMARY, space_before=10, space_after=4):
    return add_para(doc, text, size=size, bold=True, color=color,
                    align=WD_ALIGN_PARAGRAPH.LEFT, space_before=space_before, space_after=space_after)


# =========================================================
# ДОКУМЕНТ
# =========================================================
doc = Document()

for section in doc.sections:
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(2.2)

style = doc.styles['Normal']
style.font.name = 'Arial'
style.font.size = Pt(11)
style.font.color.rgb = COL_BODY

# === ИНСТРУКЦИЯ ДЛЯ ВАС (вверху, удалить перед копированием) ===
add_para(doc, '↕ ТЕМА ПИСЬМА (скопируйте в поле «Тема» почтового клиента):', size=9, bold=True, color=COL_GRAY, space_after=2)
add_para(doc, 'TITAN: первичка +1.8%, вторичка −35% — 5 задач для РД на Q3 2026',
         size=10, bold=True, color=COL_PRIMARY, space_after=10)

add_para(doc, '↕ ТЕЛО ПИСЬМА (скопируйте всё, что ниже линии, в тело письма):',
         size=9, bold=True, color=COL_GRAY, space_after=2)
add_para(doc, 'Перед отправкой замените [ИМЯ РД] на имя регионального директора (например, «Ольга Геннадьевна» для Захаровой).',
         size=9, color=COL_GRAY, space_after=10)

# === ЛИНИЯ-РАЗДЕЛИТЕЛЬ ===
add_para(doc, '—' * 55, size=10, color=COL_GRAY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)

# === ТЕЛО ПИСЬМА (начало копирования) ===

# Приветствие
add_para(doc, '[ИМЯ РД], добрый день.', size=11, bold=True, color=COL_PRIMARY, space_after=8)

add_para(doc,
    'Подвели итоги по TITAN за 1 полугодие 2026. В приложении — Excel с детализацией по каждому '
    'дистрибьютору вашего региона. В этом письме — главное, что требует вашего внимания до конца Q3.',
    size=11, space_after=10, line_spacing=1.4)

# Главный парадокс
add_heading(doc, 'Главный парадокс сезона')

add_runs_para(doc, [
    ('Первичка (завод → дистрибьютор) ', False, COL_BODY),
    ('выросла на +1.8%', True, COL_GREEN),
    (': 90 504 → 92 094 кг. ', False, COL_BODY),
    ('Вторичка (дистрибьютор → розница) ', False, COL_BODY),
    ('упала на −35%', True, COL_RED),
    (': 30 589 → 20 046 кг. Это значит, что дистрибьюторы закупают TITAN на склад, '
     'но не продают его дальше в розницу. Складские запасы растут, а реальные продажи '
     'конечным клиентам — падают.', False, COL_BODY),
], size=11, space_after=6)

add_para(doc,
    'Если до конца сезона склады не распродаются — в Q4 дистрибьюторы сократят первичную закупку, '
    'и TITAN потеряет ещё больше. Сейчас — окно возможностей: помогите дистрибьюторам продать то, '
    'что уже лежит у них на складе.',
    size=11, bold=True, color=COL_RED, space_after=10, line_spacing=1.4)

# 5 задач
add_heading(doc, '5 задач на Q3 2026')
add_para(doc, 'В порядке приоритета — от критичных к развивающим:', size=10, color=COL_GRAY, space_after=6)

# Задача 1
add_runs_para(doc, [
    ('1. Вернуть 30 потерянных клиентов (потенциал ~10 000 кг). ', True, COL_PRIMARY),
    ('В приложении — список 30 розничных клиентов, которые закупали TITAN в 2025 (>50 кг), '
     'а в 2026 — ни одной закупки. ТОП-3 по потерянному обороту: ИП Хаджи (2 506 кг, через '
     'Технологии РС), ИП Карапетян (1 163 кг, через Технологии РС), ООО Техноторг (832 кг, '
     'через Центр Красок). Действие: каждому падающему дистрибьютору — список его потерянных '
     'клиентов с предложением по возврату (акция «подарок за первую закупку после перерыва», '
     'спецскидка). Срок — до 15 сентября.', False, COL_BODY),
], size=11, space_after=8)

# Задача 2
add_runs_para(doc, [
    ('2. Аудит складских остатков у 12 активных дистрибьюторов. ', True, COL_PRIMARY),
    ('Первичка растёт, вторичка падает — значит, у дистрибьюторов затоваривание. '
     'Свяжитесь с каждым из 12 стабильных дистрибьюторов (см. лист «Дистрибьюторы по группам», '
     'зелёные строки) и запросите остатки TITAN на складе на текущую дату. Цель: понять, у кого '
     'сколько «зависло» кг, и помочь распродать до конца сезона. Срок — до 25 августа.', False, COL_BODY),
], size=11, space_after=8)

# Задача 3
add_runs_para(doc, [
    ('3. Фокус на Севастьянова (−52%) — крупнейшее падение. ', True, COL_PRIMARY),
    ('Севастьянов Э.Ю. (Дальний Восток) — падение с 1 489 до 720 кг. Главный клиент (Бурдинский, '
     '80% всех кг) сократил закупки вдвое. Панской И.А. — нужна личная встреча с Бурдинским до '
     '31 августа. Вопросы для встречи — в документе «10 вопросов РМ» (приложение).', False, COL_BODY),
], size=11, space_after=8)

# Задача 4
add_runs_para(doc, [
    ('4. Активировать 47 «спящих» договоров. ', True, COL_PRIMARY),
    ('В базе 109 договоров с дистрибьюторами, но только 46 стабильных (закупают и в 2025, и в 2026). '
     '47 договоров — без отметок о закупках. Это 44% базы — огромный неиспользованный ресурс. '
     'Действие: до 10 сентября позвонить каждому из ваших «спящих» дистрибьюторов и выяснить причину: '
     'либо активировать (предложить MML + стартовую акцию), либо закрыть договор (чтобы база '
     'отражала реальность).', False, COL_BODY),
], size=11, space_after=8)

# Задача 5
add_runs_para(doc, [
    ('5. Масштабировать опыт Уралстройсервиса (+55%). ', True, COL_PRIMARY),
    ('Уралстройсервис-опт (Урал) — рост с 327 до 508 кг, РТТ выросли с 29 до 58 (×2). '
     'АЮСС-ВСТК (Дальний Восток) — стабильный рост +2-9%, привлёк 5 новых клиентов. '
     'Что они делают правильно? Опишите в отчёте — их опыт тиражируем на других дистрибьюторов. '
     'Срок — до 5 сентября.', False, COL_BODY),
], size=11, space_after=10)

# Что ждём
add_heading(doc, 'Что от вас ждём к 1 сентября')
add_bullet(doc, 'Список 3-5 потерянных клиентов по каждому падающему дистрибьютору — с причиной ухода и предложением для возврата')
add_bullet(doc, 'Остатки TITAN на складе у 12 стабильных дистрибьюторов (в кг по SKU)')
add_bullet(doc, 'Краткий отчёт по встрече с Севастьяновым/Бурдинским (если ДВ — ваш регион)')
add_bullet(doc, 'Статус по 47 «спящим» договорам: кого активируем, кого закрываем')
add_bullet(doc, '1 страница: что Уралстройсервис и АЮСС делают правильно — для тиражирования')

# Поддержка
add_heading(doc, 'Поддержка и материалы')
add_para(doc, 'Все материалы для работы — в приложении к письму:', size=11, space_after=4)
add_bullet(doc, 'Excel «TITAN_анализ_дистрибьюторов_для_программы_v2» — детализация по 109 договорам + 30 потерянных клиентов + 16 новых клиентов')
add_bullet(doc, 'Word «10 вопросов РМ для падающих дистрибьюторов» — готовые вопросы для встреч')
add_bullet(doc, 'PPT «3 цифры программы» — для совещаний с дистрибьюторами')
add_bullet(doc, 'Excel «MML_TITAN_с_продажами_v2» — матрица ассортимента для предложения дистрибьюторам')

add_para(doc, '', size=6, space_after=6)
add_para(doc,
    'Если нужна помощь — пишите. Если есть идеи, как улучшить программу TITAN — обязательно '
    'делитесь. Мы в одной лодке, и конец сезона — это время, когда ещё можно повлиять на '
    'результат года.',
    size=11, space_after=10, line_spacing=1.4)

# Подпись
add_para(doc, 'С уважением,', size=11, space_after=2)
add_para(doc, '[Ваше имя]', size=11, bold=True, color=COL_PRIMARY, space_after=0)
add_para(doc, '[Должность], Empils / TITAN', size=10, color=COL_GRAY, space_after=0)
add_para(doc, '[телефон] | [email]', size=10, color=COL_GRAY, space_after=10)

# Приложения
add_para(doc, '', size=6, space_after=6)
add_para(doc, 'Приложения к письму:', size=10, bold=True, color=COL_ACCENT, space_after=2)
add_bullet(doc, 'TITAN_анализ_дистрибьюторов_для_программы_v2.xlsx', size=10, color=COL_GRAY)
add_bullet(doc, 'TITAN_10_вопросов_РМ_для_падающих_дистр.docx', size=10, color=COL_GRAY)
add_bullet(doc, 'TITAN_3_цифры_программы.pptx', size=10, color=COL_GRAY)
add_bullet(doc, 'MML_TITAN_с_продажами_v2.xlsx', size=10, color=COL_GRAY)

doc.save(OUT)
print(f'OK: {OUT}')
import os
print(f'Размер: {os.path.getsize(OUT) / 1024:.1f} KB')
