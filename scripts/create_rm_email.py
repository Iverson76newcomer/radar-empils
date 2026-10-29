#!/usr/bin/env python3
"""
Word-документ: текст письма для рассылки по региональным менеджерам TITAN.
Цель: мотивация на работу с неактивными дистрибьюторами и возврат потерянных клиентов.
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = '/home/z/my-project/download/Письмо_РМ_TITAN_мотивация.docx'

COL_PRIMARY = RGBColor(0x1E, 0x3A, 0x5F)
COL_ACCENT = RGBColor(0xF4, 0x79, 0x20)
COL_BODY = RGBColor(0x1A, 0x1F, 0x36)
COL_GRAY = RGBColor(0x6B, 0x72, 0x80)
COL_RED = RGBColor(0xC6, 0x28, 0x28)
COL_GREEN = RGBColor(0x16, 0xA3, 0x4A)


def add_para(doc, text='', size=11, bold=False, color=COL_BODY, align=None,
             font='Arial', space_before=0, space_after=4, line_spacing=1.3):
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


def add_heading(doc, text, level=2):
    if level == 1:
        size, color, sb, sa = 16, COL_PRIMARY, 18, 8
    elif level == 2:
        size, color, sb, sa = 13, COL_PRIMARY, 12, 6
    else:
        size, color, sb, sa = 11, COL_ACCENT, 8, 4
    return add_para(doc, text, size=size, bold=True, color=color,
                    align=WD_ALIGN_PARAGRAPH.LEFT, space_before=sb, space_after=sa)


def add_bullet(doc, text, level=0, size=11, bold=False, color=COL_BODY):
    p = doc.add_paragraph(style='List Bullet')
    pf = p.paragraph_format
    pf.left_indent = Cm(1.0 + level * 0.7)
    pf.space_before = Pt(0)
    pf.space_after = Pt(2)
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


def add_numbered(doc, num, title, text, accent_color=COL_ACCENT):
    """Нумерованный пункт с заголовком и описанием."""
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(0.5)
    pf.space_before = Pt(6)
    pf.space_after = Pt(2)
    pf.line_spacing = 1.3
    r1 = p.add_run(f'{num}. ')
    r1.font.name = 'Arial'
    r1.font.size = Pt(11)
    r1.font.bold = True
    r1.font.color.rgb = accent_color
    r2 = p.add_run(title)
    r2.font.name = 'Arial'
    r2.font.size = Pt(11)
    r2.font.bold = True
    r2.font.color.rgb = COL_PRIMARY
    # Описание
    p2 = doc.add_paragraph()
    pf = p2.paragraph_format
    pf.left_indent = Cm(1.2)
    pf.space_before = Pt(0)
    pf.space_after = Pt(4)
    pf.line_spacing = 1.3
    r = p2.add_run(text)
    r.font.name = 'Arial'
    r.font.size = Pt(10.5)
    r.font.color.rgb = COL_BODY


# =========================================================
# ДОКУМЕНТ
# =========================================================
doc = Document()

for section in doc.sections:
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(2.0)

style = doc.styles['Normal']
style.font.name = 'Arial'
style.font.size = Pt(11)
style.font.color.rgb = COL_BODY

# === ШАПКА ПИСЬМА ===
add_para(doc, 'ТЕХНИЧЕСКАЯ ЧАСТЬ ПИСЬМА (для вас — что заполнить в Unisender)', size=9, bold=True, color=COL_GRAY,
         space_after=2)
add_para(doc, '', size=4, space_after=4)

# Subject — 3 варианта для A/B
add_para(doc, 'Subject (3 варианта для A/B теста):', size=10, bold=True, color=COL_ACCENT, space_after=2)
add_bullet(doc, 'Вариант 1 (нейтральный): TITAN — результаты 1 полугодия и приоритеты на Q3', size=10)
add_bullet(doc, 'Вариант 2 (проблемный): TITAN: первичка +1.8%, вторичка −35% — что это значит для вас', size=10, bold=True, color=COL_RED)
add_bullet(doc, 'Вариант 3 (мотивирующий): 5 задач по TITAN на Q3 2026 — ваш фокус до конца сезона', size=10, color=COL_GREEN)

add_para(doc, 'Preheader (превью в инбоксе):', size=10, bold=True, color=COL_ACCENT, space_after=2, space_before=6)
add_para(doc, 'Первичка растёт, вторичка падает — склады затовариваются. 5 конкретных задач для РМ до конца Q3.',
         size=10, color=COL_GRAY, space_after=10)

# === РАЗДЕЛИТЕЛЬ ===
add_para(doc, '—' * 50, size=9, color=COL_GRAY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)

# === ТЕЛО ПИСЬМА ===
add_para(doc, 'Коллеги, добрый день.', size=11, bold=True, color=COL_PRIMARY, space_after=6)

add_para(doc,
    'Подвели итоги по TITAN за 1 полугодие 2026. Цифры — в приложении (Excel-файл с детализацией по '
    'каждому дистрибьютору). В этом письме — главное, что требует вашего внимания до конца Q3.',
    size=11, color=COL_BODY, space_after=8, line_spacing=1.4)

# === ГЛАВНЫЙ ПАРАДОКС ===
add_heading(doc, 'Главный парадокс сезона', level=2)

add_para(doc,
    'Первичка (завод → дистрибьютор) выросла на +1.8%: 90 504 → 92 094 кг. '
    'Вторичка (дистрибьютор → розница) упала на −35%: 30 589 → 20 046 кг. '
    'Это значит, что дистрибьюторы закупают TITAN на склад, но не продают его дальше в розницу. '
    'Складские запасы растут, а реальные продажи конечным клиентам — падают.',
    size=11, color=COL_BODY, space_after=6, line_spacing=1.4)

add_para(doc,
    'Если до конца сезона склады не распродаются — в Q4 дистрибьюторы сократят первичную закупку, '
    'и TITAN потеряет ещё больше. Сейчас — окно возможностей: помогите дистрибьюторам продать то, '
    'что уже лежит у них на складе.',
    size=11, bold=True, color=COL_RED, space_after=10, line_spacing=1.4)

# === 5 ЗАДАЧ ===
add_heading(doc, '5 задач на Q3 2026', level=2)

add_para(doc, 'В порядке приоритета — от критичных к развивающим:', size=10, color=COL_GRAY, space_after=6)

# Задача 1
add_numbered(doc, 1, 'Вернуть 30 потерянных клиентов (потенциал ~10 000 кг)',
    'В приложении — список 30 розничных клиентов, которые закупали TITAN в 2025 году (>50 кг), '
    'а в 2026 — ни одной закупки. ТОП-3 по потерянному обороту: ИП Хаджи (2 506 кг, через Технологии РС), '
    'ИП Карапетян (1 163 кг, через Технологии РС), ООО Техноторг (832 кг, через Центр Красок). '
    'Действие: каждому из ваших падающих дистрибьюторов — список его потерянных клиентов с предложением '
    'по возврату (акция «подарок за первую закупку после перерыва», спецскидка). Срок — до 15 сентября.',
    COL_RED)

# Задача 2
add_numbered(doc, 2, 'Аудит складских остатков у 12 активных дистрибьюторов',
    'Первичка растёт, вторичка падает — значит, у дистрибьюторов затоваривание. '
    'Свяжитесь с каждым из 12 стабильных дистрибьюторов (см. лист «Дистрибьюторы по группам», '
    'зелёные строки) и запросите остатки TITAN на складе на текущую дату. '
    'Цель: понять, у кого сколько «зависло» кг, и помочь распродать до конца сезона. '
    'Срок — до 25 августа.',
    COL_ACCENT)

# Задача 3
add_numbered(doc, 3, 'Фокус на Севастьянова (−52%) — крупнейшее падение',
    'Севастьянов Э.Ю. (Дальний Восток) — падение с 1 489 до 720 кг. '
    'Главный клиент (Бурдинский, 80% всех кг) сократил закупки вдвое. '
    'РМ Панской И.А. — нужна личная встреча с Бурдинским до 31 августа. '
    'Вопросы для встречи — в документе «10 вопросов РМ» (приложение).',
    COL_RED)

# Задача 4
add_numbered(doc, 4, 'Активировать 47 «спящих» договоров',
    'В базе 109 договоров с дистрибьюторами, но только 46 стабильных (закупают и в 2025, и в 2026). '
    '47 договоров — без отметок о закупках. Это 44% базы — огромный неиспользованный ресурс. '
    'Действие: до 10 сентября позвонить каждому из ваших «спящих» дистрибьюторов и выяснить причину: '
    'либо активировать (предложить MML + стартовую акцию), либо закрыть договор (чтобы база отражала реальность).',
    COL_ACCENT)

# Задача 5
add_numbered(doc, 5, 'Масштабировать опыт Уралстройсервиса (+55%)',
    'Уралстройсервис-опт (Урал) — рост с 327 до 508 кг, РТТ выросли с 29 до 58 (×2). '
    'АЮСС-ВСТК (Дальний Восток) — стабильный рост +2-9%, привлёк 5 новых клиентов. '
    'Что они делают правильно? Опишите в отчёте — их опыт тиражируем на других дистрибьюторов. '
    'Срок — до 5 сентября.',
    COL_GREEN)

# === ЧТО ОТ ВАС ЖДЁМ ===
add_heading(doc, 'Что от вас ждём до 1 сентября', level=2)
add_bullet(doc, 'Список 3-5 потерянных клиентов по каждому падающему дистрибьютору — с причиной ухода и предложением для возврата', size=10)
add_bullet(doc, 'Остатки TITAN на складе у 12 стабильных дистрибьюторов (в кг по SKU)', size=10)
add_bullet(doc, 'Краткий отчёт по встрече с Севастьяновым/Бурдинским (если ДВ — ваш регион)', size=10)
add_bullet(doc, 'Статус по 47 «спящим» договорам: кого активируем, кого закрываем', size=10)
add_bullet(doc, '1 страница: что Уралстройсервис и АЮСС делают правильно — для тиражирования', size=10)

# === ПОДДЕРЖКА ===
add_heading(doc, 'Поддержка', level=2)
add_para(doc,
    'Все материалы для работы — в приложении:',
    size=11, color=COL_BODY, space_after=4)
add_bullet(doc, 'Excel «TITAN_анализ_дистрибьюторов_для_программы_v2» — детализация по 109 договорам + 30 потерянных клиентов + 16 новых клиентов', size=10)
add_bullet(doc, 'Word «10 вопросов РМ для падающих дистрибьюторов» — готовые вопросы для встреч', size=10)
add_bullet(doc, 'PPT «3 цифры программы» — для совещаний с дистрибьюторами', size=10)
add_bullet(doc, 'Excel «MML_TITAN_с_продажами» — матрица ассортимента для предложения дистрибьюторам', size=10)

add_para(doc, '', size=6, space_after=6)
add_para(doc,
    'Если нужна помощь — пишите. Если у вас есть идеи, как улучшить программу TITAN — '
    'обязательно делитесь. Мы в одной лодке, и конец сезона — это время, когда ещё можно '
    'повлиять на результат года.',
    size=11, color=COL_BODY, space_after=10, line_spacing=1.4)

# Подпись
add_para(doc, 'С уважением,', size=11, color=COL_BODY, space_after=2)
add_para(doc, '[Ваше имя]', size=11, bold=True, color=COL_PRIMARY, space_after=0)
add_para(doc, '[Должность], Empils / TITAN', size=10, color=COL_GRAY, space_after=0)
add_para(doc, '[телефон] | [email]', size=10, color=COL_GRAY, space_after=10)

# === ПРИЛОЖЕНИЯ ===
add_para(doc, '', size=6, space_after=6)
add_para(doc, '—' * 50, size=9, color=COL_GRAY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
add_para(doc, 'Приложения к письму:', size=10, bold=True, color=COL_ACCENT, space_after=2)
add_bullet(doc, 'TITAN_анализ_дистрибьюторов_для_программы_v2.xlsx (основной отчёт)', size=9)
add_bullet(doc, 'TITAN_10_вопросов_РМ_для_падающих_дистр.docx', size=9)
add_bullet(doc, 'TITAN_3_цифры_программы.pptx', size=9)
add_bullet(doc, 'MML_TITAN_с_продажами_v2.xlsx', size=9)

# === ПРИМЕЧАНИЕ ДЛЯ ВАС ===
add_para(doc, '', size=8, space_after=10)
add_para(doc, '—' * 50, size=9, color=COL_GRAY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
add_para(doc, 'ПРИМЕЧАНИЕ ДЛЯ ВАС (не для рассылки)', size=9, bold=True, color=COL_GRAY, space_after=2)
add_para(doc,
    'Выше — тело письма. Subject и preheader — вверху документа. '
    'Перед отправкой замените [Ваше имя], [Должность], [телефон], [email] на свои данные. '
    'Приложения прикрепите к письму в Unisender (4 файла).',
    size=9, color=COL_GRAY, space_after=4, line_spacing=1.3)

doc.save(OUT)
print(f'OK: {OUT}')
import os
print(f'Размер: {os.path.getsize(OUT) / 1024:.1f} KB')
