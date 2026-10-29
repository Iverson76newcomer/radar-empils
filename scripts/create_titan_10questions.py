#!/usr/bin/env python3
"""
Word: список из 10 вопросов РМ для Оберон, БРОЗЭКС, Севастьянов.
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = '/home/z/my-project/download/TITAN_10_вопросов_РМ_для_падающих_дистр.docx'

COL_PRIMARY = RGBColor(0x1E, 0x3A, 0x5F)
COL_ACCENT = RGBColor(0xF4, 0x79, 0x20)
COL_BODY = RGBColor(0x1A, 0x1F, 0x36)
COL_GRAY = RGBColor(0x6B, 0x72, 0x80)
COL_RED = RGBColor(0xC6, 0x28, 0x28)
COL_GREEN = RGBColor(0x16, 0xA3, 0x4A)


def add_para(doc, text='', size=11, bold=False, color=COL_BODY, align=None,
             font='Arial', space_before=0, space_after=4, line_spacing=1.3,
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
    rFonts.set(qn('w:cs'), 'Arial')
    return p


def add_heading(doc, text, level=1):
    if level == 1:
        size, color, sb, sa = 16, COL_PRIMARY, 18, 8
    elif level == 2:
        size, color, sb, sa = 13, COL_PRIMARY, 12, 6
    else:
        size, color, sb, sa = 11, COL_ACCENT, 8, 4
    return add_para(doc, text, size=size, bold=True, color=color,
                    align=WD_ALIGN_PARAGRAPH.LEFT, space_before=sb, space_after=sa)


doc = Document()

for section in doc.sections:
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.0)

style = doc.styles['Normal']
style.font.name = 'Arial'
style.font.size = Pt(11)
style.font.color.rgb = COL_BODY

# Заголовок
add_para(doc, 'ПРОГРАММА ПРОДВИЖЕНИЯ TITAN', size=11, bold=True, color=COL_ACCENT,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_before=10, space_after=2)
add_para(doc, '10 вопросов РМ для встреч с падающими дистрибьюторами',
         size=20, bold=True, color=COL_PRIMARY,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=4)
add_para(doc, 'Технологии РС/Оберон (−86%), БРОЗЭКС (−81%), Севастьянов (−52%), Центр Красок (−28%)',
         size=12, color=COL_GRAY, align=WD_ALIGN_PARAGRAPH.CENTER,
         space_before=0, space_after=8, line_spacing=1.2)
add_para(doc, 'Контекст: АКБ −15%, РТТ −6%, кг −35% (без СтройТерминала). ВСЕ 3 метрики падают.',
         size=11, bold=True, color=COL_RED, align=WD_ALIGN_PARAGRAPH.CENTER,
         space_before=0, space_after=20, line_spacing=1.2)

# Контекст
add_heading(doc, 'Контекст для РМ', level=2)
add_para(doc,
    'Перед встречей с каждым из 4 падающих дистрибьюторов РМ должен держать в голове '
    'контекст: ВСЕ три метрики TITAN падают за полугодие 2025→2026 — АКБ на 15%, РТТ на 6%, '
    'кг на 35% (без СтройТерминала, который не предоставлял статистику в 2025 и искажал картину). '
    '4 компании ниже дают большую часть падения кг. Цель встречи — не «разобраться, что происходит» '
    '(мы это уже знаем), а получить конкретные ответы на 10 вопросов и зафиксировать '
    'договорённости по возврату кг. Каждый вопрос сформулирован под конкретную проблему, '
    'выявленную в отчёте.',
    size=11, color=COL_BODY, space_after=8, line_spacing=1.4)

# Раздел по каждому дистрибьютору
distributors = [
    {
        'name': 'Технологии, Ресурсы Строительства (бывш. Оберон)',
        'rm': 'Еремеев Василий Николаевич (Москва)',
        'kg25': 7947, 'kg26': 1082, 'delta': '-86%',
        'lost_clients': 9,
        'biggest_loss': 'ИП Хаджи Иван Петрович (2 506 кг → 0)',
        'questions': [
            ('Что произошло с ТОП-3 клиентами — Хаджи (2 506 кг), Карапетян (1 163 кг), Гречиха (899 кг)?',
             'Эти 3 клиента = 4 568 кг потерь. Это 91% падения Оберона. Хаджи и Карапетян — армянские фамилии, возможно, ушли к конкурентам по двери/воротам (HD4, Dufa). Узнаём точно — куда ушли.'),
            ('Кто из 9 ушедших клиентов ещё обратим — и какие условия нужны для возврата?',
             'Из 9 потерянных: 6 клиентов были с закупками 6 раз за полугодие (т.е. регулярные). Это не «случайный отток», а системный. Просим РМ назвать поимённо: кого можно вернуть за 1 месяц, кого за 3, кого не вернуть вообще.'),
            ('Почему при падении кг на 86% компания всё же привлекла 5 новых клиентов (МЕТАЛЛСТРОЙ 331 кг, Хачатрян 225 кг)?',
             'Это значит — Технологии РС не «умирают», у них есть приток. Что они делают правильно с новыми? Можно ли это масштабировать на возврат ушедших?'),
            ('Какая сейчас выкладка TITAN в магазинах — и сколько SKU из 30 MML там реально представлены?',
             'Гипотеза: доля TITAN в АКБ упала с 5.6% до 3.9% — то есть даже те, кто покупает, берут меньше. Возможно, TITAN просто убрали с активной полки. Просим фото выкладки из 3 ключевых точек.'),
            ('Какой план возврата кг на Q3 2026 — конкретно в цифрах по каждому из ТОП-5 потерянных клиентов?',
             'Цель — получить от РМ конкретный план: «Хаджи возвращаем за счёт X, Карапетян — за счёт Y». Без плана встреча считается провальной.'),
        ]
    },
    {
        'name': 'БРОЗЭКС',
        'rm': 'Захарова Ольга Геннадьевна (Урал)',
        'kg25': 2154, 'kg26': 417, 'delta': '-81%',
        'lost_clients': 'крупнейший внутренний',
        'biggest_loss': 'Эмпилс ЗАО (1 179 кг → 92 кг) — ВНУТРЕННИЙ КЛИЕНТ',
        'questions': [
            ('Почему Эмпилс ЗАО (завод-поставщик!) перестал закупать у БРОЗЭКСа — с 1 179 кг до 92 кг?',
             'Это аномалия: Эмпилс сам производит TITAN, но раньше закупал через БРОЗЭКС. Что изменилось — прямые поставки? Конфликт по ценам? Эта одна позиция = 55% всего падения БРОЗЭКСа.'),
            ('БРОЗЭКС — это Урал. Кто из конкурентов (Tikkurila, Эксперт, Alpina) забрал долю TITAN на уральских РТТ?',
             'РТТ БРОЗЭКСа упали с 59 до 34 (−42%). Это не потеря клиентов, а потеря розничных точек. Что происходит с самой сетью БРОЗЭКС — закрывают точки? Или TITAN убрали из ассортимента?'),
            ('Какая цена закупки TITAN у БРОЗЭКСа по сравнению с другими уральскими дистрибьюторами (Зыков, Протасов, Уралстройсервис)?',
             'Гипотеза: БРОЗЭКС теряет ценовую конкурентоспособность. Уралстройсервис растёт +55% на том же рынке — значит, проблема не в рынке, а в БРОЗЭКСе. Сравниваем прайсы.'),
            ('Какие 3 SKU TITAN БРОЗЭКС готов закладывать в MML для восстановления — и какой объём?',
             'Цель — перевести разговор из «у нас всё плохо» в «вот 3 SKU, которые мы можем качать». Минимум 3 SKU × 50 кг/мес × 3 мес = 450 кг. Это уже +50% к текущим 417 кг.'),
        ]
    },
    {
        'name': 'Севастьянов Эдуард Юрьевич',
        'rm': 'Панская Ирина Александровна (Дальний Восток)',
        'kg25': 1489, 'kg26': 720, 'delta': '-52%',
        'lost_clients': 'крупнейший -1',
        'biggest_loss': 'ИП Бурдинский Андрей Владимирович (1 185 кг → 583 кг, −50%)',
        'questions': [
            ('Почему Бурдинский сократил закупки вдвое (1 185 → 583 кг) — это временно или он переводит долю на другой бренд?',
             'Бурдинский = 80% всех кг Севастьянова. Если он уходит к конкуренту — Севастьянов теряет смысл как дистрибьютор TITAN. Нужна личная встреча РМ с Бурдинским.'),
            ('РТТ Севастьянова почти не изменились (14 → 13) — то есть точки те же, но в каждой берут меньше. Что происходит внутри точек?',
             'Это «молчаливый отток» — никто не ушёл, просто TITAN постепенно вытесняется с полки. Просим РМ сделать 5 фото выкладок в течение 2 недель — сравним с фото год назад.'),
            ('Севастьянов — Дальний Восток. Есть ли логистическая проблема (сроки поставки, остатки на складе), которая мешает качественно продавать TITAN?',
             'ДВ — особый кейс: длинное плечо доставки, возможны перебои. Если TITAN лежит на складе 2 месяца — это объясняет, почему РТТ берут реже. Узнаём остатки на складе Севастьянова прямо сейчас.'),
        ]
    },
    {
        'name': 'Центр Красок',
        'rm': 'Богородицкая Надежда Владимировна (Поволжье)',
        'kg25': 6663, 'kg26': 4822, 'delta': '-28%',
        'lost_clients': 6,
        'biggest_loss': 'ИП Урахчин (868 кг), ООО Техноторг (832 кг) — оба ушли в 0',
        'questions': [
            ('Центр Красок привлёк 8 из 16 новых клиентов TITAN по всей России — что они делают правильно, и можно ли это тиражировать на других дистрибьюторов?',
             'Это позитивная часть разговора. Просим РМ описать механику привлечения: какие SKU ставят новым клиентам, какие акции, как ТП работает. Это материал для ролика «кейс успеха» в ТЗ.'),
            ('Куда ушли Урахчин (868 кг) и Техноторг (832 кг) — это 1 700 кг потерь. Можно ли их вернуть?',
             'Оба клиента — Краснодарский край. Возможно, ушли к местному конкуренту («АфинаЮг»?). Это уже третья позиция в списке потерь — Урахчин, Техноторг, Строй-Ритм (120 кг).'),
            ('РТТ Центра Красок упали с 130 до 98 (−24%) — но АКБ выросли с 627 до 485... нет, АКБ тоже упали. Это уже не «точки те же, берут меньше», а «точек меньше». Что происходит с сетью?',
             'Центр Красок — крупнейший дистрибьютор TITAN по числу РТТ (130 в 2025). Если РТТ сокращаются на 24% — это структурная проблема, а не конъюнктурная. Нужно понять: точки закрываются, или TITAN из них выводят?'),
            ('Какой план по РТТ на Q3-Q4 2026 — хотя бы вернуть 130?',
             'Цель — зафиксировать: сколько РТТ Центр Красок планирует иметь на конец года. Если меньше 100 — это сигнал, что дистрибьютор strategically уходит из TITAN.'),
        ]
    },
]

for d in distributors:
    # Заголовок дистрибьютора
    add_heading(doc, f'{d["name"]} ({d["delta"]})', level=1)
    add_para(doc, f'РМ: {d["rm"]}    |    Кг: {d["kg25"]} → {d["kg26"]}    |    Главные потери: {d["biggest_loss"]}',
             size=10, color=COL_GRAY, space_after=8, line_spacing=1.3)
    
    for i, (q, why) in enumerate(d['questions'], 1):
        # Вопрос
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.left_indent = Cm(0.5)
        pf.space_before = Pt(8)
        pf.space_after = Pt(2)
        pf.line_spacing = 1.3
        r1 = p.add_run(f'Вопрос {i}. ')
        r1.font.name = 'Arial'
        r1.font.size = Pt(11)
        r1.font.bold = True
        r1.font.color.rgb = COL_ACCENT
        r2 = p.add_run(q)
        r2.font.name = 'Arial'
        r2.font.size = Pt(11)
        r2.font.bold = True
        r2.font.color.rgb = COL_PRIMARY
        # Зачем спрашивать
        p2 = doc.add_paragraph()
        pf = p2.paragraph_format
        pf.left_indent = Cm(1.2)
        pf.space_before = Pt(0)
        pf.space_after = Pt(4)
        pf.line_spacing = 1.3
        r1 = p2.add_run('Зачем спрашивать: ')
        r1.font.name = 'Arial'
        r1.font.size = Pt(10)
        r1.font.bold = True
        r1.font.italic = True
        r1.font.color.rgb = COL_GRAY
        r2 = p2.add_run(why)
        r2.font.name = 'Arial'
        r2.font.size = Pt(10)
        r2.font.italic = True
        r2.font.color.rgb = COL_BODY

    doc.add_paragraph()  # пустая строка между дистрибьюторами

# Итоговая секция — что должно быть на выходе
doc.add_page_break()
add_heading(doc, 'Что РМ должен привезти с каждой встречи', level=1)
add_para(doc,
    'Встреча считается успешной, если по каждому дистрибьютору РМ привозит 4 конкретных артефакта:',
    size=11, color=COL_BODY, space_after=8, line_spacing=1.4)

add_bullet(doc, 'Ответы на все 10 вопросов (по каждому дистрибьютору свой набор из 3-5 вопросов) — в письменной форме, не «в общих чертах».')
add_bullet(doc, 'Список из 3-5 конкретных клиентов, которых можно вернуть в течение Q3 2026, с указанием причины ухода и предложения для возврата.')
add_bullet(doc, 'План по кг на Q3 2026 — в цифрах. Минимум: вернуть 30% от падения (т.е. для Технологии РС — это +2 000 кг к текущим 1 082 кг).')
add_bullet(doc, 'Фото выкладки TITAN из 3 ключевых точек каждого дистрибьютора — для сравнения с фото год назад и аудита MML.')

# Финальная напоминалка
add_para(doc, '', size=8, space_after=10)
add_heading(doc, 'Принцип встречи', level=2)
add_para(doc,
    'Главная задача РМ — не «продать TITAN дистрибьютору», а понять, почему дистрибьютор '
    'перестал продавать TITAN своим клиентам. Ответ почти всегда не в цене TITAN, а в том, '
    'что происходит внутри дистрибьютора: либо его собственная сеть сжимается, либо TITAN '
    'вытесняется с полки конкурентом, либо РТТ переходят на другой бренд. Все 10 вопросов '
    'построены так, чтобы РМ вышел за рамки «у нас всё нормально, просто сезон такой» и '
    'получил конкретику.',
    size=11, color=COL_BODY, space_after=10, line_spacing=1.4)

add_para(doc,
    'ВАЖНО: по итогам анализа (без СтройТерминала) ВСЕ три метрики TITAN падают — '
    'АКБ на 15%, РТТ на 6%, кг на 35%. Это не «локальная проблема 4 дистрибьюторов», '
    'а системное сжатие. Если хотя бы один из 4 (Технологии РС, БРОЗЭКС, Севастьянов, '
    'Центр Красок) не вернёт хотя бы 30% потерянных кг до конца Q3 2026 — программа '
    'продвижения TITAN за год не выйдет в плюс.',
    size=11, bold=True, color=COL_RED, space_after=10, line_spacing=1.4)

doc.save(OUT)
print(f'OK: {OUT}')
import os
print(f'Размер: {os.path.getsize(OUT) / 1024:.1f} KB')
