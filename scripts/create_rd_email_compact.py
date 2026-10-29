#!/usr/bin/env python3
"""
Word: письмо РД TITAN на половину А4 — компактный формат.
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = '/home/z/my-project/download/Письмо_РД_TITAN_пол_А4.docx'

COL_PRIMARY = RGBColor(0x1E, 0x3A, 0x5F)
COL_ACCENT = RGBColor(0xF4, 0x79, 0x20)
COL_BODY = RGBColor(0x1A, 0x1F, 0x36)
COL_GRAY = RGBColor(0x6B, 0x72, 0x80)
COL_RED = RGBColor(0xC6, 0x28, 0x28)
COL_GREEN = RGBColor(0x16, 0xA3, 0x4A)


def add_para(doc, text='', size=10, bold=False, color=COL_BODY, align=None,
             font='Arial', space_before=0, space_after=3, line_spacing=1.15):
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


def add_runs_para(doc, runs, size=10, space_before=0, space_after=3, line_spacing=1.15):
    p = doc.add_paragraph()
    pf = p.paragraph_format
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


def add_bullet_compact(doc, text, bold_part='', size=9.5):
    """Компактный буллет: жирный заголовок + обычный текст в одном пункте."""
    p = doc.add_paragraph(style='List Bullet')
    pf = p.paragraph_format
    pf.left_indent = Cm(0.5)
    pf.space_before = Pt(0)
    pf.space_after = Pt(2)
    pf.line_spacing = 1.1
    if bold_part:
        run = p.add_run(bold_part)
        run.font.name = 'Arial'
        run.font.size = Pt(size)
        run.font.bold = True
        run.font.color.rgb = COL_PRIMARY
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.append(rFonts)
        rFonts.set(qn('w:ascii'), 'Arial')
        rFonts.set(qn('w:hAnsi'), 'Arial')
    run = p.add_run(text)
    run.font.name = 'Arial'
    run.font.size = Pt(size)
    run.font.color.rgb = COL_BODY
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), 'Arial')
    rFonts.set(qn('w:hAnsi'), 'Arial')


# =========================================================
# ДОКУМЕНТ
# =========================================================
doc = Document()

# Узкие поля для максимума места
for section in doc.sections:
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)

style = doc.styles['Normal']
style.font.name = 'Arial'
style.font.size = Pt(10)
style.font.color.rgb = COL_BODY

# === ШАПКА (тема) ===
add_para(doc, 'ТЕМА: TITAN — 5 задач для РД на Q3 2026 (первичка +1.8%, вторичка −35%)',
         size=9, bold=True, color=COL_GRAY, space_after=6, line_spacing=1.0)

# === ТЕЛО ===
add_para(doc, '[ИМЯ РД], добрый день.', size=11, bold=True, color=COL_PRIMARY, space_after=4)

add_para(doc,
    'Подвели итоги TITAN за 1 полугодие. Главное: первичка +1.8% (90 504 → 92 094 кг), '
    'вторичка −35% (30 589 → 20 046 кг). Дистрибьюторы закупают на склад, но не продают в '
    'розницу. Если до конца сезона не распродадут — Q4 потеряем ещё больше.',
    size=10, space_after=4, line_spacing=1.2)

# 5 задач — компактно
add_para(doc, '5 задач на Q3 2026:', size=10, bold=True, color=COL_ACCENT, space_before=4, space_after=2)

add_bullet_compact(doc, ' — вернуть 30 потерянных клиентов (потенциал ~10 000 кг, список в приложении). Срок: 15 сентября.',
                   bold_part='1. Вернуть потерянных клиентов')

add_bullet_compact(doc, ' — запросить остатки TITAN у 12 стабильных дистрибьюторов (первичка растёт, вторичка падает = затоваривание). Срок: 25 августа.',
                   bold_part='2. Аудит складских остатков')

add_bullet_compact(doc, ' (−52%) — личная встреча с Бурдинским (80% кг, сократил закупки вдвое). Срок: 31 августа.',
                   bold_part='3. Фокус на Севастьянова')

add_bullet_compact(doc, ' — 47 из 109 договоров без закупок (44% базы). Активировать или закрыть. Срок: 10 сентября.',
                   bold_part='4. Активировать «спящих»')

add_bullet_compact(doc, ' (+55%, РТТ ×2) и АЮСС (+9%, 5 новых клиентов) — что делают правильно? Срок: 5 сентября.',
                   bold_part='5. Тиражировать Уралстройсервис')

# Что ждём
add_para(doc, 'Жду к 1 сентября:', size=10, bold=True, color=COL_ACCENT, space_before=4, space_after=2)
add_para(doc,
    '① список потерянных клиентов по падающим дистрибьюторам с предложением возврата  ② '
    'остатки TITAN на складах 12 стабильных  ③ отчёт по встрече с Севастьяновым/Бурдинским  '
    '④ статус по 47 «спящим»  ⑤ 1 страница: кейс Уралстройсервиса и АЮСС.',
    size=9.5, space_after=4, line_spacing=1.15)

# Материалы
add_para(doc,
    'Материалы — в приложении: Excel-анализ 109 дистрибьюторов, 10 вопросов РМ, '
    'PPT «3 цифры», MML-матрица. Пишите, если нужна помощь.',
    size=9.5, color=COL_GRAY, space_after=4, line_spacing=1.2)

# Подпись
add_para(doc, 'С уважением,', size=10, space_after=0, line_spacing=1.0)
add_para(doc, '[Ваше имя], [Должность] · Empils / TITAN · [телефон] · [email]',
         size=9, color=COL_GRAY, space_after=4, line_spacing=1.0)

# Проверка размера
doc.save(OUT)
print(f'OK: {OUT}')
import os
print(f'Размер: {os.path.getsize(OUT) / 1024:.1f} KB')

# Конвертация в PDF для проверки страниц
import subprocess
r = subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', '/tmp', OUT],
                   capture_output=True, text=True)
r2 = subprocess.run(['pdfinfo', '/tmp/Письмо_РД_TITAN_пол_А4.pdf'], capture_output=True, text=True)
for line in r2.stdout.split('\n'):
    if 'Pages' in line or 'Page size' in line:
        print(line)
