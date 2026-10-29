#!/usr/bin/env python3
"""Ultra-compact partnership program — 2 pages max."""

import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT_PATH = '/home/z/my-project/download/Программа_Центр_Красок_2026_2стр.docx'

P_PRIMARY = RGBColor(0xC6, 0x28, 0x28)
P_DARK = RGBColor(0x1A, 0x1F, 0x36)
P_BODY = RGBColor(0x2D, 0x31, 0x42)
P_SECONDARY = RGBColor(0x5A, 0x60, 0x80)
P_ACCENT = RGBColor(0x16, 0xA3, 0x4A)


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
        color = P_PRIMARY if level == 1 else P_DARK
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6 if level == 1 else 4)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(text)
    run.font.bold = True
    run.font.size = Pt(13 if level == 1 else 11)
    run.font.color.rgb = color
    run.font.name = 'Calibri'


def add_p(doc, text, size=10, bold=False, color=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.1
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = 'Calibri'
    if color:
        run.font.color.rgb = color
    return p


def add_b(doc, text):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.line_spacing = 1.1
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
        p.paragraph_format.space_after = Pt(0)
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
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.1
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
style.font.size = Pt(10)
style.paragraph_format.line_spacing = 1.1
style.paragraph_format.space_after = Pt(2)

section = doc.sections[0]
section.left_margin = Cm(1.8)
section.right_margin = Cm(1.5)
section.top_margin = Cm(1.5)
section.bottom_margin = Cm(1.2)

# ===== TITLE BAR =====
title_table = doc.add_table(rows=1, cols=1)
title_cell = title_table.rows[0].cells[0]
set_cell_bg(title_cell, 'C62828')
title_cell.text = ''
p = title_cell.paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(6)
p.paragraph_format.space_after = Pt(2)
run = p.add_run('Empils • TITAN — Программа партнёрского сотрудничества')
run.font.size = Pt(14)
run.font.bold = True
run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
p2 = title_cell.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
p2.paragraph_format.space_after = Pt(6)
run = p2.add_run('Дистрибьютор: «Центр Красок», Краснодар (ЮФО)  •  Период: 06.07.2026 — 31.12.2026  •  Основание: протокол встречи от 01.07.2026')
run.font.size = Pt(9)
run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

# ===== ВСТУПИТЕЛЬНЫЙ АБЗАЦ =====
intro = doc.add_paragraph()
intro.paragraph_format.space_before = Pt(6)
intro.paragraph_format.space_after = Pt(4)
intro.paragraph_format.line_spacing = 1.15
intro.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
run = intro.add_run(
    'Компания «Эмпилс» и дистрибьютор «Центр Красок» связаны многолетним партнёрством, '
    'основанным на взаимном доверии, прозрачности и совместном стремлении к развитию рынка '
    'лакокрасочных материалов в Южном федеральном округе. В непростых экономических реалиях '
    '2026 года особенно важно укреплять отношения, поддерживать друг друга финансовыми и '
    'маркетинговыми инструментами, инвестировать в обучение команд и совместно формировать '
    'ценностное предложение для конечного покупателя. Настоящая программа — практический шаг '
    'навстречу партнёру, учитывающий обратную связь, полученную на встрече 1 июля 2026 года.'
)
run.font.size = Pt(10)
run.font.italic = True
run.font.name = 'Calibri'
run.font.color.rgb = P_BODY

# ===== 1. ФИНАНСОВЫЕ УСЛОВИЯ =====
add_h(doc, '1. Финансовые условия')
make_table(doc,
    ['Условие', 'Размер', 'Порядок выплаты'],
    [
        ['Ежемесячный бонус', '8% от закупок', 'Безоговорочный, в течение 10 раб. дней после месяца'],
        ['Полугодовой бонус', '3% от закупок', 'В течение 20 раб. дней после окончания полугодия'],
        ['Отсрочка платежа', '60 дней', 'При соблюдении графика платежей по предыдущим поставкам'],
    ],
    col_widths=[3.5, 3.5, 10.5])

# ===== 2. ВЫЕЗДНОЕ ОБУЧЕНИЕ =====
add_h(doc, '2. Выездное обучение торговой команды')
add_p(doc, '«Эмпилс» проводит обучение (офлайн, на территории дистрибьютора) для 15–25 сотрудников: '
           'торговых представителей, супервайзеров, тренеров. Дата согласовывается сторонами.')
add_b(doc, 'Модули: о компании «Эмпилс» → продукты TITAN (упор на TITAN 3в1) → методика ТСО (совокупная стоимость владения) → примеры ТСО для розницы и B2B → практикум с выкрасами (TITAN vs Hammerite vs ПФ-115) → работа с возражениями → тестирование и сертификат')
add_b(doc, 'Материалы (шпаргалка, сравнительная таблица, калькулятор ТСО, сертификаты) — за счёт «Эмпилс»; командировочные тренера — за счёт «Эмпилс»; аренда помещения — за счёт дистрибьютора')

# ===== 3. ЭКСКУРСИЯ НА ЗАВОД =====
add_h(doc, '3. Экскурсия на завод «Эмпилс» (Ростов-на-Дону)')
add_p(doc, 'Приглашается делегация до 10 человек. Программа: экскурсия по производству и лаборатории, '
           'дегустация продуктов (TITAN 3в1, Avangarde, Newton), презентация новинок, круглый стол с '
           'руководством. Дата — по согласованию (предпочтительно сентябрь—октябрь 2026). '
           'Программа, обед, сувениры — за счёт «Эмпилс»; проезд и проживание — 50/50.')

# ===== 4. ВЫКРАСЫ TITAN =====
add_h(doc, '4. Обеспечение выкрасами по TITAN')
add_p(doc, '«Эмпилс» бесплатно поставляет образцы покрытий (выкрасы) для демонстрации покупателям и '
           'на обучении: 5 цветов RAL (чёрный, белый, красно-коричневый, серый, зелёный), сравнительные '
           'образцы TITAN vs ПФ-115 и TITAN vs Hammerite, образец «до/после через год». Базовый набор — '
           '1 комплект на филиал, дополнительные — по запросу (до 5). Первая поставка — в течение 14 дней '
           'после подписания. Доставка — за счёт «Эмпилс».')

# ===== 5. КАЛЕНДАРНЫЙ ПЛАН =====
add_h(doc, '5. Календарный план')
make_table(doc,
    ['Срок', 'Мероприятие'],
    [
        ['06.07.2026', 'Старт программы'],
        ['До 20.07.2026', 'Первая поставка выкрасов'],
        ['Июль—Авг 2026', 'Выездное обучение торговой команды'],
        ['Ежемесячно', 'Выплата ежемесячного бонуса (8%)'],
        ['Сент—Окт 2026', 'Экскурсия делегации на завод «Эмпилс»'],
        ['31.12.2026', 'Окончание программы, расчёт полугодового бонуса (3%)'],
        ['Январь 2027', 'Анализ результатов, решение о продлении'],
    ],
    col_widths=[3.5, 14])

# ===== 6. ОЖИДАЕМЫЕ РЕЗУЛЬТАТЫ =====
add_h(doc, '6. Ожидаемые результаты')
add_b(doc, 'Рост продаж TITAN: +20–30% к первому полугодию 2026')
add_b(doc, 'Обучение: ≥20 сотрудников прошли тестирование, ≥80% применяют ТСО')
add_b(doc, 'Восстановление сокращённого ассортимента + расширение присутствия TITAN на полке')

# ===== 7. КООРДИНАЦИЯ =====
add_h(doc, '7. Координация')
add_p(doc, 'Ответственные назначаются с обеих сторон. Ежемесячные онлайн-созвоны, ежеквартальные очные '
           'встречи. Оперативные вопросы — в течение 1–2 рабочих дней. Программа вступает в силу '
           '06.07.2026, действует до 31.12.2026. Не является публичной офертой.')

# ===== ПОДПИСИ =====
sig_table = doc.add_table(rows=2, cols=2)
sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, header in enumerate(['АО «Эмпилс»', '«Центр Красок»']):
    cell = sig_table.rows[0].cells[i]
    cell.text = ''
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(header)
    run.font.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = P_PRIMARY
    set_cell_bg(cell, 'F5F5F5')
    set_cell_borders(cell)
    cell2 = sig_table.rows[1].cells[i]
    cell2.text = ''
    p2 = cell2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(18)
    p2.paragraph_format.space_after = Pt(2)
    run = p2.add_run('____________ / ____________   М.П.')
    run.font.size = Pt(9)
    run.font.color.rgb = P_SECONDARY
    set_cell_borders(cell2)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(4)
run = p.add_run('Дата: «____» ____________ 2026 г.')
run.font.size = Pt(10)
run.font.color.rgb = P_BODY

# Save
doc.save(OUT_PATH)
print(f'Document saved: {OUT_PATH}')
print(f'Size: {os.path.getsize(OUT_PATH)} bytes')
