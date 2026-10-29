#!/usr/bin/env python3
"""
Техническое задание для дизайнера и типографии — POS-материалы TITAN.
3 инструмента: Шелфтокер, Воблер, Брендированный ценник.
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

OUT = '/home/z/my-project/download/ТЗ_POSM_TITAN_шелфтокер_воблер_ценник.docx'

RED = RGBColor(0xC6, 0x28, 0x28)
DARK = RGBColor(0x1A, 0x1F, 0x36)
GRAY = RGBColor(0x60, 0x60, 0x60)
BLACK = RGBColor(0x00, 0x00, 0x00)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)


def set_para_format(p, line_spacing=1.3, space_after=6, space_before=0,
                    first_line_indent=0, alignment=None):
    pf = p.paragraph_format
    pf.line_spacing = line_spacing
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(space_before)
    if first_line_indent:
        pf.first_line_indent = Cm(first_line_indent)
    if alignment is not None:
        p.alignment = alignment


def add_para(doc, text='', size=11, bold=False, color=DARK,
             alignment=WD_ALIGN_PARAGRAPH.LEFT, italic=False,
             line_spacing=1.3, space_after=6, space_before=0,
             first_line_indent=0, font_name='Arial'):
    p = doc.add_paragraph()
    set_para_format(p, line_spacing=line_spacing, space_after=space_after,
                    space_before=space_before,
                    first_line_indent=first_line_indent, alignment=alignment)
    if text:
        run = p.add_run(text)
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = color
        run.font.name = font_name
    return p


def add_divider(doc, color='C62828'):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(8)
    pPr = p._element.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '12')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), color)
    pBdr.append(bottom)
    pPr.append(pBdr)
    return p


def add_section_label(doc, text, color=RED):
    p = doc.add_paragraph()
    set_para_format(p, line_spacing=1.2, space_after=4, space_before=10)
    run = p.add_run(text)
    run.font.size = Pt(12)
    run.font.bold = True
    run.font.color.rgb = color
    run.font.name = 'Arial'
    rPr = run._element.get_or_add_rPr()
    caps = OxmlElement('w:caps')
    caps.set(qn('w:val'), 'true')
    rPr.append(caps)
    return p


def add_table_row(table, row_idx, cells_data, bold_cols=None):
    """cells_data: список строк"""
    bold_cols = bold_cols or []
    for ci, text in enumerate(cells_data):
        cell = table.cell(row_idx, ci)
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(text)
        run.font.size = Pt(10)
        run.font.name = 'Arial'
        run.font.color.rgb = DARK
        if ci in bold_cols:
            run.font.bold = True


def add_styled_table(doc, headers, rows, col_widths=None):
    """Создать таблицу со стилем."""
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'
    
    # Заголовки
    for ci, h in enumerate(headers):
        cell = table.cell(0, ci)
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        run.font.size = Pt(10)
        run.font.bold = True
        run.font.color.rgb = WHITE
        run.font.name = 'Arial'
        # Заливка
        shading = OxmlElement('w:shd')
        shading.set(qn('w:val'), 'clear')
        shading.set(qn('w:fill'), '1E3A5F')
        cell._element.get_or_add_tcPr().append(shading)
    
    # Данные
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = table.cell(ri + 1, ci)
            cell.text = ''
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(str(val))
            run.font.size = Pt(10)
            run.font.name = 'Arial'
            run.font.color.rgb = DARK
            if ci == 0:
                run.font.bold = True
    
    # Ширина колонок
    if col_widths:
        for ci, w in enumerate(col_widths):
            for cell in table.columns[ci].cells:
                cell.width = Cm(w)
    
    return table


# ==============================================================
# СОЗДАНИЕ ДОКУМЕНТА
# ==============================================================

def main():
    doc = Document()

    for section in doc.sections:
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(2.0)
        section.right_margin = Cm(1.5)

    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)

    # === ЗАГОЛОВОК ===
    p = doc.add_paragraph()
    set_para_format(p, line_spacing=1.2, space_after=4)
    run = p.add_run('TITAN')
    run.font.size = Pt(20)
    run.font.bold = True
    run.font.color.rgb = RED
    run.font.name = 'Arial'

    add_para(doc, 'Техническое задание для дизайнера и типографии',
             size=14, color=DARK, bold=True, space_after=4)
    add_para(doc, 'POS-материалы для выкладки TITAN на полке: шелфтокер, воблер, брендированный ценник',
             size=11, color=GRAY, italic=True, space_after=4)
    add_para(doc, 'Empils  •  2026  •  Бренд: TITAN (ОРЕОЛ TITAN)',
             size=10, color=GRAY, italic=True, space_after=2)

    add_divider(doc)

    # === ОБЩАЯ ИНФОРМАЦИЯ ===
    add_section_label(doc, 'Общая информация')

    add_styled_table(doc,
        ['Параметр', 'Значение'],
        [
            ['Бренд', 'TITAN (ОРЕОЛ TITAN)'],
            ['Производитель', 'ЗАО «Эмпилс», Ростов-на-Дону'],
            ['Продукт', 'Грунт-эмаль 3в1 по ржавчине TITAN'],
            ['Слоган (основной)', '«Красишь 1 раз — служит 10 лет»'],
            ['Слоган (доп.)', '«У коррозии нет шансов»'],
            ['Сайт', 'www.titanpaints.ru'],
            ['Цвета бренда', 'Красный #C62828, Жёлтый #FBBF24, Белый #FFFFFF, Чёрный #1A1A1A'],
            ['QR-код', 'Ведёт на www.titanpaints.ru'],
            ['Масштаб', 'Розничные точки классов A, B, C (федеральные DIY, локальные сети, супермаркеты)'],
        ],
        col_widths=[5, 12]
    )

    add_divider(doc)

    # ==============================================================
    # 1. ШЕЛФТОКЕР
    # ==============================================================
    add_section_label(doc, '1. Шелфтокер — техническое задание')

    add_para(doc, 'Назначение', size=12, bold=True, color=RED, space_before=8, space_after=4)
    add_para(doc,
             'Шелфтокер устанавливается на край полки с выкладкой TITAN '
             '(крепится на торец полки с помощью клеевого слоя или клипсы). '
             'Привлекает внимание покупателя к продукту, транслирует ключевое УТП '
             'и направляет на сайт через QR-код.',
             size=11, space_after=6)

    add_para(doc, 'Текстовая информация на макете', size=12, bold=True, color=RED, space_before=8, space_after=4)

    add_styled_table(doc,
        ['Зона', 'Текст', 'Шрифт / Стиль'],
        [
            ['Логотип', 'TITAN', 'Крупно, белым на красном фоне. Фирменный стиль Empils.'],
            ['Слоган (основной)', 'Красишь 1 раз — служит 10 лет', 'Крупно, жирно. Белым на красном или жёлтым на красном.'],
            ['Доп. текст', 'Грунт-эмаль 3в1 по ржавчине', 'Мелким шрифтом, под слоганом. Белым.'],
            ['QR-код', 'QR-код → www.titanpaints.ru', 'В правом нижнем углу. Размер квадрата — не менее 25×25 мм.'],
            ['Подпись QR', 'Сканируй → видео + калькулятор', 'Мелким шрифтом под QR-кодом.'],
        ],
        col_widths=[3.5, 7, 6]
    )

    add_para(doc, 'Визуальные элементы', size=12, bold=True, color=RED, space_before=8, space_after=4)
    add_para(doc, '• Фон — красный (#C62828) с градиентом или без', size=11, space_after=2)
    add_para(doc, '• Логотип TITAN — в левой или центральной части', size=11, space_after=2)
    add_para(doc, '• QR-код — в правой части, на белом квадрате-подложке', size=11, space_after=2)
    add_para(doc, '• Изображение банки TITAN 3в1 (опционально) — слева', size=11, space_after=2)
    add_para(doc, '• Жёлтая полоса (#FBBF24) — разделитель под логотипом', size=11, space_after=6)

    add_para(doc, 'Технические параметры для типографии', size=12, bold=True, color=RED, space_before=8, space_after=4)

    add_styled_table(doc,
        ['Параметр', 'Значение'],
        [
            ['Тип', 'Шелфтокер (без стопера) — плоский, на торец полки'],
            ['Размер (ширина × высота)', '300 × 100 мм'],
            ['Форма', 'Прямоугольная, нижний край с фигурным вырезом (визуальный акцент)'],
            ['Материал', 'Картон мелованный 300 г/м²'],
            ['Покрытие', 'Ламинация глянцевая (односторонняя)'],
            ['Печать', '4+0 (полноцвет, лицевая сторона)'],
            ['Вырубка', 'По контуру + фигурный вырез снизу'],
            ['Крепление', 'Двусторонний скотч (на оборотную сторону) или клипса-держатель полки'],
            ['Тираж (ориентир.)', '500 шт (пилот) → 2000+ шт (масштаб)'],
            ['Срок изготовления', '7-10 рабочих дней после утверждения макета'],
        ],
        col_widths=[5, 12]
    )

    add_divider(doc)

    # ==============================================================
    # 2. ВОБЛЕР
    # ==============================================================
    add_section_label(doc, '2. Воблер — техническое задание')

    add_para(doc, 'Назначение', size=12, bold=True, color=RED, space_before=8, space_after=4)
    add_para(doc,
             'Воблер подвешивается над полкой с выкладкой TITAN на гибкой ножке. '
             'Привлекает внимание за счёт движения («колыхается» при проходе покупателя). '
             'Транслирует короткий слоган — «ловушку внимания».',
             size=11, space_after=6)

    add_para(doc, 'Текстовая информация на макете', size=12, bold=True, color=RED, space_before=8, space_after=4)

    add_styled_table(doc,
        ['Зона', 'Текст', 'Шрифт / Стиль'],
        [
            ['Логотип', 'TITAN', 'Крупно, белым на красном фоне.'],
            ['Слоган', 'Красишь 1 раз\nслужит 10 лет', 'Крупно, жирно. Белым или жёлтым. В 2 строки.'],
            ['Доп. текст', 'Грунт-эмаль 3в1', 'Мелким шрифтом, под слоганом.'],
        ],
        col_widths=[3.5, 7, 6]
    )

    add_para(doc, 'Визуальные элементы', size=12, bold=True, color=RED, space_before=8, space_after=4)
    add_para(doc, '• Фон — жёлтый (#FBBF24) с красной окантовкой', size=11, space_after=2)
    add_para(doc, '• Логотип TITAN — вверху, красным на жёлтом', size=11, space_after=2)
    add_para(doc, '• Слоган — центр, крупно', size=11, space_after=2)
    add_para(doc, '• Ножка — прозрачная гибкая пластиковая полоса (30-40 мм)', size=11, space_after=6)

    add_para(doc, 'Технические параметры для типографии', size=12, bold=True, color=RED, space_before=8, space_after=4)

    add_styled_table(doc,
        ['Параметр', 'Значение'],
        [
            ['Тип', 'Воблер на гибкой ножке'],
            ['Размер плашки (Ш × В)', '120 × 80 мм'],
            ['Форма плашки', 'Прямоугольная со скруглёнными углами (R=5 мм)'],
            ['Ножка', 'Прозрачный ПВХ, 30-40 мм, клеевой квадрат 15×15 мм'],
            ['Материал плашки', 'Картон мелованный 350 г/м²'],
            ['Покрытие', 'Ламинация матовая (двусторонняя)'],
            ['Печать', '4+0 (полноцвет, лицевая сторона)'],
            ['Вырубка', 'По контуру + отверстие для ножки (Ø4 мм)'],
            ['Крепление', 'Клеевой квадрат на ножке → крепится к полке/ценнику'],
            ['Тираж (ориентир.)', '1000 шт (пилот) → 3000+ шт (масштаб)'],
            ['Срок изготовления', '7-10 рабочих дней после утверждения макета'],
        ],
        col_widths=[5, 12]
    )

    add_divider(doc)

    # ==============================================================
    # 3. БРЕНДИРОВАННЫЙ ЦЕННИК
    # ==============================================================
    add_section_label(doc, '3. Брендированный ценник — техническое задание')

    add_para(doc, 'Назначение', size=12, bold=True, color=RED, space_before=8, space_after=4)
    add_para(doc,
             'Брендированный ценник устанавливается на полку рядом с банками TITAN. '
             'Переводит цену банки в «цену за год» — показывает покупателю, что '
             'TITAN дешевле, чем кажется. Главное: не просто цена, а УТП «дешевле, '
             'если считать на 10 лет».',
             size=11, space_after=6)

    add_para(doc, 'Текстовая информация на макете', size=12, bold=True, color=RED, space_before=8, space_after=4)

    add_styled_table(doc,
        ['Зона', 'Текст', 'Шрифт / Стиль'],
        [
            ['Логотип', 'TITAN', 'Вверху, мелко. Белым на красной плашке.'],
            ['Заголовок', 'ЦЕНА ЗА ГОД', 'Крупно, жирно. Жёлтым на красном.'],
            ['Цена за год', '532 ₽/год', 'Очень крупно, белым на красном. Главная цифра.'],
            ['Расчёт', '5 320 ₽ за банку ÷ 10 лет', 'Мелким шрифтом, под ценой за год. Белым.'],
            ['Слоган', 'Красишь 1 раз — служит 10 лет', 'Внизу ценника. Белым, курсивом.'],
        ],
        col_widths=[3.5, 7, 6]
    )

    add_para(doc, 'Визуальные элементы', size=12, bold=True, color=RED, space_before=8, space_after=4)
    add_para(doc, '• Фон — красный (#C62828)', size=11, space_after=2)
    add_para(doc, '• «Цена за год» — самая крупная цифра на ценнике', size=11, space_after=2)
    add_para(doc, '• Поле для розничной цены (пустое) — внизу, на белом фоне, для ценника магазина', size=11, space_after=2)
    add_para(doc, '• Жёлтая полоса (#FBBF24) — разделитель', size=11, space_after=6)

    add_para(doc, 'Технические параметры для типографии', size=12, bold=True, color=RED, space_before=8, space_after=4)

    add_styled_table(doc,
        ['Параметр', 'Значение'],
        [
            ['Тип', 'Брендированный ценник (двусторонний)'],
            ['Размер (Ш × В)', '70 × 40 мм (стандартный размер полочного ценника)'],
            ['Форма', 'Прямоугольная, без скруглений'],
            ['Лицевая сторона', 'Брендированная: логотип TITAN + «Цена за год» + слоган'],
            ['Оборотная сторона', 'Белая — поле для цены магазина (ценник наклеивается)'],
            ['Материал', 'Картон мелованный 250 г/м²'],
            ['Покрытие', 'Ламинация глянцевая (лицевая сторона)'],
            ['Печать', '4+0 (лицевая) + 1+0 (оборот: белый фон)'],
            ['Вырубка', 'Прямая (прямоугольник)'],
            ['Крепление', 'Клипса-держатель для полки (заказывается отдельно) или двусторонний скотч'],
            ['Тираж (ориентир.)', '2000 шт (пилот) → 5000+ шт (масштаб)'],
            ['Срок изготовления', '5-7 рабочих дней после утверждения макета'],
        ],
        col_widths=[5, 12]
    )

    add_divider(doc)

    # ==============================================================
    # ОБЩИЕ ТРЕБОВАНИЯ
    # ==============================================================
    add_section_label(doc, 'Общие требования к макетам')

    add_para(doc, 'Форматы файлов для типографии', size=12, bold=True, color=RED, space_before=8, space_after=4)
    add_para(doc, '• Основной формат: PDF (Press Quality), CMYK, 300 dpi', size=11, space_after=2)
    add_para(doc, '• Дополнительный: AI (Adobe Illustrator) или EPS — векторный исходник', size=11, space_after=2)
    add_para(doc, '• Шрифты: переведены в кривые (outlines)', size=11, space_after=2)
    add_para(doc, '• Отступы (вылеты): по 3 мм с каждой стороны (bleed)', size=11, space_after=2)
    add_para(doc, '• Цветовая модель: CMYK (не RGB!)', size=11, space_after=6)

    add_para(doc, 'Цветовая палитра (CMYK)', size=12, bold=True, color=RED, space_before=8, space_after=4)

    add_styled_table(doc,
        ['Цвет', 'HEX', 'CMYK', 'Где используется'],
        [
            ['Красный TITAN', '#C62828', 'C0 M100 Y80 K20', 'Фон шелфтокера, ценника, логотип'],
            ['Жёлтый', '#FBBF24', 'C0 M25 Y95 K0', 'Полосы, акценты, воблер'],
            ['Белый', '#FFFFFF', 'C0 M0 Y0 K0', 'Текст на красном фоне'],
            ['Чёрный', '#1A1A1A', 'C0 M0 Y0 K100', 'Дополнительный текст'],
        ],
        col_widths=[4, 3, 4, 6]
    )

    add_para(doc, 'Шрифты', size=12, bold=True, color=RED, space_before=8, space_after=4)
    add_para(doc, '• Основной шрифт: Arial (или аналог — Roboto, Helvetica)', size=11, space_after=2)
    add_para(doc, '• Заголовки: Arial Bold', size=11, space_after=2)
    add_para(doc, '• Основной текст: Arial Regular', size=11, space_after=2)
    add_para(doc, '• Все шрифты в макете должны быть переведены в кривые', size=11, space_after=6)

    add_para(doc, 'QR-код — требования', size=12, bold=True, color=RED, space_before=8, space_after=4)
    add_para(doc, '• Ссылка: https://www.titanpaints.ru', size=11, space_after=2)
    add_para(doc, '• Размер: не менее 25 × 25 мм (для корректного сканирования)', size=11, space_after=2)
    add_para(doc, '• Фон: белый квадрат-подложка вокруг QR-кода (отступ 3 мм)', size=11, space_after=2)
    add_para(doc, '• Контраст: чёрный QR-код на белом фоне', size=11, space_after=2)
    add_para(doc, '• Уровень коррекции ошибок: M (15%)', size=11, space_after=6)

    add_para(doc, 'Маркировка', size=12, bold=True, color=RED, space_before=8, space_after=4)
    add_para(doc, 'На оборотной стороне каждого POS-материала (мелким шрифтом):', size=11, space_after=4)
    add_para(doc, '© ЗАО «Эмпилс», 2026  •  www.titanpaints.ru  •  Ростов-на-Дону, пер. 1-й Машиностроительный, 21',
             size=10, color=GRAY, italic=True, space_after=6)

    add_divider(doc)

    # ==============================================================
    # ПРИЛОЖЕНИЕ
    # ==============================================================
    add_section_label(doc, 'Приложение: визуальные примеры')

    add_para(doc, 'Изображения POS-материалов для reference', size=12, bold=True, color=RED, space_before=8, space_after=4)

    # Встраиваем визуализации из папки titan_posm_visuals
    posm_dir = '/home/z/my-project/download/titan_posm_visuals'
    visuals = [
        ('01_shelftalker.png', '1. Шелфтокер — визуализация'),
        ('02_wobbler.png', '2. Воблер — визуализация'),
        ('03_price_tag.png', '3. Брендированный ценник — визуализация'),
    ]

    for fname, caption in visuals:
        path = os.path.join(posm_dir, fname)
        if os.path.exists(path):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run()
            run.add_picture(path, width=Cm(12))
            add_para(doc, caption, size=10, color=GRAY, italic=True,
                     alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)

    # Сохранение
    doc.save(OUT)
    print(f'OK: {OUT}')


if __name__ == '__main__':
    main()
