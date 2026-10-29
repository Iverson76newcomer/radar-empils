"""
Отчёт по переговорам с сетью супермаркетов «12 месяцев» (Казахстан)
Выставка, переговоры с действующим партнёром.
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ORANGE = RGBColor(0xF4, 0x76, 0x29)
DARK   = RGBColor(0x31, 0x31, 0x31)
GREY   = RGBColor(0x6B, 0x72, 0x80)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
RED    = RGBColor(0xC0, 0x39, 0x2B)
GREEN  = RGBColor(0x16, 0xA3, 0x4A)
AMBER  = RGBColor(0xB8, 0x86, 0x0B)
BLUE   = RGBColor(0x1E, 0x3A, 0x5F)

doc = Document()
style = doc.styles['Normal']
style.font.name = 'Arial'
style.font.size = Pt(11)
style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')

for section in doc.sections:
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)

def set_spacing(p, before=0, after=6, line=1.3):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line

def add_run(p, text, size=11, bold=False, color=None, italic=False):
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.bold = bold
    r.italic = italic
    if color:
        r.font.color.rgb = color
    r.font.name = 'Arial'
    return r

def add_para(text, size=11, bold=False, color=None, italic=False, align='left', after=6, before=0):
    p = doc.add_paragraph()
    set_spacing(p, before=before, after=after, line=1.3)
    p.alignment = {'center': WD_ALIGN_PARAGRAPH.CENTER, 'right': WD_ALIGN_PARAGRAPH.RIGHT,
                   'left': WD_ALIGN_PARAGRAPH.LEFT, 'justify': WD_ALIGN_PARAGRAPH.JUSTIFY}.get(align, WD_ALIGN_PARAGRAPH.LEFT)
    add_run(p, text, size=size, bold=bold, color=color, italic=italic)
    return p

def add_h1(text, color=DARK, size=20, before=0, after=8, orange_bar=True):
    p = doc.add_paragraph()
    set_spacing(p, before=before, after=after, line=1.1)
    add_run(p, text, size=size, bold=True, color=color)
    if orange_bar:
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        left = OxmlElement('w:left')
        left.set(qn('w:val'), 'single')
        left.set(qn('w:sz'), '24')
        left.set(qn('w:space'), '8')
        left.set(qn('w:color'), 'F47629')
        pBdr.append(left)
        pPr.append(pBdr)
        p.paragraph_format.left_indent = Cm(0.3)
    return p

def add_h2(text, color=ORANGE, size=14, before=10, after=4):
    p = doc.add_paragraph()
    set_spacing(p, before=before, after=after, line=1.2)
    add_run(p, text, size=size, bold=True, color=color)
    return p

def add_h3(text, color=DARK, size=12, before=8, after=3):
    p = doc.add_paragraph()
    set_spacing(p, before=before, after=after, line=1.2)
    add_run(p, text, size=size, bold=True, color=color)
    return p

def add_bullet(text, size=11, bold=False):
    p = doc.add_paragraph(style='List Bullet')
    set_spacing(p, before=0, after=3, line=1.25)
    p.paragraph_format.left_indent = Cm(0.6)
    r = p.runs[0] if p.runs else p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.name = 'Arial'
    if bold:
        r.bold = True
    return p

def set_cell_bg(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def set_cell_borders(cell, color='D0D0D0', sz='4'):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), sz)
        el.set(qn('w:color'), color)
        tcBorders.append(el)
    tcPr.append(tcBorders)

def add_table(headers, rows, col_widths, font_size=10, header_bg='313131'):
    tbl = doc.add_table(rows=1 + len(rows), cols=len(headers))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    for i, w in enumerate(col_widths):
        for cell in tbl.columns[i].cells:
            cell.width = Cm(w)
    for i, h in enumerate(headers):
        c = tbl.rows[0].cells[i]
        set_cell_bg(c, header_bg)
        set_cell_borders(c, color='313131', sz='4')
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_spacing(p, before=2, after=2, line=1.2)
        add_run(p, h, size=font_size, bold=True, color=WHITE)
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            c = tbl.rows[ri + 1].cells[ci]
            if ri % 2 == 1:
                set_cell_bg(c, 'F8F4F0')
            set_cell_borders(c, color='D0D0D0', sz='4')
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = c.paragraphs[0]
            set_spacing(p, before=2, after=2, line=1.25)
            add_run(p, str(val), size=font_size)
            if ci == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.runs[0].bold = True
                p.runs[0].font.color.rgb = ORANGE
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p = doc.add_paragraph()
    set_spacing(p, before=0, after=4, line=1.0)
    return tbl

def add_callout(title, body, bg='FFF4ED', bar='F47629'):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    cell = tbl.cell(0, 0)
    cell.width = Cm(17.0)
    set_cell_bg(cell, bg)
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge, c in (('top', bg), ('right', bg), ('bottom', bg), ('left', bar)):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), '36' if edge == 'left' else '4')
        el.set(qn('w:color'), c)
        tcBorders.append(el)
    tcPr.append(tcBorders)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    if title:
        p = cell.paragraphs[0]
        set_spacing(p, before=2, after=2, line=1.2)
        add_run(p, title, size=12, bold=True, color=ORANGE)
    if body:
        p2 = cell.add_paragraph()
        set_spacing(p2, before=0, after=2, line=1.3)
        add_run(p2, body, size=11, color=DARK)
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for edge, val in (('top', '120'), ('left', '200'), ('bottom', '120'), ('right', '200')):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:w'), val)
        el.set(qn('w:type'), 'dxa')
        tcMar.append(el)
    tcPr.append(tcMar)
    p = doc.add_paragraph()
    set_spacing(p, before=0, after=4, line=1.0)
    return tbl

def add_alert_box(title, body, color='red'):
    """Красный/жёлтый/зелёный блок для важных выводов."""
    color_map = {
        'red':    ('FEE2E2', 'C0392B'),
        'amber':  ('FEF3C7', 'B8860B'),
        'green':  ('D1FAE5', '166534'),
        'blue':   ('DBEAFE', '1E3A5F'),
    }
    bg, bar = color_map.get(color, color_map['red'])
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    cell = tbl.cell(0, 0)
    cell.width = Cm(17.0)
    set_cell_bg(cell, bg)
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge, c in (('top', bg), ('right', bg), ('bottom', bg), ('left', bar)):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), '36' if edge == 'left' else '4')
        el.set(qn('w:color'), c)
        tcBorders.append(el)
    tcPr.append(tcBorders)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    if title:
        p = cell.paragraphs[0]
        set_spacing(p, before=2, after=2, line=1.2)
        add_run(p, title, size=12, bold=True, color=RGBColor.from_string(bar))
    if body:
        p2 = cell.add_paragraph()
        set_spacing(p2, before=0, after=2, line=1.3)
        add_run(p2, body, size=11, color=DARK)
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for edge, val in (('top', '120'), ('left', '200'), ('bottom', '120'), ('right', '200')):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:w'), val)
        el.set(qn('w:type'), 'dxa')
        tcMar.append(el)
    tcPr.append(tcMar)
    p = doc.add_paragraph()
    set_spacing(p, before=0, after=4, line=1.0)
    return tbl

def page_break():
    p = doc.add_paragraph()
    r = p.add_run()
    r.add_break(WD_BREAK.PAGE)

def hr():
    p = doc.add_paragraph()
    set_spacing(p, before=2, after=6, line=1.0)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:color'), 'F47629')
    pBdr.append(bottom)
    pPr.append(pBdr)

print("Utils ready")

# ============================================================
# ЗАГОЛОВОК
# ============================================================
add_h1('Отчёт по переговорам', size=24, before=0, after=2, orange_bar=True)
add_para('Сеть супермаркетов «12 месяцев» (Казахстан, г. Алматы)',
         size=14, bold=True, color=DARK, after=2)
add_para('Контекст: встреча на выставке с действующим партнёром',
         size=11, italic=True, color=GREY, after=10)

hr()

# Метаданные встречи
add_table(
    headers=['Параметр', 'Значение'],
    rows=[
        ['Контрагент',          'Сеть супермаркетов «12 месяцев» (Алматы, Казахстан)'],
        ['Тип встречи',         'Переговоры на выставке с действующим партнёром'],
        ['Формат',              'Личная встреча с представителем сети'],
        ['Предмет обсуждения',  'Текущая выкладка Эмпилс (TITAN, NEWTONE, AVANTGARDE), запросы на новинки, повторный заход TITAN'],
        ['Источник данных',      'Аудиозапись → расшифровка через Whisper'],
        ['Дата встречи',        'В период проведения выставки (дата уточняется)'],
    ],
    col_widths=[5.0, 12.0],
    font_size=10
)

page_break()

# ============================================================
# КРАТКОЕ РЕЗЮМЕ
# ============================================================
add_h1('Краткое резюме', size=18, before=0, after=6)

add_para(
    'Встреча с представителем сети супермаркетов «12 месяцев» (Алматы) показала '
    'двоякую картину по сотрудничеству с Эмпилс. С одной стороны — текущие позиции '
    'Эмпилс в сети слабеют: TITAN выведен с полки (осталось 6–7 банок остатка), '
    'NEWTONE и AVANTGARDE также были убраны после неудачной попытки продажи. С другой '
    'стороны — партнёр чётко обозначил два направления, по которым есть активный '
    'ежедневный спрос покупателей и куда Эмпилс может предложить продукт: '
    'эпоксидные наливные полы («жидкий кафель») и двух/трёхкомпонентные эпоксидные '
    'краски для гаражных полов и уличных работ. Это ключевая возможность для повторного '
    'захода в сеть — уже с новой продуктовой линейкой.',
    after=8
)

# Главный вывод в красной плашке
add_alert_box('Главный вывод встречи',
    'TITAN как самостоятельный бренд в сети «12 месяцев» пока не работает — нужен повторный '
    'заход не ранее февраля 2027 года с пересмотром цены и обучением продавцов. Однако партнёр '
    'готов рассматривать принципиально новые продукты — эпоксидные полы и компонентные '
    'эпоксидные краски, по которым у него ежедневно есть покупательский спрос. Это ниша, '
    'которую в Казахстане уже активно занимают китайские поставщики.',
    color='amber')

page_break()

# ============================================================
# СОСТОЯНИЕ ПО ТЕКУЩИМ БРЕНДАМ ЭМПИЛС
# ============================================================
add_h1('Состояние по текущим брендам Эмпилс', size=18, before=0, after=6)

add_para(
    'На момент встречи в сети «12 месяцев» складывается следующая ситуация по брендам '
    'Эмпилс. Все три ключевых бренда (TITAN, NEWTONE, AVANTGARDE) на текущий момент '
    'не работают или выведены с полки. Партёр просит устранить причины и вернуться к '
    'разговору в начале 2027 года.',
    after=8
)

add_table(
    headers=['Бренд', 'Статус', 'Что произошло', 'Комментарий партнёра'],
    rows=[
        ['TITAN',
         '🔴 Выведен с полки',
         'Партнёр сам попросил убрать. Остаток 6–7 банок.',
         'Причина официально не названа. Партнёр: «Не знаю, не сказали».'],
        ['NEWTONE',
         '🔴 Не пошёл',
         'Поставили на полку, простоял полгода, продали 1 банку.',
         '«Если хороший продавец будет — любые краски купят. Полгода стоял — только одну банку продали».'],
        ['AVANTGARDE',
         '🔴 Не пошёл',
         'Аналогично NEWTONE — убран с полки после неудачной попытки.',
         'Партнёр: «Не нашли покупателя. Через державку без промоутера — нам намерение».'],
        ['ВД (водно-дисперсионные)',
         '🔴 Не представлен',
         'Ранее пробовали поставить — продажи нулевые.',
         '«В воде вы у нас вообще не продавали. Поставили — не продавали».'],
        ['Алкидные краски (ОРЕОЛ)',
         '🟢 №1 в категории',
         'Лидер продаж в алкидной категории.',
         '«Алкидные краски у нас номер один. Оптима — топ».'],
        ['ОРЕОЛ 3в1',
         '🟢 Хорошо узнаваем',
         'Бренд известен покупателю, продаётся.',
         '«Три в одном — уже знают, понятно. Выходит нормально».'],
    ],
    col_widths=[2.8, 2.5, 5.5, 6.2],
    font_size=9
)

add_alert_box('Причины неудач (анализ партнёра)',
    'Партнёр прямо назвал три причины неудач по NEWTONE / AVANTGARDE / TITAN:\n'
    '1. Слабый продавец — «если хороший продавец будет, любые краски купят».\n'
    '2. Отсутствие промоутера — продукт лежал на полке без поддержки.\n'
    '3. По TITAN — причина официально не названа, но в переговорах прозвучало, что '
    '«никакой задаток, никакой отсрочки не работаю» — то есть возможны финансовые '
    'условия или изменения в дистрибьюции.',
    color='red')

page_break()

# ============================================================
# ЖЕЛАЕМЫЕ ПРОДУКТЫ-НОВИНКИ (КЛЮЧЕВОЙ БЛОК)
# ============================================================
add_h1('Желаемые продукты-новинки', size=20, before=0, after=4)
add_para('Ключевой блок отчёта — запросы партнёра на новые продукты, по которым есть '
         'активный ежедневный покупательский спрос в Казахстане.',
         size=12, italic=True, color=GREY, after=10)

# Большая выделенная плашка — важность блока
add_alert_box('Важно: это главная коммерческая находка переговоров',
    'Партнёр впервые озвучил два направления, по которым у него в сети есть '
    'ежедневный незакрытый спрос. Сейчас этот спрос закрывают китайские поставщики, '
    'которые в 2026 году активно зашли на казахстанский рынок. Эмпилс имеет шанс '
    'занять эту нишу, если оперативно предложит продукт.',
    color='green')

# ---------- Новинка 1 ----------
add_h2('Новинка 1. Эпоксидный наливной пол («жидкий кафель»)', before=8, after=4)

add_table(
    headers=['Параметр', 'Данные'],
    rows=[
        ['Что хочет партнёр',
         'Эпоксидный наливной пол — готовый комплект в коробке (краска + отвердитель + декор).'],
        ['Спрос в магазине',
         '«Каждый день спрашивают» — постоянный покупательский спрос.'],
        ['Целевая аудитория',
         'Частные покупатели — ремонт в хрущёвках, 5–9-этажках (замена старого кафеля), кухни, гаражи.'],
        ['Объём фасовки',
         '20 кг (на 12–18 м²) — партнёр именно такую фасовку и запрашивает.'],
        ['Цена ориентир',
         '70–80 тыс. тенге за 26 кг (ориентир на китайские аналоги).'],
        ['Конкуренты',
         'Китайские поставщики (зашли в 2026 году), реализация через Ozon, Kaspi.'],
        ['Особые требования',
         'Рекламная заявка производителей: стойкость до −40 / −45 °C (уличное применение).'],
        ['Канал продаж',
         'Розница + интернет-магазин Kaspi. Готовы к паллетной выкладке и паллетной отгрузке.'],
        ['Запрос на пилот',
         'Паллета 30 штук на точку — для теста продаж.'],
    ],
    col_widths=[4.5, 12.5],
    font_size=10
)

# ---------- Новинка 2 ----------
add_h2('Новинка 2. Двух/трёхкомпонентная эпоксидная краска', before=10, after=4)

add_table(
    headers=['Параметр', 'Данные'],
    rows=[
        ['Что хочет партнёр',
         'Компонентная эпоксидная краска для полов — готовый комплект в коробке: краска + '
         'отвердитель (+ возможен третий компонент — декоративная крошка).'],
        ['Сфера применения',
         'Гаражные полы, уличные работы, наливные полы. «Машина сверху ходить — срок, ходим».'],
        ['Спрос',
         'Спрашивают большие объёмы — от 10 до 20 кг.'],
        ['Объём фасовки',
         'Большая фасовка — 10 кг / 20 кг. Маленькая уже есть у конкурентов.'],
        ['Целевая аудитория',
         'B2B-покупатели (гаражи, СТО, малый бизнес) + продвинутые DIY-клиенты.'],
        ['Конкурентная ситуация',
         '«Китайцы уже зашли — у них это есть, потребители прямо спрашивают».'],
        ['У партнёра сейчас',
         'Эпоксидная краска уже есть в продаже у других поставщиков (китайские), но партнёр '
         'не уверен в качестве — хочет проверить продукт Эмпилс.'],
        ['Запрос на пилот',
         '10 кг / 20 кг на пробу — «можно 10 сделать, попробовать».'],
    ],
    col_widths=[4.5, 12.5],
    font_size=10
)

# ---------- Сводная таблица новинок ----------
add_h2('Сводная таблица запросов по новинкам', before=10, after=4)

add_table(
    headers=['№', 'Продукт', 'Фасовка', 'Целевая цена', 'Спрос', 'Пилот'],
    rows=[
        ['1', 'Эпоксидный наливной пол («жидкий кафель»)',
         '20 кг',
         '70–80 тыс. ₸ за 26 кг',
         'Ежедневный',
         '30 шт. на точку'],
        ['2', 'Эпоксидная компонентная краска для пола',
         '10 / 20 кг',
         'уточняется',
         'Большой объём',
         '10 / 20 кг на пробу'],
    ],
    col_widths=[0.8, 6.0, 2.0, 3.0, 2.5, 3.2],
    font_size=10
)

# Главный вывод по новинкам
add_alert_box('Что делать Эмпилс с этими запросами',
    '1. Срочно проверить: есть ли в портфеле Эмпилс эпоксидный наливной пол или '
    'компонентная эпоксидная краска. По словам представителя, у Эмпилс такой продукт '
    'есть («у нас есть продукт, мы его продаём другими каналами») — но в категорию '
    'DIY-сети он пока не выводился.\n'
    '2. Подготовить образцы: паллета 30 шт. наливного пола + 10/20 кг компонентной '
    'эпоксидной краски — отправить партнёру на тест.\n'
    '3. Подготовить расчёт расхода и сравнение с китайскими аналогами по цене/качеству.\n'
    '4. Разработать промо-материалы: видеоинструкция, демо-полы в магазине (партнёр '
    'готов выделить место внутри магазина для демонстрации).',
    color='green')

page_break()

# ============================================================
# КОНКУРЕНТНОЕ ОКРУЖЕНИЕ
# ============================================================
add_h1('Конкурентное окружение в сети', size=18, before=0, after=6)

add_para(
    'В категории «лакокрасочные материалы» в сети «12 месяцев» представлена следующая '
    'картина. На полках соседствуют местные казахстанские бренды, российские бренды '
    '(включая Эмпилс) и продукция Tikkurila. В сегменте эпоксидных полов и компонентных '
    'красок активную экспансию ведут китайские поставщики.',
    after=8
)

add_table(
    headers=['Бренд', 'Статус в сети', 'Категория'],
    rows=[
        ['Радуга (Российский)',          '🟢 Представлен',                 'ЛКМ масс-сегмент'],
        ['Tikkurila',                     '🟢 Представлена',                'Премиум-ЛКМ'],
        ['Dulux',                         '🟢 Представлен',                 'Премиум-ЛКМ'],
        ['ОРЕОЛ (Эмпилс) — алкидные',    '🟢 №1 в алкидной категории',     'Алкидные эмали'],
        ['ОРЕОЛ (Эмпилс) — 3в1',         '🟢 Хорошо узнаваем',             'Грунт-эмали по ржавчине'],
        ['TITAN (Эмпилс)',                '🔴 Выведен с полки',             'Премиум-грунт-эмаль 3в1'],
        ['NEWTONE (Эмпилс)',              '🔴 Не пошёл, убран',            'ВД-краски'],
        ['AVANTGARDE (Эмпилс)',           '🔴 Не пошёл, убран',            'ВД-краски премиум'],
        ['Китайские бренды (эпоксид)',    '🟡 Активный заход в 2026',      'Эпоксидные полы и краски'],
    ],
    col_widths=[5.0, 5.5, 6.5],
    font_size=10
)

add_h2('Главная угроза: китайские поставщики эпоксидных полов', before=8, after=4)
add_para(
    'В 2026 году китайские компании активно вышли на казахстанский рынок с эпоксидными '
    'наливными полами и компонентными красками. Каналы их продаж: розница (DIY-сети, '
    'включая «12 месяцев»), Ozon, Kaspi. Цена 26 кг комплекта — 70–80 тыс. тенге. '
    'Партнёр прямо говорит: «В этом году тренд, очень много покупателей его спрашивают». '
    'Если Эмпилс не предложит свой продукт в течение 6–12 месяцев — ниша будет полностью '
    'занята китайцами и стоимость входа в категорию значительно вырастет.',
    after=8
)

page_break()

# ============================================================
# ЦЕНОВАЯ ПОЛИТИКА И УСЛОВИЯ
# ============================================================
add_h1('Ценовая политика и условия работы', size=18, before=0, after=6)

add_h2('Финансовые условия партнёра', before=4, after=4)
add_bullet('Каспи-банк — рассрочка 0% на 24 месяца, 2–3 раза в год. Главный финансовый '
           'инструмент покупателя в Казахстане.', bold=True)
add_bullet('Раньше у сети был кешбэк — временно убрали («сейчас наши мои»).')
add_bullet('Партнёр по TITAN: «никакой задаток, никакой отсрочки — не работаю». Это '
           'прямое указание на проблему с финансовыми условиями по текущему TITAN.')

add_h2('Каналы продаж в Казахстане', before=8, after=4)
add_bullet('Kaspi — крупнейший онлайн-канал (аналог支付宝/Alipay в Китае). Огромный '
           'интернет-магазин, через него идут почти все продажи.')
add_bullet('Ozon — также активный канал, особенно для китайских поставщиков эпоксидных '
           'полов (партнёр сам видел рекламу и заказывал образцы с Ozon).')
add_bullet('Розничная полка в магазинах сети — приоритет для физического присутствия '
           'бренда, импульсных покупок и доверия.')

page_break()

# ============================================================
# СРОКИ И СЛЕДУЮЩИЕ ШАГИ
# ============================================================
add_h1('Сроки и следующие шаги', size=18, before=0, after=6)

add_h2('Окно возможностей по повторному заходу TITAN', before=4, after=4)
add_alert_box('Сроки партнёра по TITAN',
    '• В этом (2026) году место на полке уже занято — повторный заход невозможен.\n'
    '• Февраль 2027 года — партнёр готов вернуться к разговору «к следующему сезону».\n'
    '• Условие повторного захода: «лучше предложение по цене и продавцы».\n'
    '• К 2027 году Эмпилс должен учесть ошибки прошлой попытки: '
    'обучить продавцов сети, предложить промо-поддержку, проверить финансовые условия.',
    color='amber')

add_h2('План действий Эмпилс (предложение)', before=10, after=4)

add_table(
    headers=['№', 'Шаг', 'Срок', 'Ответственный'],
    rows=[
        ['1', 'Проверить наличие в портфеле Эмпилс эпоксидного наливного пола и '
         'компонентной эпоксидной краски. Подготовить образцы.',
         'до 15.10.2026', 'Технический директор Эмпилс'],
        ['2', 'Подготовить пилотную партию: паллета 30 шт. наливного пола + 10/20 кг '
         'эпоксидной краски. Отправить партнёру в Алматы на тест.',
         'до 01.11.2026', 'Региональный менеджер'],
        ['3', 'Подготовить сравнительный расчёт расхода и цены vs китайские аналоги. '
         'Обосновать преимущество продукта Эмпилс.',
         'до 15.11.2026', 'Маркетинг / ТП'],
        ['4', 'Подготовить промо-материалы: видеоинструкция, демо-полы, POSm для '
         'эпоксидных продуктов.',
         'до 01.12.2026', 'Маркетинг'],
        ['5', 'Провести повторную встречу с партнёром (онлайн или в Алматы) — '
         'презентация новинок и согласование пилота.',
         'до 15.12.2026', 'Региональный менеджер'],
        ['6', 'По TITAN: подготовить улучшенное коммерческое предложение (цена + '
         'обучение продавцов + промоутер).',
         'до 01.02.2027', 'Коммерческий директор'],
        ['7', 'Повторный заход TITAN в сеть «12 месяцев» — пилот в 1–2 точках.',
         'Февраль 2027', 'Региональный менеджер'],
    ],
    col_widths=[0.8, 9.5, 3.0, 3.7],
    font_size=9
)

page_break()

# ============================================================
# ИТОГ ВСТРЕЧИ
# ============================================================
add_h1('Итог встречи', size=18, before=0, after=6)

add_para(
    'Встреча с сетью «12 месяцев» дала два важных результата. Первый — негативный: '
    'TITAN, NEWTONE и AVANTGARDE на текущий момент в сети не работают, повторный заход '
    'возможен не ранее февраля 2027 года с пересмотром цены и подготовкой продавцов. '
    'Второй — позитивный и стратегически важный: партнёр чётко обозначил два продукта, '
    'которые у него ежедневно спрашивают покупатели — эпоксидный наливной пол и '
    'компонентная эпоксидная краска. Это категория, в которую Эмпилс пока не выходил в '
    'DIY-канале Казахстана, но где у компании есть реальный продукт и реальное окно '
    'возможностей до того, как китайские поставщики полностью займут нишу.',
    after=8
)

add_alert_box('Ключевая рекомендация',
    'Сфокусироваться в ближайшие 3 месяца не на повторном заходе TITAN (окно закрыто до '
    'февраля 2027), а на выводе в сеть «12 месяцев» эпоксидных продуктов Эмпилс. '
    'Это иная категория, иной покупатель, иной чек — и здесь у Эмпилс есть преимущество '
    'первого российского производителя на казахстанском рынке в нише, которую пока '
    'делят только китайские бренды. Параллельно — готовить улучшенное предложение по '
    'TITAN к февралю 2027 года.',
    color='green')

# ---------- Финал ----------
hr()
add_para('Отчёт по переговорам  ·  Сеть супермаркетов «12 месяцев» (Казахстан)  ·  Сентябрь 2026',
         size=10, italic=True, color=GREY, align='center', after=2)
add_para('Источник: аудиозапись → расшифровка через Whisper. Анализ и структурирование — '
         'команда Эмпилс.',
         size=10, italic=True, color=GREY, align='center')

# ---------- Сохранение ----------
for table in doc.tables:
    if len(table.rows) > 0:
        tr = table.rows[0]._tr
        trPr = tr.get_or_add_trPr()
        tblHeader = OxmlElement('w:tblHeader')
        tblHeader.set(qn('w:val'), 'true')
        trPr.append(tblHeader)
        for row in table.rows:
            trPr = row._tr.get_or_add_trPr()
            cantSplit = OxmlElement('w:cantSplit')
            cantSplit.set(qn('w:val'), 'true')
            trPr.append(cantSplit)

out = '/home/z/my-project/download/Отчёт_переговоры_12_месяцев_Казахстан.docx'
doc.save(out)
print(f'OK saved: {out}')
import os
print(f'Size: {os.path.getsize(out)} bytes')
