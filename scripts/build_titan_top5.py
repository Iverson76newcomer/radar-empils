"""
TITAN — Топ-5 идей вирусных роликов с медийными личностями
Лаконичный Word-документ (A4, без обложки, без оглавления).
Цветовая схема: оранжевый #F47629 / тёмный #313131 / серый #6B7280.
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ---------- Цвета ----------
ORANGE = RGBColor(0xF4, 0x76, 0x29)
DARK   = RGBColor(0x31, 0x31, 0x31)
GREY   = RGBColor(0x6B, 0x72, 0x80)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT  = RGBColor(0xFA, 0xF6, 0xF2)
BG_ALT = RGBColor(0xF8, 0xF4, 0xF0)

doc = Document()

# Базовый стиль
style = doc.styles['Normal']
style.font.name = 'Arial'
style.font.size = Pt(11)
style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')

# Поля + A4
for section in doc.sections:
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(2.2)
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)

# ---------- Утилиты ----------
def set_spacing(p, before=0, after=6, line=1.3):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line

def add_para(text, size=11, bold=False, color=None, italic=False, align='left', after=6, before=0):
    p = doc.add_paragraph()
    set_spacing(p, before=before, after=after, line=1.3)
    p.alignment = {'center': WD_ALIGN_PARAGRAPH.CENTER, 'right': WD_ALIGN_PARAGRAPH.RIGHT,
                   'left': WD_ALIGN_PARAGRAPH.LEFT, 'justify': WD_ALIGN_PARAGRAPH.JUSTIFY}.get(align, WD_ALIGN_PARAGRAPH.LEFT)
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.bold = bold
    r.italic = italic
    if color:
        r.font.color.rgb = color
    r.font.name = 'Arial'
    return p

def add_h1(text, color=DARK, size=22, before=0, after=8, orange_bar=True):
    p = doc.add_paragraph()
    set_spacing(p, before=before, after=after, line=1.1)
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(size)
    r.font.color.rgb = color
    r.font.name = 'Arial'
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
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(size)
    r.font.color.rgb = color
    r.font.name = 'Arial'
    return p

def add_h3(text, color=DARK, size=12, before=8, after=3):
    p = doc.add_paragraph()
    set_spacing(p, before=before, after=after, line=1.2)
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(size)
    r.font.color.rgb = color
    r.font.name = 'Arial'
    return p

def add_bullet(text, size=11, bold=False, indent=0):
    p = doc.add_paragraph(style='List Bullet')
    set_spacing(p, before=0, after=3, line=1.25)
    p.paragraph_format.left_indent = Cm(0.6 + indent * 0.5)
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
        r = p.add_run(h)
        r.bold = True
        r.font.size = Pt(font_size)
        r.font.color.rgb = WHITE
        r.font.name = 'Arial'
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            c = tbl.rows[ri + 1].cells[ci]
            if ri % 2 == 1:
                set_cell_bg(c, 'F8F4F0')
            set_cell_borders(c, color='D0D0D0', sz='4')
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = c.paragraphs[0]
            set_spacing(p, before=2, after=2, line=1.25)
            r = p.add_run(str(val))
            r.font.size = Pt(font_size)
            r.font.name = 'Arial'
            if ci == 0:
                r.bold = True
                r.font.color.rgb = ORANGE
            elif ci == len(row) - 1:
                r.italic = True
                r.font.color.rgb = GREY
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    # Пустой абзац после
    p = doc.add_paragraph()
    set_spacing(p, before=0, after=4, line=1.0)
    return tbl

def add_callout(title, body):
    """Цветной блок с оранжевой полосой слева."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    cell = tbl.cell(0, 0)
    cell.width = Cm(16.5)
    set_cell_bg(cell, 'FFF4ED')
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge, c in (('top', 'FFF4ED'), ('right', 'FFF4ED'), ('bottom', 'FFF4ED'), ('left', 'F47629')):
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
        r = p.add_run(title)
        r.bold = True
        r.font.size = Pt(11)
        r.font.color.rgb = ORANGE
        r.font.name = 'Arial'
    if body:
        p2 = cell.add_paragraph()
        set_spacing(p2, before=0, after=2, line=1.3)
        r2 = p2.add_run(body)
        r2.font.size = Pt(10)
        r2.font.color.rgb = DARK
        r2.font.name = 'Arial'
    # Поля ячейки
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
add_h1('TITAN — Топ-8 идей вирусных роликов', size=24, before=0, after=2, orange_bar=True)
add_para('С медийными личностями для повышения узнаваемости бренда в России',
         size=12, italic=True, color=GREY, after=4)
add_para('Цель: познакомить аудиторию выбранной медийной личности с брендом TITAN и создать '
         'видео с высоким вирусным потенциалом. ЦА бренда — мужчины 30–55 лет, DIY-покупатели, '
         'владельцы дач/гаражей/авто. Все кандидаты проверены: без статуса иноагента.',
         size=11, after=8)

hr()

# Сводная таблица
add_h2('Сводка 8 идей', before=4, after=6)
add_table(
    headers=['№', 'Личность', 'Формат', 'Целевая аудитория', 'Охват'],
    rows=[
        ['1', 'Гурам Амарян',         'Музыкальная пародия-песня',     'М 30–55, широкая',     '4,8 млн'],
        ['2', 'Тимур Батрутдинов',    'Пародия «3 поколения краски»',  'М 30–50, Comedy',      '2,3 млн'],
        ['3', 'Nail (Найл)',           'Документалка про завод Эмпилс', 'М 25–45, премиум',     '6,2 млн'],
        ['4', 'Тимофей Баженов',       'Научные испытания краски',      'М 25–50, факт-ориент.', '5,8 млн'],
        ['5', 'Антон Лапенко',         'Абсурдный скетч с Инженером',   'М+Ж 25–40, мемы',      '3,5 млн'],
        ['6', 'Дмитрий Пучков',        'Брутальный «перевод»',          'М 35–60, муж. тематика', '4,0 млн'],
        ['7', 'Гарик Харламов',        'Абсурдный скетч-челлендж',     'М+Ж 25–45, мемы',      '6,0 млн'],
        ['8', 'Игорь Лифт',            'Честный тест TITAN vs ПФ-115',  'М 30–50, DIY-ядро',    '2,5 млн'],
    ],
    col_widths=[1.0, 3.5, 5.0, 4.0, 3.0],
    font_size=10
)

page_break()

# ============================================================
# ИДЕЯ 1 — ГУРАМ АМАРЯН
# ============================================================
add_h1('Идея 1. Гурам Амарян — «Песня про забор»', size=18, before=0, after=6)

add_h3('Почему Амарян', color=ORANGE)
add_para('4,8 млн подписчиков. Жанр — вирусные музыкальные клипы-пародии с реальным бытом и '
         'приборами. Ядро аудитории — мужчины 30–55 лет, ровно наша ЦА. Клипы «Про тёщу», '
         '«Про соседа» собирают 10–20 млн просмотров. Припевы расходятся на цитаты и звучат '
         'в мужских компаниях — там, где решения о покупке ЛКМ реально принимаются.', after=8)

add_h3('Концепция ролика', color=ORANGE)
add_para('Музыкальный клип в стиле блатного шансона/кантри. Герой — мужик, который каждые 3 года '
         'перекрашивает забор ПФ-115: «Вчера опять красил, сегодня опять ржавый». Сосед покрасил '
         'забор TITANом 10 лет назад — и забыл про него. Припев: «Я крашу, крашу, крашу — у меня '
         'всё в краске, / А сосед на TITAN сел и сидит, и пьёт чаю». В финале — Амарян с банкой '
         'TITAN, нестройный хор соседей-мужиков подпевает.', after=8)

add_h3('Вирусный триггер', color=ORANGE)
add_para('Песня должна быть такой, чтобы её хотелось переслать другу: «Смотри, это про тебя!». '
         'Прилипчивый мотив + абсурдная история + цитируемая строчка из припева. Короткие нарезки '
         'по 15–30 сек для Shorts/Reels/VK Клипов.', after=8)

add_callout('Ключевая фраза',
            '«Красишь 1 раз — служит 10 лет». Это слоган бренда, заложенный в припев песни.')

page_break()

# ============================================================
# ИДЕЯ 2 — ТИМУР БАТРУТДИНОВ
# ============================================================
add_h1('Идея 2. Тимур Батрутдинов — «3 поколения краски»', size=18, before=0, after=6)

add_h3('Почему Батрутдинов', color=ORANGE)
add_para('Comedy Club, 2,3 млн подписчиков на YouTube. Мастер пародий на типажи: «Я — дед, '
         'я — отец, я — сын». Мужская аудитория 30–50 лет узнаёт в его персонажах себя, отцов, '
         'дедов. Безопасный, понятный юмор — не рисковый для бренда.', after=8)

add_h3('Концепция ролика', color=ORANGE)
add_para('Батрутдинов играет трёх мужиков одного возраста через грим:', after=4)
add_bullet('Дед — красит забор кистью с ПФ-115, плачет, ругает советскую краску.')
add_bullet('Отец — красит валиком с ГФ + ПФ-115, тоже страдает, но говорит «зато дёшево».')
add_bullet('Сын — молодой мужик с банкой TITAN, красит за 1,5 часа, потом сидит с пивом.')
add_para('Все трое встречаются у одного забора через 10 лет. Дед и Отец прибежали снова с '
         'банками — Сын сидит на скамейке, забор блестит. Финал: «И кто из нас умнее?» — '
         'обращение к камере.', after=8, before=4)

add_h3('Вирусный триггер', color=ORANGE)
add_para('Пародийный формат — то, в чём Батрутдинов силен. Расходится на мемы: «Я — дед, я — '
         'отец, я — сын». Можно развить в серию роликов с разными типажами.', after=8)

add_callout('Ключевая фраза',
            '«И кто из нас умнее?» — риторический вопрос, на который отвечает сам продукт TITAN.')

page_break()

# ============================================================
# ИДЕЯ 3 — NAIL (НАЙЛ)
# ============================================================
add_h1('Идея 3. Nail (Найл) — «Как делают краску 124 года»', size=18, before=0, after=6)

add_h3('Почему Nail', color=ORANGE)
add_para('6,2 млн подписчиков на основном канале + 4,5 млн на Nail Status. Длинные документалки '
         'про бизнесы и производства: «как устроен X», «как делают Y». Смотрят мужчины 25–45 лет, '
         'именно те, кто готов платить за премиум-продукт. Nail не делает рекламу — он делает '
         'исследование бренда, и это работает сильнее любой прямой рекламы.', after=8)

add_h3('Концепция ролика', color=ORANGE)
add_para('40-минутный лонгрид. Nail приезжает в Ростов-на-Дону на завод «Эмпилс»:', after=4)
add_bullet('История: с 1901 года, 124 года производства, пережил революцию, войну, перестройку.')
add_bullet('Современность: лаборатория, где 10 лет крутят образцы под УФ-лампой; испытательный '
           'центр «Рутил»; цех, где варят грунт-эмаль.')
add_bullet('Выезд на объект: Nail вместе с технологом красит старые ворота TITANом.')
add_bullet('Финал — Nail на фоне 10-летнего забора, покрашенного TITANом: «Сколько лет этой краске?»',
           bold=True)
add_para('', after=4)

add_h3('Вирусный триггер', color=ORANGE)
add_para('Формат «заглянуть туда, куда не пускают». Завод с историей, реальные испытания, '
         'конкретные цифры — это работает на доверие и узнаваемость одновременно. Видео такого '
         'формата живёт долго и собирает просмотры годами.', after=8)

add_callout('Ключевая фраза',
            '«124 года делают краску — и до сих пор не стыдно». Конкретный, доказательный подход, '
            'который ценит аудитория Nail.')

page_break()

# ============================================================
# ИДЕЯ 4 — ТИМОФЕЙ БАЖЕНОВ
# ============================================================
add_h1('Идея 4. Тимофей Баженов — «TITAN против стихии»', size=18, before=0, after=6)

add_h3('Почему Баженов', color=ORANGE)
add_para('Программа «Редакция» — 5,8 млн подписчиков, документальные расследования и научно-'
         'популярные фильмы. Баженов умеет делать зрелищно и научно: испытания, эксперименты, '
         'тесты на прочность. Аудитория — мужчины 25–50 лет, любопытные, ценящие факты.', after=8)

add_h3('Концепция ролика', color=ORANGE)
add_para('Не «история завода», а «испытание краски». Баженов берёт 4 образца одинаковых '
         'металлических пластин, красит 4 красками (ПФ-115, дешёвая грунт-эмаль, Dali, TITAN), '
         'и подвергает их:', after=4)
add_bullet('1000 часов под УФ-лампой (имитация 10 лет солнца).')
add_bullet('Соляной туман 30 дней (имитация приморского климата).')
add_bullet('Заморозка −40 °C / нагрев +80 °C (имитация перепадов).')
add_bullet('Механические царапины.')
add_para('В финале — показывает пластины, считает стоимость года эксплуатации каждой. TITAN '
         'выигрывает по совокупной стоимости.', after=6, before=4)
add_para('Финальный кадр: Баженов на фоне пластины, покрытой TITAN: «У коррозии нет шансов — '
         'это не маркетинг, это физика».', after=8, bold=True)

add_h3('Вирусный триггер', color=ORANGE)
add_para('Зрелищные испытания, цифры, наука. Аудитория «Редакции» любит такое и активно делится. '
         'Видео работает на доверие: «если Баженов проверил и подтвердил — значит, не врут».', after=8)

add_callout('Ключевая фраза',
            '«У коррозии нет шансов — это не маркетинг, это физика». Слоган бренда, подтверждённый '
            'экспериментом.')

page_break()

# ============================================================
# ИДЕЯ 5 — АНТОН ЛАПЕНКО
# ============================================================
add_h1('Идея 5. Антон Лапенко — «Инженер на даче»', size=18, before=0, after=6)

add_h3('Почему Лапенко', color=ORANGE)
add_para('3,5 млн подписчиков, фирменные персонажи — Инженер, Толян, Дядя Боря. Аудитория '
         'помоложе (25–40 лет), но очень лояльная и активная в распространении мемов. Его '
         'абсурдный стиль идеально подходит для вирусного ролика в соцсетях.', after=8)

add_h3('Концепция ролика', color=ORANGE)
add_para('Герой Лапенко — Инженер — приезжает на дачу к Дяде Боре. Дядя Боря жалуется: «Ворота '
         'ржавые, опять красить надо, в прошлый раз красил в 2014-м». Инженер достаёт банку '
         'TITAN и начинает вещать с энтузиазмом инженера: «Двойной барьерный эффект! '
         'Модифицированный алкид! От −40 до +80!» — пока красит. Дядя Боря слушает, кивает, в '
         'финале говорит: «Серёга, ты главное не размахивайся, а то опять на 10 лет забудешь». '
         'Через 10 лет Инженер возвращается — ворота блестят, он торжественно: «Я же говорил!».', after=8)

add_h3('Вирусный триггер', color=ORANGE)
add_para('Лапенко — мастер мемов. Фраза «ты главное не размахивайся» или «опять на 10 лет '
         'забудешь» может уйти в народ. Короткие нарезки с Инженером отлично пойдут в '
         'Shorts/Reels/VK Клипов.', after=8)

add_callout('Ключевая фраза',
            '«Серёга, ты главное не размахивайся — а то опять на 10 лет забудешь». Абсурдная '
            'фраза, которая через юмор фиксирует преимущество «10 лет без перекраски».')

page_break()

# ============================================================
# ИДЕЯ 6 — ДМИТРИЙ ПУЧКОВ
# ============================================================
add_h1('Идея 6. Дмитрий Пучков — «У коррозии нет шансов»', size=18, before=0, after=6)

add_h3('Почему Пучков', color=ORANGE)
add_para('4 млн подписчиков, ядро аудитории — мужчины 35–60 лет, любители оружия, техники, '
         'брутальной темы. Ценит качество, не переносит «дешёвую рекламу». Если бренд попадает '
         'в его тон — она работает как рекомендация «своего парня».', after=8)

add_h3('Концепция ролика', color=ORANGE)
add_para('Видео в фирменном стиле «Гоблин переводит» — Пучков голосом за кадром озвучивает '
         '«перевод» реального боя: Мужик vs Ржавчина. Ржавчина наступает на гараж, на ворота, '
         'на забор — как армия врага. Мужик в растерянности хватает ПФ-115, грунтовку, '
         'преобразователь — ничего не помогает, всё смывается дождём. Тогда появляется банка '
         'TITAN — как герой боевика. Пучков: «И вот тут началось настоящее кино». Финал — через '
         '10 лет тот же мужик с биноклем пытается найти ржавчину на воротах — не находит.', after=8)

add_h3('Вирусный триггер', color=ORANGE)
add_para('Брутальный юмор, узнаваемый голос, отсылки к классике боевиков. Фраза «У коррозии нет '
         'шансов» — слоган бренда и одновременно каламбур из контекста ролика. Аудитория Пучкова '
         'любит делиться такими видео в Telegram-каналах «мужской тематики».', after=8)

add_callout('Ключевая фраза',
            '«У коррозии нет шансов» — слоган бренда, произнесённый фирменным голосом Пучкова '
            'как финальная реплика боевика.')

page_break()

# ============================================================
# ИДЕЯ 7 — ГАРИК ХАРЛАМОВ
# ============================================================
add_h1('Идея 7. Гарик Харламов — «Я покрасил всё»', size=18, before=0, after=6)

add_h3('Почему Харламов', color=ORANGE)
add_para('6 млн подписчиков на YouTube, 5 млн в Instagram. Делает абсурдные скетчи, которые '
         'расходятся на мемы. Аудитория — мужчины и женщины 25–45 лет, широкий охват, не только '
         'DIY-ядро. Идеально подходит для бренда, который хочет выйти за пределы «только мужской» '
         'аудитории.', after=8)

add_h3('Концепция ролика', color=ORANGE)
add_para('Скетч-челлендж в стиле «Один день из жизни Харламова». Он находит банку TITAN у себя '
         'в гараже — и решает, что это «магическая банка, которая красит всё». Сначала красит '
         'гаражные ворота (нормально). Потом — любимый автомобиль жены (она в шоке). Потом — '
         'диван в гостиной (она кричит). Потом — кота (тот убегает, оставляя оранжевые следы — '
         'очевидно, кота не красят, это CGI или шутка). В финале — жена в ярости, а Харламов '
         'стоит на фоне идеально покрашенного гаража, держит банку: «Но ворота-то — 10 лет будут '
         'как новые!». Жена, скрестив руки: «Десять лет ты здесь и будешь жить».', after=8)

add_h3('Вирусный триггер', color=ORANGE)
add_para('Харламов — мастер абсурда и мемов. Фразы вроде «Но ворота-то!» могут уйти в народ. '
         'Короткие нарезки на 15–30 секунд отлично пойдут в Reels/TikTok/Shorts. Главный приём '
         '— довести ситуацию до абсурда, а потом резюмировать реальным преимуществом продукта.', after=8)

add_callout('Ключевая фраза',
            '«Но ворота-то — 10 лет будут как новые!» — абсурдная ситуация сводится к реальному '
            'УТП бренда, и зритель запоминает «10 лет».')

page_break()

# ============================================================
# ИДЕЯ 8 — ИГОРЬ ЛИФТ
# ============================================================
add_h1('Идея 8. Игорь Лифт — «Битва красок: TITAN vs ПФ-115»', size=18, before=0, after=6)

add_h3('Почему Лифт', color=ORANGE)
add_para('2,5 млн подписчиков, ядро аудитории — мужчины 30–50 лет, делают ремонт сами. Это '
         'точечное попадание в ЦА: человек, который смотрит Лифта, через неделю идёт в '
         'Лемана/Петрович и покупает краску. Меньший охват, чем у Дудя или Харламова — но '
         'максимальная конверсия в покупку.', after=8)

add_h3('Концепция ролика', color=ORANGE)
add_para('Челлендж-формат: «Покрасил одни ворота двумя красками — и забыл на год». Лифт берёт '
         'старые ржавые ворота, делит их пополам. Левую половину красит комплектом «ГФ + '
         'ПФ-115» (как привыкли все). Правую — одной банкой TITAN 3в1. Снимает таймлапс: сколько '
         'времени ушло, сколько слоёв, сколько денег. Через 6 месяцев возвращается с камерой — '
         'сравнивает. Через год — снова. В финале — конкретный расчёт:', after=6)
add_bullet('Левая половина: 3 перекраски, 8 250 ₽.')
add_bullet('Правая половина: 1 покраска, 2 900 ₽.')
add_bullet('Экономия — 5 350 ₽ и 4 выходных.', bold=True)

add_h3('Вирусный триггер', color=ORANGE)
add_para('Формат «честный тест» — самый доверенный в строительной нише. Не реклама, а '
         'эксперимент. Видео такого формата смотрят перед походом в магазин — и решение о '
         'покупке TITAN принимается прямо там. Дополнительно: короткие нарезки с ключевыми '
         'моментами (через 6 месяцев / через год) для Shorts/Reels.', after=8)

add_callout('Ключевая фраза',
            '«Экономия — 5 350 ₽ и 4 выходных». Конкретные цифры вместо абстрактных обещаний — '
            'это то, что запоминается и работает на конверсию в покупку.')

page_break()

# ============================================================
# РЕКОМЕНДАЦИЯ
# ============================================================
add_h1('Рекомендация: связка из 3 роликов', size=18, before=0, after=6)

add_para('Не делайте одну идею — сделайте 3 параллельно для разных целей. Эта связка работает '
         'как воронка: от узнаваемости к доверию к покупке.', after=8)

add_table(
    headers=['Этап воронки', 'Личность', 'Задача', 'Ожидаемый эффект'],
    rows=[
        ['Узнаваемость',  'Гурам Амарян',     'Запустить слоган в народ', 'Охват 10+ млн просмотров'],
        ['Доверие',       'Nail (Найл)',       'Премиальное позиционирование', 'Долгоживущий контент, доверие к бренду'],
        ['Конверсия',     'Игорь Лифт',        'Честный тест TITAN vs ПФ-115', 'Решение о покупке перед походом в магазин'],
    ],
    col_widths=[3.0, 4.0, 5.0, 4.5],
    font_size=10
)

add_para('Логика воронки:', bold=True, before=8, after=4)
add_bullet('Амарян — широкая аудитория видит песню, запоминает слоган «Красишь 1 раз — служит '
           '10 лет», возникает интерес к бренду.')
add_bullet('Nail — заинтересованные смотрят 40-минутный лонгрид, узнают про завод Эмпилс, '
           'испытания, историю. Формируется доверие.')
add_bullet('Лифт — те, кто готов купить, смотрят честный тест перед походом в Лемана/Петрович. '
           'Решение о покупке принимается на основе цифр.')

add_callout('Итог',
            'Связка Амарян → Nail → Лифт закрывает все три этапа потребительского пути: '
            'узнавание → доверие → покупка. Запускать ролики рекомендуется с разницей в '
            '2–3 недели для создания эффекта «везде вижу TITAN».')

# ---------- Финал ----------
hr()
add_para('TITAN  ·  Топ-8 идей вирусных роликов  ·  Версия 1.1  ·  Август 2026',
         size=10, italic=True, color=GREY, align='center', after=2)
add_para('Документ подготовлен как концептуальный материал для PR-стратегии бренда TITAN.',
         size=10, italic=True, color=GREY, align='center')

# ---------- Сохранение ----------
# Повтор шапок таблиц + cantSplit
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

out = '/home/z/my-project/download/TITAN_Топ-8_идей_вирусных_роликов.docx'
doc.save(out)
print(f'OK saved: {out}')
import os
print(f'Size: {os.path.getsize(out)} bytes')
