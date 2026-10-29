#!/usr/bin/env python3
"""
Унификация визуального стиля дорожной карты TITAN.
1. Цвета шапок по единой логике
2. Шрифт Roboto Mono → Arial (слайд 1)
3. Нумерация этапов
4. Шапка на слайд 9
5. Футеры + жёлтые полосы
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import copy

INPUT = '/home/z/my-project/upload/Titan -Дорожная_карта_fin.pptx'
OUTPUT = '/home/z/my-project/download/Titan_Дорожная_карта_fin.pptx'

# === ЦВЕТА ===
RED = RGBColor(0xC6, 0x28, 0x28)
DARK_RED = RGBColor(0x8E, 0x00, 0x00)
GREEN = RGBColor(0x16, 0xA3, 0x4A)
DARK_GREEN = RGBColor(0x0B, 0x6E, 0x31)
DARK = RGBColor(0x1A, 0x1F, 0x36)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GRAY = RGBColor(0x9C, 0xA3, 0xAF)
LIGHT_GRAY = RGBColor(0xE5, 0xE7, 0xEB)
ORANGE = RGBColor(0xF4, 0x79, 0x20)
YELLOW = RGBColor(0xFB, 0xBF, 0x24)
DARK_BLUE = RGBColor(0x1E, 0x3A, 0x5F)
BLUE = RGBColor(0x25, 0x63, 0xEB)

# Светлые варианты (которые нужно заменить на насыщенные)
LIGHT_BLUE = RGBColor(0xDB, 0xEA, 0xFE)
LIGHT_RED = RGBColor(0xFE, 0xE2, 0xE2)
LIGHT_ORANGE = RGBColor(0xFF, 0xED, 0xD5)
LIGHT_GREEN = RGBColor(0xDC, 0xFC, 0xE7)

# Карта целевых цветов шапок по слайдам
SLIDE_COLORS = {
    1: None,         # Титул — не трогаем
    2: DARK_BLUE,    # Этап 0 — поиск партнёра
    3: RED,          # Этап 1 — первичные
    4: ORANGE,       # Этап 2 — вторичные
    5: GREEN,        # Этап 3 — третичные
    6: GREEN,        # POS-материалы (детализация этапа 3)
    7: GREEN,        # Контент (детализация этапа 3)
    8: GREEN,        # Outdoor (детализация этапа 3)
    9: RED,          # Планограммы — бренд TITAN
    10: DARK_BLUE,   # MML сравнение — аналитика
    11: GREEN,       # Заработок — деньги = позитив
    12: DARK_BLUE,   # Детальный расчёт — аналитика
    13: DARK_BLUE,   # Обучение ТП — процесс
    14: DARK_BLUE,   # План Б — управление
    15: DARK_BLUE,   # KPI — контроль
}

# Карта замены нумерации этапов
STAGE_REPLACEMENTS = {
    'ЭТАП 0 – ПОИСК': 'ЭТАП 0 — ПОИСК НОВОГО ПАРТНЁРА',
    'ЭТАП 0 -ПОИСК': 'ЭТАП 0 — ПОИСК НОВОГО ПАРТНЁРА',
    'ЭТАП 4 — POS-МАТЕРИАЛЫ': 'ЭТАП 3 — POS-МАТЕРИАЛЫ НА ПОЛКЕ',
    'ЭТАП 4 -POS': 'ЭТАП 3 — POS-МАТЕРИАЛЫ НА ПОЛКЕ',
    'ЭТАП 4 - POS': 'ЭТАП 3 — POS-МАТЕРИАЛЫ НА ПОЛКЕ',
    'ЭТАП 5 — КОНТЕНТ': 'ЭТАП 3 — КОНТЕНТ И ОБУЧЕНИЕ',
    'ЭТАП 5 -КОНТЕНТ': 'ЭТАП 3 — КОНТЕНТ И ОБУЧЕНИЕ',
    'ЭТАП 5 - КОНТЕНТ': 'ЭТАП 3 — КОНТЕНТ И ОБУЧЕНИЕ',
    'ЭТАП 6 -ПРОДВИЖЕНИЕ': 'ЭТАП 3 — ПРОДВИЖЕНИЕ (OUTDOOR + DIGITAL)',
    'ЭТАП 6 - ПРОДВИЖЕНИЕ': 'ЭТАП 3 — ПРОДВИЖЕНИЕ (OUTDOOR + DIGITAL)',
    'ЭТАП 6 —ПРОДВИЖЕНИЕ': 'ЭТАП 3 — ПРОДВИЖЕНИЕ (OUTDOOR + DIGITAL)',
}

# Светлые цвета, которые меняем на насыщенные
COLOR_REPLACEMENTS = {
    str(LIGHT_BLUE): DARK_BLUE,
    str(LIGHT_RED): RED,
    str(LIGHT_ORANGE): ORANGE,
    str(LIGHT_GREEN): GREEN,
    'DBEAFE': DARK_BLUE,
    'FEE2E2': RED,
    'FFEDD5': ORANGE,
    'DCFCE7': GREEN,
}


def replace_text_in_shape(shape, replacements):
    """Заменить текст в фигуре по словарю замен."""
    if not shape.has_text_frame:
        return
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            for old_text, new_text in replacements.items():
                if old_text in run.text:
                    run.text = run.text.replace(old_text, new_text)


def fix_shape_color(shape, target_color):
    """Заменить цвет заливки фигуры на целевой."""
    if shape.fill.type is not None:
        try:
            current = str(shape.fill.fore_color.rgb)
            # Если цвет светлый — меняем на насыщенный
            if current in COLOR_REPLACEMENTS:
                shape.fill.fore_color.rgb = COLOR_REPLACEMENTS[current]
            # Если цвет жёлтый (#FBBF24) и целевой не жёлтый — меняем
            elif current == 'FBBF24' and target_color != YELLOW:
                shape.fill.fore_color.rgb = target_color
        except:
            pass


def fix_header_colors(slide, slide_num):
    """Исправить цвета шапки слайда."""
    target_color = SLIDE_COLORS.get(slide_num)
    if target_color is None:
        return
    
    for shape in slide.shapes:
        # Шапка — прямоугольник в верхней части (y < 1 inch)
        if shape.top is not None and shape.top < 914400:
            if shape.shape_type in [MSO_SHAPE.RECTANGLE, MSO_SHAPE.ROUNDED_RECTANGLE]:
                fix_shape_color(shape, target_color)


def fix_fonts(slide, slide_num):
    """Заменить Roboto Mono на Arial."""
    if slide_num != 1:
        return
    
    for shape in slide.shapes:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if run.font.name and 'Roboto' in run.font.name:
                        run.font.name = 'Arial'


def fix_stage_numbering(slide, slide_num):
    """Исправить нумерацию этапов."""
    if slide_num < 2 or slide_num > 8:
        return
    
    for shape in slide.shapes:
        replace_text_in_shape(shape, STAGE_REPLACEMENTS)


def add_header_to_slide_9(slide):
    """Добавить шапку на слайд 9 (планограммы), если её нет."""
    # Проверим — есть ли уже шапка
    has_header = False
    for shape in slide.shapes:
        if shape.top is not None and shape.top < 914400:
            if shape.shape_type in [MSO_SHAPE.RECTANGLE, MSO_SHAPE.ROUNDED_RECTANGLE]:
                has_header = True
                break
    
    if has_header:
        return
    
    # Добавляем шапку
    slide_w = Inches(13.333)
    header = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, slide_w, Inches(0.65))
    header.fill.solid()
    header.fill.fore_color.rgb = RED
    header.line.fill.background()
    header.shadow.inherit = False
    
    # Жёлтая полоса
    yellow_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(0.65), slide_w, Inches(0.04))
    yellow_bar.fill.solid()
    yellow_bar.fill.fore_color.rgb = YELLOW
    yellow_bar.line.fill.background()
    yellow_bar.shadow.inherit = False
    
    # Заголовок
    tb = slide.shapes.add_textbox(Inches(0.4), Inches(0.08), Inches(8.0), Inches(0.5))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Pt(2)
    tf.margin_right = Pt(2)
    tf.margin_top = Pt(2)
    tf.margin_bottom = Pt(2)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = 'ПЛАНОГРАММЫ ПО КЛАССАМ ТТ'
    run.font.size = Pt(22)
    run.font.color.rgb = WHITE
    run.font.bold = True
    run.font.name = 'Arial'


def unify_footers(slide):
    """Унифицировать футеры."""
    FOOTER_TEXT = 'Empils  •  TITAN  •  Дорожная карта развития продаж  •  Красишь 1 раз — служит 10 лет'
    
    # Найдём существующий футер и обновим текст
    for shape in slide.shapes:
        if shape.has_text_frame:
            if shape.top is not None and shape.top > 5943600:  # > 6.5 inch
                tf = shape.text_frame
                # Очистим и заполним единым текстом
                if tf.paragraphs:
                    p = tf.paragraphs[0]
                    # Очистим существующие runs
                    for run in p.runs:
                        run.text = ''
                    if p.runs:
                        p.runs[0].text = FOOTER_TEXT
                        p.runs[0].font.size = Pt(8)
                        p.runs[0].font.color.rgb = GRAY
                        p.runs[0].font.name = 'Arial'
                        p.runs[0].font.italic = True
                    else:
                        r = p.add_run()
                        r.text = FOOTER_TEXT
                        r.font.size = Pt(8)
                        r.font.color.rgb = GRAY
                        r.font.name = 'Arial'
                        r.font.italic = True
                    p.alignment = PP_ALIGN.CENTER


def ensure_yellow_bar(slide, slide_num):
    """Убедиться, что жёлтая полоса есть под шапкой."""
    if slide_num == 1:
        return  # Титул — не трогаем
    
    # Проверим — есть ли жёлтая полоса
    has_yellow = False
    for shape in slide.shapes:
        if shape.top is not None and 600000 < shape.top < 800000:  # ~0.65-0.87 inch
            if shape.shape_type == MSO_SHAPE.RECTANGLE:
                try:
                    if shape.fill.type is not None:
                        color = str(shape.fill.fore_color.rgb)
                        if color == 'FBBF24' or color == 'FBC02D':
                            has_yellow = True
                except:
                    pass
    
    if not has_yellow:
        # Добавляем жёлтую полосу
        slide_w = Inches(13.333)
        yellow_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(0.65), slide_w, Inches(0.04))
        yellow_bar.fill.solid()
        yellow_bar.fill.fore_color.rgb = YELLOW
        yellow_bar.line.fill.background()
        yellow_bar.shadow.inherit = False


def main():
    prs = Presentation(INPUT)
    print(f'Загружено: {INPUT}')
    print(f'Слайдов: {len(prs.slides)}')
    print()
    
    for i, slide in enumerate(prs.slides, 1):
        print(f'Слайд {i}...', end=' ')
        
        # 1. Цвета шапок
        if i != 1:
            fix_header_colors(slide, i)
        
        # 2. Шрифты (слайд 1)
        fix_fonts(slide, i)
        
        # 3. Нумерация этапов
        fix_stage_numbering(slide, i)
        
        # 4. Шапка на слайд 9
        if i == 9:
            add_header_to_slide_9(slide)
        
        # 5. Жёлтая полоса
        ensure_yellow_bar(slide, i)
        
        # 6. Футеры
        unify_footers(slide)
        
        print('OK')
    
    prs.save(OUTPUT)
    print(f'\nСохранено: {OUTPUT}')


if __name__ == '__main__':
    main()
