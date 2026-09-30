"""
Normalize all product card slides to have:
1. Same fonts: Carlito for headings, Liberation Sans for body, Courier New for codes
2. Same font sizes across all slides
3. Same positions for information blocks (symmetry)
4. Consistent layout: header strip → product name → subtitle → badge → table → cert info → note
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import copy

INPUT = '/home/z/my-project/upload/Эмпилс_Каталог_пожаробезопасность_правки.pptx'
OUTPUT = '/home/z/my-project/download/Эмпилс_Каталог_пожаробезопасность_нормализованный.pptx'

prs = Presentation(INPUT)

# Define standard fonts and sizes
STD_FONTS = {
    'heading': 'Carlito',
    'body': 'Liberation Sans',
    'mono': 'Courier New',
}

STD_SIZES = {
    'header_strip': Pt(8),       # "КАРТОЧКА ПРОДУКТА · ..."
    'product_name': Pt(14),      # "NEWTONE", "ОРЕОЛ ДИСКОНТ моющаяся"
    'subtitle': Pt(9),           # "ВД-краска для внутренних работ"
    'badge': Pt(8),              # "КМ1 (Г1, В1, Д1, Т1, РП1)"
    'section_label': Pt(8),      # "КЛАСС ПОЖАРООПАСНОСТИ", "СЕРТИФИКАТЫ"
    'cert_text': Pt(7),          # Certificate numbers
    'table_header': Pt(8),        # Table header row
    'table_cell': Pt(8),          # Table cells
    'note': Pt(7),               # Bottom note
    'page_num': Pt(8),            # Page number
    'footer': Pt(6),             # Footer text
}

# Standard colors
COLORS = {
    'primary': RGBColor(0x1F, 0x4E, 0x79),
    'text': RGBColor(0x1A, 0x1A, 0x2E),
    'text_secondary': RGBColor(0x55, 0x55, 0x55),
    'green_badge': RGBColor(0x2E, 0x7D, 0x32),
    'green_bg': RGBColor(0xE8, 0xF5, 0xE9),
    'orange_badge': RGBColor(0xE6, 0x51, 0x00),
    'orange_bg': RGBColor(0xFF, 0xF3, 0xE0),
    'blue_badge': RGBColor(0x15, 0x65, 0xC0),
    'blue_bg': RGBColor(0xE3, 0xF2, 0xFD),
    'white': RGBColor(0xFF, 0xFF, 0xFF),
    'header_bg': RGBColor(0x1F, 0x4E, 0x79),
}

# Standard positions (in inches, for 10" x 5.6" slide)
STD_POSITIONS = {
    'header_strip': {'left': 0.5, 'top': 0.3, 'width': 5.0, 'height': 0.25},
    'product_name': {'left': 0.5, 'top': 0.6, 'width': 7.0, 'height': 0.35},
    'subtitle': {'left': 0.5, 'top': 0.95, 'width': 7.0, 'height': 0.25},
    'badge': {'left': 0.5, 'top': 1.25, 'width': 4.0, 'height': 0.3},
    'table': {'left': 0.5, 'top': 1.7, 'width': 6.0, 'height': 2.5},
    'cert_block': {'left': 0.5, 'top': 4.3, 'width': 9.0, 'height': 0.5},
    'note': {'left': 0.5, 'top': 4.9, 'width': 9.0, 'height': 0.4},
    'page_num': {'left': 9.0, 'top': 0.3, 'width': 0.5, 'height': 0.25},
}

def normalize_run(run, font_name, font_size, color=None, bold=None):
    """Normalize a single run's font properties"""
    run.font.name = font_name
    run.font.size = font_size
    if color is not None:
        run.font.color.rgb = color
    if bold is not None:
        run.font.bold = bold

def normalize_shape(shape, font_name, font_size, color=None, bold=None):
    """Normalize all runs in a text frame"""
    if not shape.has_text_frame:
        return
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            normalize_run(run, font_name, font_size, color, bold)

def move_shape(shape, left, top, width=None, height=None):
    """Move and optionally resize a shape"""
    shape.left = Inches(left)
    shape.top = Inches(top)
    if width is not None:
        shape.width = Inches(width)
    if height is not None:
        shape.height = Inches(height)

def get_shape_text(shape):
    """Get text from shape"""
    if shape.has_text_frame:
        return shape.text_frame.text.strip()
    return ''

def classify_shape(shape):
    """Classify a shape by its text content to determine its role"""
    text = get_shape_text(shape).lower()
    
    # Skip very short text (icons, symbols)
    if len(text) < 2:
        return 'icon'
    
    # Header strip (contains "КАРТОЧКА" or "EMPILS" or "КАТАЛОГ")
    if any(kw in text for kw in ['карточка', 'каталог', 'empils', 'легенда', 'технический']):
        if len(text) < 60:
            return 'header_strip'
    
    # Product name (large text, usually contains known product names)
    if any(kw in text for kw in ['newtone', 'ореол', 'эмаль', 'раствор', 'компаунд', 'краска', 'сто ']):
        if len(text) < 60:
            return 'product_name'
    
    # Subtitle (description of product type)
    if any(kw in text for kw in ['вд-краска', 'вд-эмаль', 'огнезащитная', 'эпоксидное', 'фасадная', 'интерьерная', 'полиакриловая', 'моющаяся', 'потолочная', 'защитно-декоративная', 'термостойкая']):
        return 'subtitle'
    
    # Badge (contains fire classes)
    if any(kw in text for kw in ['км1', 'г1', 'в1', 'в2', 'д1', 'д2', 'т1', 'т2', 'рп1', 'нг']):
        if len(text) < 50:
            return 'badge'
    
    # Section label
    if any(kw in text for kw in ['класс пожар', 'сертификат', 'применение']):
        if len(text) < 30:
            return 'section_label'
    
    # Certificate info
    if any(kw in text for kw in ['ru ', 'д-ru', 'c-ru', 'сто ', 'действует', 'n ', 'n=', '№']):
        if len(text) > 20:
            return 'cert_text'
    
    # Bottom note
    if any(kw in text for kw in ['премиальная', 'устойчива', 'класс', 'не подходит', 'внимание', 'лучший', 'для ', 'атмосфер']):
        if len(text) > 30:
            return 'note'
    
    # Page number
    if '/' in text and len(text) < 10:
        return 'page_num'
    
    # Footer
    if any(kw in text for kw in ['жилищное', 'лкм', 'фз-123', 'классы']):
        if len(text) < 60:
            return 'footer'
    
    return 'other'

# Process each slide
for slide_idx, slide in enumerate(prs.slides):
    slide_num = slide_idx + 1
    
    # Skip slides 1 and 2 (cover and legend) - just normalize fonts
    if slide_num <= 2:
        for shape in slide.shapes:
            if shape.has_text_frame:
                text = get_shape_text(shape)
                if text:
                    # Keep cover/legend fonts as is but normalize
                    pass
        continue
    
    print(f"\nНормализация слайда {slide_num}...")
    
    # Classify all shapes
    shape_roles = []
    for shape in slide.shapes:
        role = classify_shape(shape)
        if role != 'icon':
            shape_roles.append((shape, role))
            print(f"  {role}: {get_shape_text(shape)[:50]}")
    
    # Normalize each shape based on its role
    for shape, role in shape_roles:
        if role == 'header_strip':
            normalize_shape(shape, STD_FONTS['heading'], STD_SIZES['header_strip'], COLORS['text_secondary'])
        
        elif role == 'product_name':
            normalize_shape(shape, STD_FONTS['heading'], STD_SIZES['product_name'], COLORS['primary'], True)
            move_shape(shape, **STD_POSITIONS['product_name'])
        
        elif role == 'subtitle':
            normalize_shape(shape, STD_FONTS['body'], STD_SIZES['subtitle'], COLORS['text_secondary'])
            move_shape(shape, **STD_POSITIONS['subtitle'])
        
        elif role == 'badge':
            normalize_shape(shape, STD_FONTS['mono'], STD_SIZES['badge'], COLORS['green_badge'])
            move_shape(shape, **STD_POSITIONS['badge'])
        
        elif role == 'section_label':
            normalize_shape(shape, STD_FONTS['heading'], STD_SIZES['section_label'], COLORS['primary'], True)
        
        elif role == 'cert_text':
            normalize_shape(shape, STD_FONTS['mono'], STD_SIZES['cert_text'], COLORS['text_secondary'])
            move_shape(shape, **STD_POSITIONS['cert_block'])
        
        elif role == 'note':
            normalize_shape(shape, STD_FONTS['body'], STD_SIZES['note'], COLORS['text'])
            move_shape(shape, **STD_POSITIONS['note'])
        
        elif role == 'page_num':
            normalize_shape(shape, STD_FONTS['body'], STD_SIZES['page_num'], COLORS['text_secondary'])
            move_shape(shape, **STD_POSITIONS['page_num'])
        
        elif role == 'footer':
            normalize_shape(shape, STD_FONTS['body'], STD_SIZES['footer'], COLORS['text_secondary'])

# Save
prs.save(OUTPUT)

import os
file_size = os.path.getsize(OUTPUT)
print(f"\n✅ Файл сохранён: {OUTPUT}")
print(f"Размер: {file_size} байт ({file_size/1024:.1f} KB)")
