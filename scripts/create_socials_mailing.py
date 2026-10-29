#!/usr/bin/env python3
"""
Создание итогового изображения для рассылки Empils с призывом
вступить в социальные сети.
Фон: предоставленный PDF (конвертированный в PNG).
Текст: обращение к партнёрам + 3 QR-кода (VK, Telegram, MAX).
"""

import os
import qrcode
from PIL import Image, ImageDraw, ImageFont, ImageOps

# === ПУТИ ===
BG_PATH = '/tmp/background_original-1.png'
OUT_DIR = '/home/z/my-project/download'
os.makedirs(OUT_DIR, exist_ok=True)
OUT_PNG = os.path.join(OUT_DIR, 'Рассылка_Empils_соцсети.png')
OUT_PDF = os.path.join(OUT_DIR, 'Рассылка_Empils_соцсети.pdf')

# === ЦВЕТА (из брендбука Empils — чёрно-белая палитра) ===
BLACK = (26, 31, 54)        # тёмный текст
DARK = (40, 40, 40)         # подписи
GRAY = (110, 110, 110)      # мелкий текст
LIGHT_GRAY = (180, 180, 180)
WHITE = (255, 255, 255)
ACCENT_RED = (198, 40, 40)  # для «Подписывайтесь!»
EMPILS_ORANGE = (244, 121, 32)  # фирменный оранжевый Empils для заголовка

# === ШРИФТЫ ===
FONT_PATH_BOLD = '/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf'
FONT_PATH_REG = '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf'
FONT_PATH_ITALIC = '/usr/share/fonts/truetype/liberation/LiberationSans-Italic.ttf'

def font(size, bold=False, italic=False):
    if bold:
        return ImageFont.truetype(FONT_PATH_BOLD, size)
    if italic:
        return ImageFont.truetype(FONT_PATH_ITALIC, size)
    return ImageFont.truetype(FONT_PATH_REG, size)

# === СОЦСЕТИ ===
SOCIALS = [
    {
        'name': 'ВКонтакте',
        'url': 'https://vk.com/empils_brands',
        'handle': '@empils_brands',
        'color': (0, 119, 255),  # синий VK
    },
    {
        'name': 'Telegram',
        'url': 'https://t.me/empils_lkm',
        'handle': '@empils_lkm',
        'color': (0, 136, 204),  # голубой Telegram
    },
    {
        'name': 'МАХ',
        'url': 'https://max.ru/id6167008343_biz',
        'handle': 'max.ru/id6167008343_biz',
        'color': (50, 50, 50),  # тёмный МАХ
    },
]


def make_qr(url, box_size=8, border=1):
    """Сгенерировать QR-код и вернуть PIL Image."""
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=box_size,
        border=border,
    )
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color=BLACK, back_color=WHITE).convert('RGB')
    return img


def draw_text_centered(draw, cx, y, text, fnt, color):
    """Нарисовать текст по центру (по горизонтали)."""
    bbox = draw.textbbox((0, 0), text, font=fnt)
    w = bbox[2] - bbox[0]
    draw.text((cx - w // 2, y), text, font=fnt, fill=color)
    return bbox[3] - bbox[1]  # высота


def draw_text_wrapped(draw, x, y, text, fnt, color, max_width):
    """Перенос текста по словам с возвратом новой y."""
    words = text.split()
    line = ''
    line_height = fnt.size + 4
    for word in words:
        test_line = (line + ' ' + word).strip()
        bbox = draw.textbbox((0, 0), test_line, font=fnt)
        if bbox[2] - bbox[0] <= max_width:
            line = test_line
        else:
            if line:
                draw.text((x, y), line, font=fnt, fill=color)
                y += line_height
            line = word
    if line:
        draw.text((x, y), line, font=fnt, fill=color)
        y += line_height
    return y


def main():
    # === ОТКРЫВАЕМ ФОН ===
    bg = Image.open(BG_PATH).convert('RGB')
    W, H = bg.size  # 1241 x 1754
    print(f'Фон: {W}x{H}')

    # Создаём копию для рисования
    img = bg.copy()
    draw = ImageDraw.Draw(img)

    # === РАЗМЕТКА ===
    # Фон имеет:
    #   - узор из точек вверху (y=0..350)
    #   - логотип Empils (y=400..550) — НЕ ТРОГАТЬ!
    #   - чистая зона (y=600..1450) — 850 px для текста
    #   - список брендов внизу (y=1500..1700)

    CX = W // 2
    top_y = 660  # НАЧИНАЕМ НИЖЕ логотипа Empils (было 620 — опустили ещё на 40 px)
    bottom_y = 1500  # до списка брендов (фактически бренды начинаются на y=1510)

    # === 1. ЗАГОЛОВОК "Уважаемые партнёры!" ===
    y = top_y
    fnt_title = font(48, bold=True)
    draw_text_centered(draw, CX, y, 'Уважаемые партнёры!', fnt_title, EMPILS_ORANGE)
    y += 80  # отступ после заголовка

    # === 2. ОСНОВНОЙ ТЕКСТ (сокращённый, по центру, с переносом) ===
    max_text_width = W - 200  # отступы по 100 px слева/справа

    # Параграф 1 — макс. сокращено
    fnt_body = font(26)
    text1 = ('Мы ценим наше сотрудничество и приглашаем вас '
             'стать ближе к жизни компании.')
    y = draw_text_wrapped(draw, 100, y, text1, fnt_body, DARK, max_text_width)
    y += 20  # отступ между абзацами

    # Параграф 2 — макс. сокращено
    text2 = ('Присоединяйтесь к нам в социальных сетях, '
             'чтобы быть в курсе событий.')
    y = draw_text_wrapped(draw, 100, y, text2, fnt_body, DARK, max_text_width)
    y += 60  # большой отступ перед разделителем

    # === 3. РАЗДЕЛИТЕЛЬНАЯ ПОЛОСА ===
    line_y = y
    draw.line([(W // 2 - 80, line_y), (W // 2 + 80, line_y)],
              fill=LIGHT_GRAY, width=2)
    y += 45  # отступ после разделителя

    # === 4. ПРИЗЫВ "Подписывайтесь!" ===
    fnt_cta = font(38, bold=True)
    draw_text_centered(draw, CX, y, 'Подписывайтесь!', fnt_cta, EMPILS_ORANGE)
    y += 50  # отступ

    fnt_cta2 = font(22, italic=True)
    draw_text_centered(draw, CX, y, 'Будем рады видеть вас в нашем кругу', fnt_cta2, GRAY)
    y += 50  # отступ перед заголовком QR

    # === 5. ТРИ QR-КОДА В РЯД ===
    # Размеры QR — уменьшены до 180 px (было 200)
    qr_size = 180  # px
    qr_gap = 90    # отступ между QR (увеличен с 50 до 90)
    total_w = qr_size * 3 + qr_gap * 2
    start_x = (W - total_w) // 2

    # Заголовок над QR-кодами
    fnt_qr_title = font(22, bold=True)
    draw_text_centered(draw, CX, y, 'НАШИ СОЦИАЛЬНЫЕ СЕТИ', fnt_qr_title, BLACK)
    y += 30  # отступ после заголовка (стандартный)

    # Опускаем сам QR-блок ниже на 5% высоты изображения
    qr_drop = int(H * 0.05)  # 5% от высоты изображения (~88 px)
    # Ограничиваем, чтобы не наехать на бренды (мин. зазор 30px до брендов на y=1510)
    max_qr_drop = 1510 - (y + qr_size + 30)
    qr_drop = min(qr_drop, max_qr_drop)
    print(f'Опускаем QR-блок на {qr_drop} px (5% = {int(H * 0.05)} px)')

    qr_y = y + qr_drop  # ОПУСКАЕМ QR-блок
    label_y = qr_y + qr_size + 28  # отступ для подписи (увеличен с 12 до 28)

    for i, social in enumerate(SOCIALS):
        x_qr = start_x + (qr_size + qr_gap) * i

        # Генерируем QR
        qr_img = make_qr(social['url'], box_size=8, border=1)
        qr_img = qr_img.resize((qr_size, qr_size), Image.LANCZOS)

        # Рамка вокруг QR
        border_padding = 12
        border_size = qr_size + border_padding * 2
        border_x = x_qr - border_padding
        border_y = qr_y - border_padding

        # Белая подложка с тёмной рамкой
        draw.rectangle(
            [border_x - 2, border_y - 2, border_x + border_size + 2, border_y + border_size + 2],
            outline=DARK, width=2, fill=WHITE
        )

        # Вставляем QR
        img.paste(qr_img, (x_qr, qr_y))

        # Цветная полоска сверху рамки (цвет соцсети)
        draw.rectangle(
            [border_x - 2, border_y - 8, border_x + border_size + 2, border_y - 2],
            fill=social['color']
        )

        # Название соцсети (крупно, под QR)
        fnt_name = font(22, bold=True)
        bbox = draw.textbbox((0, 0), social['name'], font=fnt_name)
        name_w = bbox[2] - bbox[0]
        draw.text((x_qr + (qr_size - name_w) // 2, label_y),
                  social['name'], font=fnt_name, fill=BLACK)

        # Handle (мелким)
        fnt_handle = font(16)
        bbox = draw.textbbox((0, 0), social['handle'], font=fnt_handle)
        handle_w = bbox[2] - bbox[0]
        draw.text((x_qr + (qr_size - handle_w) // 2, label_y + 30),
                  social['handle'], font=fnt_handle, fill=GRAY)

    # === 6. СОХРАНЕНИЕ ===
    img.save(OUT_PNG, 'PNG', dpi=(300, 300))
    print(f'PNG сохранён: {OUT_PNG}')

    # Конвертируем в PDF
    img_rgb = img.convert('RGB')
    img_rgb.save(OUT_PDF, 'PDF', resolution=300.0)
    print(f'PDF сохранён: {OUT_PDF}')

    # Проверка: не выходит ли QR-блок за пределы рабочей зоны
    qr_block_bottom = label_y + 70
    print(f'Низ QR-блока: y={qr_block_bottom} (лимит: {bottom_y})')
    if qr_block_bottom > bottom_y:
        print(f'⚠️  ПРЕДУПРЕЖДЕНИЕ: QR-блок выходит за пределы рабочей зоны!')
    else:
        print(f'✅ QR-блок укладывается в рабочую зону (запас {bottom_y - qr_block_bottom}px)')


if __name__ == '__main__':
    main()
