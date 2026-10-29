#!/usr/bin/env python3
"""
Создание мини-визуализаций для 9 инструментов третичных продаж TITAN.
Каждая визуализация — PNG файл, который будет встроен в презентацию.
"""

import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.font_manager as fm
fm.fontManager.addfont('/usr/share/fonts/truetype/chinese/LiberationSans-Regular.ttf')
fm.fontManager.addfont('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, Rectangle, Circle, Polygon, FancyArrowPatch
import numpy as np

plt.rcParams['font.sans-serif'] = ['Liberation Sans', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

OUT_DIR = '/home/z/my-project/download/titan_posm_visuals'
os.makedirs(OUT_DIR, exist_ok=True)

RED = '#C62828'
DARK_RED = '#8E0000'
GREEN = '#16A34A'
DARK = '#1A1F36'
WHITE = '#FFFFFF'
GRAY = '#9CA3AF'
LIGHT_GRAY = '#F5F5F5'
ORANGE = '#F47920'
YELLOW = '#FBBF24'
BLUE = '#2563EB'
BLACK = '#000000'
METAL = '#C0C0C0'


def save_fig(fig, name):
    path = os.path.join(OUT_DIR, name)
    fig.savefig(path, dpi=150, bbox_inches='tight', facecolor=WHITE)
    plt.close(fig)
    print(f'OK: {name}')
    return path


# ==============================================================
# 1. ШЕЛФТОКЕР
# ==============================================================
def make_shelftalker():
    fig, ax = plt.subplots(figsize=(8, 4), constrained_layout=True)
    ax.set_xlim(0, 8)
    ax.set_ylim(0, 4)
    ax.set_aspect('equal')
    ax.axis('off')

    # Основная плашка (красная, наклонная — как настоящий шелфтокер)
    talker = Polygon([(0.5, 0.5), (7.5, 0.5), (7.5, 3.0), (4.0, 3.5), (0.5, 3.0)],
                     closed=True, facecolor=RED, edgecolor=DARK_RED, linewidth=2)
    ax.add_patch(talker)

    # Слоган
    ax.text(4, 2.4, '10 лет без перекраски', ha='center', va='center',
            fontsize=18, color=WHITE, fontweight='bold')
    ax.text(4, 1.7, '10 лет экономии!', ha='center', va='center',
            fontsize=18, color=YELLOW, fontweight='bold')

    # QR-код (имитация)
    np.random.seed(42)
    qr_data = np.random.randint(0, 2, (5, 5))
    for i in range(5):
        for j in range(5):
            if qr_data[i, j]:
                ax.add_patch(Rectangle((6.5 + j*0.12, 0.7 + i*0.12), 0.12, 0.12,
                                       facecolor=WHITE, edgecolor='none'))
    ax.text(7.0, 0.55, 'QR', ha='center', fontsize=7, color=WHITE)

    # Полка (линия снизу)
    ax.plot([0, 8], [0.3, 0.3], color='#8B6F47', linewidth=4)

    save_fig(fig, '01_shelftalker.png')


# ==============================================================
# 2. ВОБЛЕР
# ==============================================================
def make_wobbler():
    fig, ax = plt.subplots(figsize=(5, 5), constrained_layout=True)
    ax.set_xlim(0, 5)
    ax.set_ylim(0, 5)
    ax.set_aspect('equal')
    ax.axis('off')

    # Крепление сверху (маленький прямоугольник)
    ax.add_patch(Rectangle((2.2, 4.3), 0.6, 0.3, facecolor=GRAY, edgecolor=DARK, linewidth=1))

    # Пружинка/ножка
    x_spring = np.linspace(2.5, 2.5, 20)
    y_spring = np.linspace(4.3, 3.8, 20)
    x_wave = 2.5 + 0.1 * np.sin(np.linspace(0, 4*np.pi, 20))
    ax.plot(x_wave, y_spring, color=DARK, linewidth=2)

    # Основная плашка воблера (жёлтая, наклонная)
    wobbler = FancyBboxPatch((0.8, 1.0), 3.4, 2.8,
                              boxstyle="round,pad=0.1",
                              facecolor=YELLOW, edgecolor=DARK, linewidth=2,
                              transform=ax.transData)
    # Поворот
    t = plt.matplotlib.transforms.Affine2D().rotate_deg_around(2.5, 2.4, -8) + ax.transData
    wobbler.set_transform(t)
    ax.add_patch(wobbler)

    # Текст на воблере (тоже повёрнут)
    text_t = plt.matplotlib.transforms.Affine2D().rotate_deg_around(2.5, 2.4, -8) + ax.transData
    ax.text(2.5, 3.0, 'TITAN', ha='center', va='center',
            fontsize=22, color=RED, fontweight='bold', transform=text_t)
    ax.text(2.5, 2.3, 'Красишь 1 раз', ha='center', va='center',
            fontsize=14, color=DARK, fontweight='bold', transform=text_t)
    ax.text(2.5, 1.7, 'служит 10 лет', ha='center', va='center',
            fontsize=14, color=DARK, fontweight='bold', transform=text_t)

    save_fig(fig, '02_wobbler.png')


# ==============================================================
# 3. БРЕНДИРОВАННЫЙ ЦЕННИК
# ==============================================================
def make_price_tag():
    fig, ax = plt.subplots(figsize=(5, 4), constrained_layout=True)
    ax.set_xlim(0, 5)
    ax.set_ylim(0, 4)
    ax.set_aspect('equal')
    ax.axis('off')

    # Ценник (красный, со скошенным левым краем)
    tag = Polygon([(0.8, 0.5), (4.7, 0.5), (4.7, 3.5), (0.8, 3.5), (0.3, 2.0)],
                  closed=True, facecolor=RED, edgecolor=DARK_RED, linewidth=2)
    ax.add_patch(tag)

    # Дырочка для крепления
    ax.add_patch(Circle((0.6, 2.0), 0.08, facecolor=WHITE, edgecolor=DARK, linewidth=1))

    # "Цена за год"
    ax.text(2.7, 3.0, 'ЦЕНА ЗА ГОД', ha='center', va='center',
            fontsize=11, color=YELLOW, fontweight='bold')

    # Большая цифра
    ax.text(2.7, 2.2, '76 руб.', ha='center', va='center',
            fontsize=28, color=WHITE, fontweight='bold')

    # Сравнение
    ax.text(2.7, 1.4, 'вместо 760 руб. за банку', ha='center', va='center',
            fontsize=10, color=WHITE)
    ax.text(2.7, 0.9, '(760 руб. / 10 лет)', ha='center', va='center',
            fontsize=8, color=WHITE, style='italic')

    save_fig(fig, '03_price_tag.png')


# ==============================================================
# 4. ПОРТФОЛИО ОБЪЕКТОВ
# ==============================================================
def make_portfolio():
    fig, ax = plt.subplots(figsize=(8, 5), constrained_layout=True)
    ax.set_xlim(0, 8)
    ax.set_ylim(0, 5)
    ax.set_aspect('equal')
    ax.axis('off')

    # Заголовок
    ax.text(4, 4.6, 'ПОРТФОЛИО ОБЪЕКТОВ', ha='center', va='center',
            fontsize=14, color=RED, fontweight='bold')
    ax.text(4, 4.2, 'с расчётом совокупной стоимости владения', ha='center', va='center',
            fontsize=9, color=GRAY, style='italic')

    # 4 карточки объектов
    objects = [
        ('Ворота', '16 м²', '5 320 руб', '10 лет', RED),
        ('Забор', '40 м²', '13 800 руб', '10 лет', ORANGE),
        ('Гараж', '24 м²', '8 100 руб', '10 лет', GREEN),
        ('Решётки', '8 м²', '2 800 руб', '10 лет', BLUE),
    ]
    for i, (name, area, price, term, color) in enumerate(objects):
        x = 0.3 + i * 1.95
        # Карточка
        card = FancyBboxPatch((x, 0.5), 1.7, 3.2, boxstyle="round,pad=0.05",
                              facecolor=WHITE, edgecolor=color, linewidth=2)
        ax.add_patch(card)
        # Шапка
        ax.add_patch(Rectangle((x, 3.2), 1.7, 0.5, facecolor=color, edgecolor='none'))
        ax.text(x + 0.85, 3.45, name, ha='center', va='center',
                fontsize=11, color=WHITE, fontweight='bold')
        # Площадь
        ax.text(x + 0.85, 2.8, area, ha='center', va='center',
                fontsize=14, color=DARK, fontweight='bold')
        # Цена
        ax.text(x + 0.85, 2.1, price, ha='center', va='center',
                fontsize=12, color=color, fontweight='bold')
        # Срок
        ax.text(x + 0.85, 1.4, term, ha='center', va='center',
                fontsize=10, color=GREEN, fontweight='bold')
        ax.text(x + 0.85, 0.9, 'ССВ за 10 лет', ha='center', va='center',
                fontsize=7, color=GRAY, style='italic')

    save_fig(fig, '04_portfolio.png')


# ==============================================================
# 5. MML + ПЛАНОГРАММА
# ==============================================================
def make_planogram():
    fig, ax = plt.subplots(figsize=(8, 5), constrained_layout=True)
    ax.set_xlim(0, 8)
    ax.set_ylim(0, 5)
    ax.set_aspect('equal')
    ax.axis('off')

    # Заголовок
    ax.text(4, 4.7, 'MML + ПЛАНОГРАММА ВЫКЛАДКИ', ha='center', va='center',
            fontsize=12, color=RED, fontweight='bold')

    # 3 полки
    ral_colors = ['#1A1A1A', '#F5F5DC', '#7B3F00', '#2D5016', '#B0B0B0', '#404040', '#3D2817', '#E6B800', '#0047AB']
    for shelf_idx in range(3):
        y = 3.5 - shelf_idx * 1.3
        # Полка
        ax.add_patch(Rectangle((0.5, y - 0.05), 7, 0.1, facecolor='#8B6F47', edgecolor='#6B5635'))
        # Банки
        for i in range(7):
            x = 0.7 + i * 1.0
            color_idx = (i + shelf_idx * 2) % len(ral_colors)
            color = ral_colors[color_idx]
            # Корпус банки
            ax.add_patch(FancyBboxPatch((x, y + 0.05), 0.7, 1.0, boxstyle="round,pad=0.03",
                                        facecolor=color, edgecolor=DARK, linewidth=0.5))
            # Крышка
            ax.add_patch(Rectangle((x + 0.05, y + 1.05), 0.6, 0.08, facecolor=GRAY, edgecolor=DARK, linewidth=0.5))

    # Подпись
    ax.text(4, 0.3, '3 полки · 22 SKU · 3 линейки TITAN', ha='center', va='center',
            fontsize=9, color=GRAY, style='italic')

    save_fig(fig, '05_planogram.png')


# ==============================================================
# 6. ИНСТРУКЦИЯ ПО НАНЕСЕНИЮ (QR)
# ==============================================================
def make_instruction_qr():
    fig, ax = plt.subplots(figsize=(6, 5), constrained_layout=True)
    ax.set_xlim(0, 6)
    ax.set_ylim(0, 5)
    ax.set_aspect('equal')
    ax.axis('off')

    # Заголовок
    ax.text(3, 4.6, 'ИНСТРУКЦИЯ ПО НАНЕСЕНИЮ', ha='center', va='center',
            fontsize=12, color=RED, fontweight='bold')
    ax.text(3, 4.2, 'открытие по QR-коду', ha='center', va='center',
            fontsize=10, color=GRAY, style='italic')

    # 2 карточки: мини-инструкция + видео
    # Левая — мини-инструкция
    card1 = FancyBboxPatch((0.3, 0.5), 2.6, 3.2, boxstyle="round,pad=0.1",
                           facecolor=WHITE, edgecolor=RED, linewidth=2)
    ax.add_patch(card1)
    ax.add_patch(Rectangle((0.3, 3.2), 2.6, 0.5, facecolor=RED, edgecolor='none'))
    ax.text(1.6, 3.45, 'МИНИ', ha='center', va='center',
            fontsize=11, color=WHITE, fontweight='bold')
    ax.text(1.6, 3.0, 'ИНСТРУКЦИЯ', ha='center', va='center',
            fontsize=10, color=RED, fontweight='bold')
    # 5 шагов (иконки)
    for i in range(5):
        y = 2.5 - i * 0.45
        ax.add_patch(Circle((0.7, y), 0.15, facecolor=RED, edgecolor='none'))
        ax.text(0.7, y, str(i+1), ha='center', va='center',
                fontsize=8, color=WHITE, fontweight='bold')
        steps = ['Очистить', 'Обезжирить', 'Перемешать', 'Покрасить', 'Готово']
        ax.text(1.0, y, steps[i], ha='left', va='center',
                fontsize=9, color=DARK)

    # Правая — видео
    card2 = FancyBboxPatch((3.1, 0.5), 2.6, 3.2, boxstyle="round,pad=0.1",
                           facecolor=WHITE, edgecolor=GREEN, linewidth=2)
    ax.add_patch(card2)
    ax.add_patch(Rectangle((3.1, 3.2), 2.6, 0.5, facecolor=GREEN, edgecolor='none'))
    ax.text(4.4, 3.45, 'ВИДЕО', ha='center', va='center',
            fontsize=11, color=WHITE, fontweight='bold')
    ax.text(4.4, 3.0, 'ИНСТРУКЦИЯ', ha='center', va='center',
            fontsize=10, color=GREEN, fontweight='bold')
    # Иконка play
    ax.add_patch(Circle((4.4, 1.8), 0.5, facecolor=GREEN, edgecolor='none'))
    ax.add_patch(Polygon([(4.3, 1.5), (4.3, 2.1), (4.7, 1.8)],
                         closed=True, facecolor=WHITE, edgecolor='none'))
    ax.text(4.4, 0.9, '5 шагов · 2 мин', ha='center', va='center',
            fontsize=9, color=GRAY)

    save_fig(fig, '06_instruction_qr.png')


# ==============================================================
# 7. ФОТО/ВИДЕО POSm НА ОБЛАКЕ
# ==============================================================
def make_cloud_posm():
    fig, ax = plt.subplots(figsize=(7, 5), constrained_layout=True)
    ax.set_xlim(0, 7)
    ax.set_ylim(0, 5)
    ax.set_aspect('equal')
    ax.axis('off')

    # Заголовок
    ax.text(3.5, 4.6, 'ФОТО И ВИДЕО POSm', ha='center', va='center',
            fontsize=12, color=RED, fontweight='bold')
    ax.text(3.5, 4.2, 'на облаке — для розничных точек', ha='center', va='center',
            fontsize=9, color=GRAY, style='italic')

    # Иконка облака
    cloud_x = [1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5, 5.0, 5.5]
    cloud_y = [2.8, 3.2, 3.3, 3.4, 3.3, 3.4, 3.3, 3.2, 2.8]
    ax.fill_between(cloud_x, cloud_y, 2.5, color=BLUE, alpha=0.3)
    ax.text(3.5, 2.9, 'ОБЛАКО', ha='center', va='center',
            fontsize=11, color=BLUE, fontweight='bold')

    # 4 папки/файла
    files = [
        ('Фото POSm', RED, 0.5, 0.5),
        ('Видео выкладки', ORANGE, 2.2, 0.5),
        ('Шаблоны постов', GREEN, 3.9, 0.5),
        ('Презентации', BLUE, 5.6, 0.5),
    ]
    for name, color, x, y in files:
        # Иконка папки
        ax.add_patch(FancyBboxPatch((x, y), 1.2, 1.4, boxstyle="round,pad=0.05",
                                    facecolor=color, edgecolor=DARK, linewidth=1))
        ax.add_patch(Rectangle((x, y + 1.2), 0.4, 0.2, facecolor=color, edgecolor=DARK, linewidth=1))
        ax.text(x + 0.6, y + 0.7, name, ha='center', va='center',
                fontsize=8, color=WHITE, fontweight='bold')

    save_fig(fig, '07_cloud_posm.png')


# ==============================================================
# 8. КОНТЕНТ ДЛЯ СОЦСЕТЕЙ
# ==============================================================
def make_social_content():
    fig, ax = plt.subplots(figsize=(7, 5), constrained_layout=True)
    ax.set_xlim(0, 7)
    ax.set_ylim(0, 5)
    ax.set_aspect('equal')
    ax.axis('off')

    # Заголовок
    ax.text(3.5, 4.6, 'КОНТЕНТ ДЛЯ СОЦСЕТЕЙ', ha='center', va='center',
            fontsize=12, color=RED, fontweight='bold')
    ax.text(3.5, 4.2, 'готовые посты для дистрибьюторов', ha='center', va='center',
            fontsize=9, color=GRAY, style='italic')

    # 3 карточки постов
    posts = [
        ('Пост 1', 'TITAN 3в1\nКрасишь 1 раз —\nслужит 10 лет', RED, 0.3),
        ('Пост 2', 'ТСО расчёт\nЭкономия 61%\nза 10 лет', GREEN, 2.5),
        ('Пост 3', 'Видео\n5 шагов\nнанесения', ORANGE, 4.7),
    ]
    for name, content, color, x in posts:
        # Карточка поста
        ax.add_patch(FancyBboxPatch((x, 0.5), 2.0, 3.2, boxstyle="round,pad=0.1",
                                    facecolor=WHITE, edgecolor=color, linewidth=2))
        # Шапка
        ax.add_patch(Rectangle((x, 3.2), 2.0, 0.5, facecolor=color, edgecolor='none'))
        ax.text(x + 1.0, 3.45, name, ha='center', va='center',
                fontsize=10, color=WHITE, fontweight='bold')
        # Изображение-заглушка
        ax.add_patch(Rectangle((x + 0.2, 2.0), 1.6, 1.0, facecolor=color, edgecolor='none', alpha=0.3))
        ax.text(x + 1.0, 2.5, '[ фото ]', ha='center', va='center',
                fontsize=8, color=color, fontweight='bold')
        # Текст поста
        ax.text(x + 1.0, 1.2, content, ha='center', va='center',
                fontsize=9, color=DARK)

    save_fig(fig, '08_social_content.png')


# ==============================================================
# 9. ВИДЕОРОЛИКИ
# ==============================================================
def make_videos():
    fig, ax = plt.subplots(figsize=(7, 5), constrained_layout=True)
    ax.set_xlim(0, 7)
    ax.set_ylim(0, 5)
    ax.set_aspect('equal')
    ax.axis('off')

    # Заголовок
    ax.text(3.5, 4.6, 'ВИДЕОРОЛИКИ', ha='center', va='center',
            fontsize=12, color=RED, fontweight='bold')
    ax.text(3.5, 4.2, 'для торговых представителей и продаж', ha='center', va='center',
            fontsize=9, color=GRAY, style='italic')

    # 2 карточки видео
    videos = [
        ('Видео 1', 'Для ТП общего прайса', 'УТП · свойства · выгода\nпо TITAN', RED, 0.5),
        ('Видео 2', 'Идеальная продажа', 'Имитация процесса\nпродажи клиенту\nс потребностью по TITAN', GREEN, 3.7),
    ]
    for name, sub, desc, color, x in videos:
        # Карточка
        ax.add_patch(FancyBboxPatch((x, 0.5), 2.8, 3.2, boxstyle="round,pad=0.1",
                                    facecolor=WHITE, edgecolor=color, linewidth=2))
        # Шапка
        ax.add_patch(Rectangle((x, 3.2), 2.8, 0.5, facecolor=color, edgecolor='none'))
        ax.text(x + 1.4, 3.45, name, ha='center', va='center',
                fontsize=11, color=WHITE, fontweight='bold')
        # Экран видео
        ax.add_patch(Rectangle((x + 0.3, 1.8), 2.2, 1.2, facecolor=DARK, edgecolor=color, linewidth=1))
        # Кнопка play
        ax.add_patch(Circle((x + 1.4, 2.4), 0.3, facecolor=color, edgecolor='none'))
        ax.add_patch(Polygon([(x + 1.3, 2.2), (x + 1.3, 2.6), (x + 1.55, 2.4)],
                             closed=True, facecolor=WHITE, edgecolor='none'))
        # Подзаголовок
        ax.text(x + 1.4, 1.4, sub, ha='center', va='center',
                fontsize=10, color=color, fontweight='bold')
        # Описание
        ax.text(x + 1.4, 0.9, desc, ha='center', va='center',
                fontsize=8, color=DARK)

    save_fig(fig, '09_videos.png')


# ==============================================================
# 10. ИКОНОГРАФИКА ССВ НА БАНКЕ
# ==============================================================
def make_tco_iconography():
    fig, ax = plt.subplots(figsize=(6, 5), constrained_layout=True)
    ax.set_xlim(0, 6)
    ax.set_ylim(0, 5)
    ax.set_aspect('equal')
    ax.axis('off')

    # Заголовок
    ax.text(3, 4.7, 'ИКОНОГРАФИКА ССВ', ha='center', va='center',
            fontsize=12, color=RED, fontweight='bold')
    ax.text(3, 4.3, 'на банке', ha='center', va='center',
            fontsize=9, color=GRAY, style='italic')

    # Банка
    ax.add_patch(FancyBboxPatch((1.5, 0.5), 3, 3.5, boxstyle="round,pad=0.1",
                                facecolor=RED, edgecolor=DARK_RED, linewidth=2))
    # Крышка
    ax.add_patch(Rectangle((1.6, 3.9), 2.8, 0.2, facecolor=GRAY, edgecolor=DARK, linewidth=1))

    # Текст на банке
    ax.text(3, 3.5, 'TITAN', ha='center', va='center',
            fontsize=14, color=WHITE, fontweight='bold')
    ax.text(3, 3.1, '3в1', ha='center', va='center',
            fontsize=10, color=WHITE, fontweight='bold')

    # Иконография ССВ (3 строки)
    ax.text(3, 2.5, 'Дороже банка', ha='center', va='center',
            fontsize=9, color=YELLOW, fontweight='bold')
    ax.text(3, 2.1, 'дешевле покраска!', ha='center', va='center',
            fontsize=9, color=YELLOW, fontweight='bold')
    ax.text(3, 1.5, 'Экономия 60%', ha='center', va='center',
            fontsize=12, color=YELLOW, fontweight='bold')
    ax.text(3, 1.1, 'за 10 лет!', ha='center', va='center',
            fontsize=10, color=YELLOW, fontweight='bold')

    save_fig(fig, '10_tco_iconography.png')


# ==============================================================
# ГЕНЕРАЦИЯ ВСЕХ ВИЗУАЛИЗАЦИЙ
# ==============================================================
print('Генерация визуализаций для третичных продаж...')
make_shelftalker()
make_wobbler()
make_price_tag()
make_portfolio()
make_planogram()
make_instruction_qr()
make_cloud_posm()
make_social_content()
make_videos()
make_tco_iconography()
print(f'---')
print(f'Готово! Файлы в: {OUT_DIR}')
