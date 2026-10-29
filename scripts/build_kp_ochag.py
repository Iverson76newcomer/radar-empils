"""
Создать КП для Очаг на основе структуры КП Олимп с обновлёнными данными рынка.
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement
import copy

# === Открыть оригинальный КП Олимп ===
doc = Document("/home/z/my-project/upload/КП_Олимп_коммерческое_предложение 12.10.2026.docx")

# === Данные для замены ===
# Олимп: 5 регионов, 3,475,000 городского населения, потенциал 3.3 млрд, факт 1 млрд, цель 312 млн
# Очаг: Красноярский край + Хакасия, 1,854,000 городского населения, потенциал 2.1 млрд, факт 858 млн, цель 170 млн

# === 1. Заменить текстовые параграфы ===
replacements = {
    # Партнёр
    'Партнёр: Олимп · Новосибирская, Томская, Кемеровская обл., Алтайский край, Республика Алтай':
    'Партнёр: Очаг · Красноярский край, Республика Хакасия',
    
    # Цель
    'Цель: 312 млн ₽ товарооборота в Год 3 = 8% потенциала региона. Заработок партнёра: 109 млн ₽/год при наценке 35%.':
    'Цель: 170 млн ₽ товарооборота в Год 3 = 8% потенциала региона. Заработок партнёра: 60 млн ₽/год при наценке 35%.',
    
    # Население
    'База расчёта: городское население — 3 475 000 чел.':
    'База расчёта: городское население — 1 854 000 чел.',
    
    # Недозагрузка
    '⚠ Регион недозагружен на 69% — потенциал 2.3 млрд ₽ не освоен':
    '⚠ Регион недозагружен на 59% — потенциал 1.2 млрд ₽ не освоен',
    
    # Доля Эмпилс
    '312 млн ₽ — целевой товарооборот Эмпилс, достигаемый за 3 года. Покрытие АКБ: 1200 РТТ.':
    '170 млн ₽ — целевой товарооборот Эмпилс, достигаемый за 3 года. Покрытие АКБ: 700 РТТ.',
    
    # Отсрочка — без изменений
    
    # 75% ЭТП
    '75% ЭТП нанимается в период первичной загрузки дистрибьютора.':
    '75% ЭТП нанимается в период первичной загрузки дистрибьютора.',
    
    # Стартовый ассортимент — без изменений
    
    # Предложение действительно
    'Предложение действительно до 15.12.2026':
    'Предложение действительно до 15.12.2026',
}

for para in doc.paragraphs:
    for old, new in replacements.items():
        if old in para.text:
            # Заменить текст
            for run in para.runs:
                if old in run.text:
                    run.text = run.text.replace(old, new)
            # Если не найдено в отдельных runs, заменить в первом
            if old in para.text and new not in para.text:
                for run in para.runs:
                    run.text = ''
                para.runs[0].text = new

# === 2. Заменить таблицы ===

# Таблица 1: ПОТЕНЦИАЛ РЫНКА / ФАКТИЧЕСКИЙ ТО / НЕДОЗАГРУЗКА / ДОЛЯ ЭМПИЛС
t = doc.tables[0]
t.rows[1].cells[0].text = '2.1 млрд ₽\n27 000 т'  # Потенциал
t.rows[1].cells[1].text = '0.86 млрд ₽\n858 млн ₽'  # Факт
t.rows[1].cells[2].text = '1.24 млрд ₽\n59% рынка'  # Недозагрузка
t.rows[1].cells[3].text = '170 млн ₽\nцель на 3 года'  # Доля Эмпилс

# Таблица 2: Товарооборот по годам
t = doc.tables[1]
t.rows[1].cells[1].text = '35 млн ₽'  # Год 1
t.rows[1].cells[2].text = '12 млн ₽'
t.rows[1].cells[3].text = '2 ЭТП + ТП дистрибьютора, 300 SKU'
t.rows[2].cells[1].text = '70 млн ₽'  # Год 2
t.rows[2].cells[2].text = '24 млн ₽'
t.rows[2].cells[3].text = 'Расширение АКБ до РТТ'
t.rows[3].cells[1].text = '170 млн ₽'  # Год 3
t.rows[3].cells[2].text = '60 млн ₽'
t.rows[3].cells[3].text = 'Полное покрытие 8% потенциала'
t.rows[4].cells[1].text = '275 млн ₽'  # Итого
t.rows[4].cells[2].text = '96 млн ₽'

# Таблица 3: Бонусирование
t = doc.tables[2]
# Год 1: 35 млн
t.rows[1].cells[1].text = '35 млн ₽'
t.rows[1].cells[2].text = '1,4 млн ₽'  # 4% ассорт
t.rows[1].cells[3].text = '1,4 млн ₽'  # 4% ежемес
t.rows[1].cells[4].text = '0,35 млн ₽ (1%)'
t.rows[1].cells[5].text = '3,15 млн ₽ (9%)'
t.rows[1].cells[6].text = '15,2 млн ₽'
# Год 2: 70 млн
t.rows[2].cells[1].text = '70 млн ₽'
t.rows[2].cells[2].text = '2,8 млн ₽'
t.rows[2].cells[3].text = '2,8 млн ₽'
t.rows[2].cells[4].text = '1,4 млн ₽ (2%)'
t.rows[2].cells[5].text = '7,0 млн ₽ (10%)'
t.rows[2].cells[6].text = '31,5 млн ₽'
# Год 3: 170 млн
t.rows[3].cells[1].text = '170 млн ₽'
t.rows[3].cells[2].text = '6,8 млн ₽'
t.rows[3].cells[3].text = '6,8 млн ₽'
t.rows[3].cells[4].text = '5,1 млн ₽ (3%)'
t.rows[3].cells[5].text = '18,7 млн ₽ (11%)'
t.rows[3].cells[6].text = '78,2 млн ₽'
# Итого
t.rows[4].cells[1].text = '275 млн ₽'
t.rows[4].cells[2].text = '11,0 млн ₽'
t.rows[4].cells[3].text = '11,0 млн ₽'
t.rows[4].cells[4].text = '6,85 млн ₽'
t.rows[4].cells[5].text = '28,85 млн ₽'
t.rows[4].cells[6].text = '124,9 млн ₽'

# Таблица 4: Население по регионам
t = doc.tables[3]
# Очистить и заполнить заново
# Row 1: заголовок — оставить
# Красноярский край
t.rows[1].cells[0].text = 'Красноярский край'
t.rows[1].cells[1].text = '2 866 000'
t.rows[1].cells[2].text = 'Красноярск'
t.rows[1].cells[3].text = '1 095 000'
# Хакасия (используем row 2 — было Кемерово)
t.rows[2].cells[0].text = 'Республика Хакасия'
t.rows[2].cells[1].text = '530 000'
t.rows[2].cells[2].text = 'Абакан'
t.rows[2].cells[3].text = '187 000'
# Дополнительные города Красноярского края (row 3 — было Алтайский)
t.rows[3].cells[0].text = ' + Норильск, Ачинск, Канск, Железногорск'
t.rows[3].cells[1].text = '—'
t.rows[3].cells[2].text = '4 города'
t.rows[3].cells[3].text = '453 000'
# Дополнительные города Хакасии (row 4 — было Томская)
t.rows[4].cells[0].text = ' + Черногорск, Саяногорск'
t.rows[4].cells[1].text = '—'
t.rows[4].cells[2].text = '2 города'
t.rows[4].cells[3].text = '119 000'
# ИТОГО (row 5 — было Республика Алтай)
t.rows[5].cells[0].text = 'ИТОГО'
t.rows[5].cells[1].text = '3 396 000'
t.rows[5].cells[2].text = '8 городов'
t.rows[5].cells[3].text = '1 854 000'
# Row 6 — было ИТОГО, сделать пустой
for c in range(4):
    t.rows[6].cells[c].text = ''

# Таблица 5: Расчёт потенциала
t = doc.tables[4]
# Row 3: расчёт
t.rows[2].cells[1].text = '1 854 000 × 1123 ₽'
t.rows[2].cells[2].text = '1 854 000 × 14.59 кг'
# Row 4: потенциал
t.rows[3].cells[1].text = '2.1 млрд ₽'
t.rows[3].cells[2].text = '27 000 т'

# Таблица 6: Дистрибьюторы (заменить полностью)
t = doc.tables[5]
# Очистить существующие строки данных (rows 2-8)
# Row 1: заголовок — оставить
# Красноярский край
t.rows[1].cells[0].text = 'СУЧКОВ А.Н. (ИП)'
t.rows[1].cells[1].text = 'АВС Фарбен, Декарт, Новбытхим'
t.rows[1].cells[2].text = 'Красноярский край'
t.rows[2].cells[0].text = 'ООО «АКСИОМА»'
t.rows[2].cells[1].text = 'Лакра, Ярославские краски'
t.rows[2].cells[2].text = 'Красноярский край'
t.rows[3].cells[0].text = 'ЛИШУТИН А.В. (ИП)'
t.rows[3].cells[1].text = 'Рогнеда, Эмпилс, Кушнаренко, ТД Радуга'
t.rows[3].cells[2].text = 'Красноярский край'
t.rows[4].cells[0].text = 'ООО «КРАСХИМРЕСУРС»'
t.rows[4].cells[1].text = 'ВИТ, Эмпилс, Новбытхим'
t.rows[4].cells[2].text = 'Красноярский край'
t.rows[5].cells[0].text = 'ООО «АЛКОТРЭЙД 66»'
t.rows[5].cells[1].text = 'Оксоль'
t.rows[5].cells[2].text = 'Красноярский край'
t.rows[6].cells[0].text = 'Есиневич (ИП)'
t.rows[6].cells[1].text = 'ВИТ, Кушнаренко, Сайвер'
t.rows[6].cells[2].text = 'Республика Хакасия'
t.rows[7].cells[0].text = 'ИТОГО (6 дистрибьюторов)'
t.rows[7].cells[1].text = 'Совокупный ТО: 858 млн ₽'
t.rows[7].cells[2].text = '2 региона'

# Таблица 7: Недозагрузка
t = doc.tables[6]
t.rows[1].cells[1].text = '2.1'  # Теоретический потенциал
t.rows[1].cells[2].text = '27 000'
t.rows[1].cells[3].text = '100%'
t.rows[2].cells[0].text = 'Фактический ТО дистрибьюторов'
t.rows[2].cells[1].text = '0.86'
t.rows[2].cells[2].text = '858 млн ₽'
t.rows[2].cells[3].text = '41%'
t.rows[3].cells[0].text = 'НЕДОЗАГРУЗКА'
t.rows[3].cells[1].text = '1.24'
t.rows[3].cells[2].text = '15 700 т'
t.rows[3].cells[3].text = '59%'

# Таблица 8: Доля Эмпилс
t = doc.tables[7]
t.rows[1].cells[1].text = '2.1 млрд ₽'
t.rows[1].cells[2].text = '27 000 т'
t.rows[2].cells[0].text = 'Доля Эмпилс (8%)'
t.rows[2].cells[1].text = '170 млн ₽'
t.rows[2].cells[2].text = '2 160 т'

# Таблица 10: ЭТП
t = doc.tables[9]
t.rows[1].cells[1].text = '120 РТТ ОКБ (80 АКБ)'  # без изменений
t.rows[2].cells[0].text = '2 ЭТП покрывают'
t.rows[2].cells[1].text = '240 РТТ ОКБ (160 АКБ)'
t.rows[3].cells[0].text = 'Товарооборот через ЭТП'
t.rows[3].cells[1].text = '23 млн ₽/год'
t.rows[4].cells[0].text = 'Товарооборот через ТП дистрибьютора'
t.rows[4].cells[1].text = '12 млн ₽/год'
t.rows[5].cells[0].text = 'ИТОГО год 1'
t.rows[5].cells[1].text = '35 млн ₽'

# Таблица 11: Прогноз по годам
t = doc.tables[10]
t.rows[1].cells[1].text = '35 млн ₽'
t.rows[1].cells[3].text = '2 ЭТП + ТП дистрибьютора, 300 SKU'
t.rows[2].cells[1].text = '70 млн ₽'
t.rows[2].cells[2].text = '×2'
t.rows[2].cells[3].text = 'Расширение АКБ до 500+ РТТ + 1 ЭТП'
t.rows[3].cells[1].text = '170 млн ₽'
t.rows[3].cells[2].text = '×4,9'
t.rows[3].cells[3].text = 'Полное покрытие 8% потенциала + ЭТП'

# === Сохранить ===
out_path = "/home/z/my-project/download/КП_Очаг_коммерческое_предложение.docx"
doc.save(out_path)

import os
print(f"✓ Файл сохранён: {out_path}")
print(f"  Размер: {os.path.getsize(out_path):,} байт")

# === Проверка ===
doc2 = Document(out_path)
print(f"\n=== Проверка ===")
print(f"Абзацев: {len(doc2.paragraphs)}, таблиц: {len(doc2.tables)}")
print("\n--- Партнёр ---")
print(doc2.paragraphs[1].text)
print("\n--- Таблица 1 (потенциал) ---")
for row in doc2.tables[0].rows:
    print([c.text[:40] for c in row.cells])
print("\n--- Таблица 2 (ТО по годам) ---")
for row in doc2.tables[1].rows:
    print([c.text[:40] for c in row.cells])
print("\n--- Таблица 4 (население) ---")
for row in doc2.tables[3].rows:
    print([c.text[:40] for c in row.cells])
print("\n--- Таблица 6 (дистрибьюторы) ---")
for row in doc2.tables[5].rows:
    print([c.text[:50] for c in row.cells])
