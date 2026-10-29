"""Финальный отчёт по сравнению."""
import openpyxl
from openpyxl import load_workbook

# === Загрузить общую таблицу ===
wb = load_workbook("/home/z/my-project/upload/Belye_pyatna_obschaya_tablitsa_21_09_2026.xlsx", data_only=True)
ws = wb["Белые пятна"]

general_partners = []
for r in range(2, ws.max_row + 1):
    name = ws.cell(row=r, column=4).value
    if name and str(name).strip():
        general_partners.append({
            'row': r,
            'name': str(name).strip(),
            'inn': ws.cell(row=r, column=5).value,
            'city': ws.cell(row=r, column=3).value,
        })

print("=" * 90)
print("РЕЗУЛЬТАТ СРАВНЕНИЯ: файл «Белые пятна - Москва» vs общая CRM-таблица")
print("=" * 90)

print(f"""
Общая таблица CRM: {len(general_partners)} партнёров
Файл «Белые пятна - Москва»: 6 уникальных дистрибьюторов (по 2 регионам)

ВНИМАНИЕ: файл «Белые пятна - Москва» содержит НЕ ТОЛЬКО московских партнёров,
но и компании из Воронежской, Липецкой и Вологодской областей.
Это фактически файл по Центральному федеральному округу (ЦФО).
""")

# === Результаты ===
results = [
    {
        'name_msk': 'ООО "ВЕРТИКАЛЬ"',
        'territory': 'Московская обл. + Воронежская обл.',
        'to': '1664 (Мск) + 554 (Врн) млн руб',
        'found': True,
        'name_general': 'Вертикаль (строка 115)',
        'city_general': 'Воронеж',
        'inn_general': '5027161766',
        'note': '✓ Есть в общей таблице. В файле Москвы указана дважды (по двум регионам), в общей таблице — одной строкой.',
    },
    {
        'name_msk': 'ООО "СТРОЙТЕРМИНАЛ ЦЕНТР КРАСОК"',
        'territory': 'Московская обл.',
        'to': '920 млн руб',
        'found': True,
        'name_general': 'СтройТерминал-Центр Красок (строка 159)',
        'city_general': 'Москва',
        'inn_general': 'не указан',
        'note': '✓ Есть в общей таблице. Имя немного отличается (дефис вместо пробела).',
    },
    {
        'name_msk': 'ООО "СТРОЙБАЗА - СПЕКТР"',
        'territory': 'Липецкая обл.',
        'to': '264 млн руб',
        'found': False,
        'name_general': '—',
        'city_general': '—',
        'inn_general': '—',
        'note': '✗ ОТСУТСТВУЕТ в общей таблице. Частичные совпадения (Стройбаза №1 в Челябинске, СтройБаза-14 в Якутии) — это другие компании.',
    },
    {
        'name_msk': 'ООО Стройоптторг',
        'territory': 'Вологодская обл.',
        'to': '110 млн руб',
        'found': False,
        'name_general': '—',
        'city_general': '—',
        'inn_general': '—',
        'note': '✗ ОТСУТСТВУЕТ в общей таблице.',
    },
    {
        'name_msk': 'ООО "КОНТИНЕНТАЛЬ"',
        'territory': 'Воронежская обл.',
        'to': '107 млн руб',
        'found': False,
        'name_general': '—',
        'city_general': '—',
        'inn_general': '—',
        'note': '✗ ОТСУТСТВУЕТ в общей таблице.',
    },
    {
        'name_msk': 'ООО "Аплайн"',
        'territory': 'Вологодская обл.',
        'to': '55 млн руб',
        'found': False,
        'name_general': '—',
        'city_general': '—',
        'inn_general': '—',
        'note': '✗ ОТСУТСТВУЕТ в общей таблице.',
    },
]

found_count = sum(1 for r in results if r['found'])
not_found_count = sum(1 for r in results if not r['found'])

print(f"{'№':<4} {'Дистрибьютор':<35} {'Регион':<22} {'ТО, млн':<12} {'В общей таблице?':<20}")
print("-" * 95)
for i, r in enumerate(results, 1):
    status = '✓ ДА' if r['found'] else '✗ НЕТ'
    print(f"{i:<4} {r['name_msk']:<35} {r['territory']:<22} {r['to']:<12} {status:<20}")

print(f"\n{'=' * 90}")
print(f"ИТОГ: {found_count} из {len(results)} найдены в общей таблице, {not_found_count} отсутствуют")
print(f"{'=' * 90}")

print(f"\n{'=' * 90}")
print("ДЕТАЛИ ПО НЕНАЙДЕННЫМ:")
print(f"{'=' * 90}")
for r in results:
    if not r['found']:
        print(f"\n✗ {r['name_msk']}")
        print(f"   Регион: {r['territory']}")
        print(f"   ТО: {r['to']} млн руб")
        print(f"   {r['note']}")

print(f"\n{'=' * 90}")
print("ДЕТАЛИ ПО НАЙДЕННЫМ:")
print(f"{'=' * 90}")
for r in results:
    if r['found']:
        print(f"\n✓ {r['name_msk']}")
        print(f"   В файле Москвы: регион={r['territory']}, ТО={r['to']}")
        print(f"   В общей таблице: {r['name_general']}, город={r['city_general']}, ИНН={r['inn_general']}")
        print(f"   {r['note']}")
