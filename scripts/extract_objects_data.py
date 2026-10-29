"""Извлечение текста и URL изображений из 5 страниц объектов."""
import json
import re
import os

objects = [
    {'id': 1, 'name': 'Ретрит-центр «Валдай»', 'file': '/tmp/obj1.json', 'url': 'https://sabivalday.com/placement', 'location': 'Валдай, Центральная Россия'},
    {'id': 2, 'name': 'Эко-парк Soloma', 'file': '/tmp/obj2.json', 'url': 'https://soloma27.ru/', 'location': 'Хабаровский край'},
    {'id': 3, 'name': 'Баня-юрта BraMari', 'file': '/tmp/obj3.json', 'url': 'https://banimed.ru/bramari/', 'location': 'Челябинск'},
    {'id': 4, 'name': 'База отдыха «Рыбацкие просторы» (Юрт Хаус)', 'file': '/tmp/obj4.json', 'url': 'https://fishprostor.ru/yurt-house.html', 'location': 'Астраханская область'},
    {'id': 5, 'name': 'Агротуристический комплекс «Голубая лагуна»', 'file': '/tmp/obj5.json', 'url': 'https://vk.ru/lagunabuida', 'location': 'Башкортостан, с. Буйда'},
]

for obj in objects:
    print(f"\n{'='*80}")
    print(f"ОБЪЕКТ {obj['id']}: {obj['name']}")
    print(f"URL: {obj['url']}")
    print(f"Локация: {obj['location']}")
    print(f"{'='*80}")

    if not os.path.exists(obj['file']):
        print(f"  ФАЙЛ НЕ НАЙДЕН")
        continue

    with open(obj['file']) as f:
        d = json.load(f)
    data = d.get('data', {})
    html = data.get('html', '')
    title = data.get('title', '')
    desc = data.get('description', '')

    print(f"Title: {title[:200]}")
    print(f"Description: {desc[:300]}")
    print(f"HTML length: {len(html)}")

    # Извлекаем изображения — ищем все <img src="...">
    img_patterns = [
        r'<img[^>]+src=["\']([^"\']+)["\']',
        r'<img[^>]+data-src=["\']([^"\']+)["\']',
        r'<img[^>]+data-original=["\']([^"\']+)["\']',
    ]
    images = []
    for pat in img_patterns:
        images.extend(re.findall(pat, html, re.IGNORECASE))

    # Фильтруем — убираем иконки, лого, трекеры
    filtered = []
    for img in images:
        if not img or img.startswith('data:'):
            continue
        # Пропускаем иконки/логотипы/мелкие
        skip_keywords = ['icon', 'logo', 'sprite', 'pixel', 'tracker', 'facebook.com/tr',
                         'googletag', 'vk.com/images/', 'avatar', 'emoji', '.svg', '1x1',
                         'placeholder', 'no-image', 'blank.gif']
        if any(kw in img.lower() for kw in skip_keywords):
            continue
        # Если относительный URL — добавим базовый домен
        if img.startswith('//'):
            img = 'https:' + img
        elif img.startswith('/'):
            from urllib.parse import urlparse
            base = urlparse(obj['url'])
            img = f"{base.scheme}://{base.netloc}{img}"
        elif not img.startswith('http'):
            continue
        # Уникальные
        if img not in filtered:
            filtered.append(img)

    print(f"\nНайдено изображений: {len(filtered)}")
    for i, img in enumerate(filtered[:10]):
        print(f"  {i+1}. {img[:150]}")

    # Сохраняем список URL
    obj['images'] = filtered[:8]

# Сохраняем итог
with open('/home/z/my-project/scripts/objects_data.json', 'w', encoding='utf-8') as f:
    json.dump(objects, f, ensure_ascii=False, indent=2, default=str)

print(f"\n\n=== Данные сохранены в objects_data.json ===")
for obj in objects:
    print(f"  {obj['id']}. {obj['name']} → {len(obj.get('images', []))} фото")
