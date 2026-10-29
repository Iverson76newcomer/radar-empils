"""Скачивание фотографий объектов с сайтов."""
import os
import urllib.request
import json

PHOTOS_DIR = '/home/z/my-project/download/titan_objects_photos'
os.makedirs(PHOTOS_DIR, exist_ok=True)

# Финальный список объектов с найденными фото-URL
objects_final = [
    {
        'id': 1,
        'name': 'Ретрит-центр SABI CLUB Валдай',
        'url': 'https://sabivalday.com/placement',
        'location': 'Валдай, Новгородская область, Центральная Россия',
        'description': (
            'Ретрит-центр «SABI Club» — закрытая территория в живописном лесу на берегу озера '
            'Валдай. На территории — комфортабельные дома с детально продуманным дизайнерским '
            'интерьером и аутентичные юрты. Площадка для проведения ретритов, семинаров и '
            'корпоративных мероприятий.'
        ),
        'materials': 'Грунт-эмаль TITAN (металлоконструкции)',
        'images': [
            'https://static.tildacdn.com/tild6234-6434-4162-b532-353166346630/photo.jpg',
            'https://static.tildacdn.com/tild3331-3434-4939-a138-303832333164/____.jpg',
            'https://static.tildacdn.com/tild6234-3339-4636-b861-356337626665/__1.jpg',
            'https://static.tildacdn.com/tild6661-3435-4562-a464-636461323261/__.jpg',
            'https://static.tildacdn.com/tild6165-3463-4234-a237-346136336361/3_2.jpg',
            'https://static.tildacdn.com/tild3530-3635-4134-b231-613665343030/IMG_6182.jpg',
        ],
    },
    {
        'id': 2,
        'name': 'Эко-парк Soloma',
        'url': 'https://soloma27.ru/',
        'location': 'Хабаровский край, Дальний Восток',
        'description': (
            'Эко-парк Soloma — безмятежный отдых в полном единении с природой. Комплекс '
            'деревянных эко-домов и инфраструктуры в экологическом стиле. Использованы '
            'комбинированные материалы: металлические конструкции с покрытием TITAN, '
            'деревянные рейки AVANTGARDE, яхтный лак PROFILUX.'
        ),
        'materials': 'TITAN (металл) + AVANTGARDE (деревянные рейки) + PROFILUX (яхтный лак)',
        'images': [
            'https://static.tildacdn.com/tild3637-3461-4437-a136-333139653339/3.png',
            'https://static.tildacdn.com/tild6632-6338-4265-b861-616136363663/photo.png',
            'https://static.tildacdn.com/tild6430-3634-4466-b632-656565633364/__.jpg',
            'https://static.tildacdn.com/tild6334-3861-4334-b338-343432633362/photo.jpg',
            'https://static.tildacdn.com/tild3935-6362-4535-b939-666138653264/photo.jpg',
            'https://static.tildacdn.com/tild6362-3564-4166-b235-376538633462/photo.jpg',
            'https://static.tildacdn.com/tild6438-6632-4265-a336-303666363363/photo.PNG',
            'https://static.tildacdn.com/tild3736-3863-4337-a439-646234306631/2.png',
        ],
    },
    {
        'id': 3,
        'name': 'Баня-юрта BraMari',
        'url': 'https://banimed.ru/bramari/',
        'location': 'Челябинск',
        'description': (
            'Баня-юрта BraMari — уникальный проект круглой бани-юрты. Каркас и металлические '
            'конструкции обработаны грунт-эмалью TITAN по ржавчине (3 в 1). Внешняя обшивка и '
            'декоративные элементы выполнены с использованием краски AVANTGARDE. На странице '
            'объекта есть фото каркаса до обшивки — показан этап антикоррозийной защиты.'
        ),
        'materials': 'TITAN (каркас, металл) + AVANTGARDE (декор)',
        'images': [
            'https://banimed.ru/wp-content/uploads/2026/03/kpb8pxgrzbgq8jhqitzjp-1.jpg',
            'https://banimed.ru/wp-content/uploads/2026/03/ffcr8klsurtrj484clrhq-1.jpg',
            'https://banimed.ru/wp-content/uploads/2026/04/img_0842.jpeg',
            'https://banimed.ru/wp-content/uploads/2026/04/2026-02-28-10.00.59.jpg',
            'https://banimed.ru/wp-content/uploads/2026/04/dscf5685.jpg',
            'https://banimed.ru/wp-content/uploads/2026/03/u_2.jpg',
            'https://banimed.ru/wp-content/uploads/2026/03/u_3.jpg',
            'https://banimed.ru/wp-content/uploads/2026/03/u_1.jpg',
        ],
    },
    {
        'id': 4,
        'name': 'База отдыха «Рыбацкие просторы» — Юрт Хаус',
        'url': 'https://fishprostor.ru/yurt-house.html',
        'location': 'Астраханская область',
        'description': (
            'Юрт Хаус на базе отдыха «Рыбацкие просторы» — современный комфорт в атмосфере '
            'этнического стиля. Пение птиц, свежий речной воздух и уют современной юрты на '
            'берегу Волги. Каркас юрты — металлический, обработан грунт-эмалью TITAN 3 в 1 по '
            'ржавчине. Есть фото каркаса на странице объекта.'
        ),
        'materials': 'Грунт-эмаль TITAN (металлический каркас юрты)',
        'images': [
            'https://fishprostor.ru/img/yurt/3d.jpg',
            'https://fishprostor.ru/img/yurt/header-yort.jpg',
            'https://fishprostor.ru/img/yurt/3.jpg',
            'https://fishprostor.ru/img/yurt/2.jpg',
            'https://fishprostor.ru/img/yurt/9.jpg',
            'https://fishprostor.ru/img/yurt/7.jpg',
            'https://fishprostor.ru/img/yurt/12.jpg',
            'https://fishprostor.ru/img/yurt/10.jpg',
            'https://fishprostor.ru/img/yurt/11.jpg',
        ],
    },
    {
        'id': 5,
        'name': 'Агротуристический комплекс «Голубая лагуна»',
        'url': 'https://vk.ru/lagunabuida',
        'location': 'с. Буйда, Республика Башкортостан',
        'description': (
            'Агротуристический комплекс «Голубая лагуна» — место семейного отдыха на природе. '
            'Лес, поля, пруд. На территории — бани, беседки, места для рыбалки и активного '
            'отдыха. Металлические конструкции (беседки, навесы, ограждения) обработаны '
            'грунт-эмалью TITAN 3 в 1 по ржавчине для долговечной защиты от атмосферных '
            'воздействий.'
        ),
        'materials': 'Грунт-эмаль TITAN (металлоконструкции, навесы, ограждения)',
        'images': [
            'https://sun1-25.vkuserphoto.ru/ZdV8d-XHc_5GsnwSSBvJSCZ7Dv5TUEdZuQILrw/dTxtHF9uF1M.png',
            'https://sun1-25.vkuserphoto.ru/6DD0xSghpEY4ZOnSoOQGo5hvOqZld01WrEb-dA/I-cI-hyWOi0.png',
        ],
    },
]

# Скачиваем фото
print("Скачивание фотографий объектов...")
print("=" * 80)

for obj in objects_final:
    obj_dir = os.path.join(PHOTOS_DIR, f"obj{obj['id']}")
    os.makedirs(obj_dir, exist_ok=True)
    print(f"\nОбъект {obj['id']}: {obj['name']}")
    print(f"  Локальных фото: {len(obj['images'])}")
    downloaded = []
    for i, img_url in enumerate(obj['images'], 1):
        # Определяем расширение
        ext = '.jpg'
        for e in ['.png', '.jpeg', '.jpg', '.webp', '.PNG']:
            if e in img_url:
                ext = e.lower()
                if ext == '.jpeg':
                    ext = '.jpg'
                break
        out_path = os.path.join(obj_dir, f"photo_{i:02d}{ext}")
        try:
            req = urllib.request.Request(img_url, headers={
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
            })
            with urllib.request.urlopen(req, timeout=20) as resp:
                if resp.status == 200:
                    data = resp.read()
                    if len(data) > 3000:  # минимальный размер фото
                        with open(out_path, 'wb') as f:
                            f.write(data)
                        downloaded.append(out_path)
                        print(f"  ✓ {i}. {os.path.basename(out_path)} ({len(data)//1024} KB)")
                    else:
                        print(f"  ✗ {i}. Файл слишком маленький ({len(data)} байт)")
        except Exception as e:
            print(f"  ✗ {i}. Ошибка: {str(e)[:60]}")
    obj['local_images'] = downloaded

# Сохраняем финальные данные
with open('/home/z/my-project/scripts/objects_final.json', 'w', encoding='utf-8') as f:
    json.dump(objects_final, f, ensure_ascii=False, indent=2, default=str)

print(f"\n{'='*80}")
print(f"ИТОГО:")
total = sum(len(obj.get('local_images', [])) for obj in objects_final)
print(f"  Скачано {total} фото для {len(objects_final)} объектов")
for obj in objects_final:
    print(f"  {obj['id']}. {obj['name']}: {len(obj.get('local_images', []))} фото")
