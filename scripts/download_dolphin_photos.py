"""Скачивание фото тюрьмы Черный дельфин."""
import os
import urllib.request

PHOTOS_DIR = '/home/z/my-project/download/titan_objects_photos/obj6'
os.makedirs(PHOTOS_DIR, exist_ok=True)

# Отобранные красивые фото
photos = [
    # Главное фото здания с лебедем/дельфином (lada.rest)
    ('https://lada.tagster.club/wp-content/uploads/sites/76/2018/11/black-dolphin.jpg', 'photo_01.jpg'),
    # Фото с entermedia.io — основной корпус
    ('https://entermedia.io/wp-content/uploads/2021/04/pr-min.jpg', 'photo_02.jpg'),
    # Фото с lapsha.media — корпус/озеро
    ('https://lapsha.media/storage/uploads/2024/05/ozero.jpg', 'photo_03.jpg'),
    ('https://lapsha.media/storage/uploads/2024/05/fsin.jpg', 'photo_04.jpg'),
    # Фото с lemur59.ru — корпус 21
    ('http://lemur59.ru/sites/default/files/images/%D0%BA%D0%BE%D1%80%D0%BF%D1%83%D1%81%2021.jpg', 'photo_05.jpg'),
    ('http://lemur59.ru/sites/default/files/images/%D0%BA%D0%BE%D1%80%D0%BF%D1%83%D1%81%D0%B0%20image4.jpg', 'photo_06.jpg'),
    ('http://lemur59.ru/sites/default/files/images/%D0%B8%D0%B7%D0%BE%D0%BB%D1%8F%D1%82%D0%BE%D1%80%20src_1b4bb198a0e902dd-01.jpg', 'photo_07.jpg'),
    ('http://lemur59.ru/sites/default/files/images/%D0%BE%D0%B7%D0%B5%D1%80%D0%BE%20001-01.jpg', 'photo_08.jpg'),
]

downloaded = []
for url, filename in photos:
    out_path = os.path.join(PHOTOS_DIR, filename)
    try:
        req = urllib.request.Request(url, headers={
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })
        with urllib.request.urlopen(req, timeout=20) as resp:
            if resp.status == 200:
                data = resp.read()
                if len(data) > 3000:
                    with open(out_path, 'wb') as f:
                        f.write(data)
                    downloaded.append(out_path)
                    print(f"✓ {filename} ({len(data)//1024} KB)")
                else:
                    print(f"✗ {filename} — слишком маленький ({len(data)} байт)")
    except Exception as e:
        print(f"✗ {filename} — {str(e)[:60]}")

print(f"\nСкачано {len(downloaded)} фото")
