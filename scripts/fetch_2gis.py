"""Получить данные через 2ГИС и другие источники."""
import urllib.request
import re

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml',
    'Accept-Language': 'ru-RU,ru;q=0.9,en;q=0.8',
}

urls = [
    # Континенталь Воронеж — 2ГИС
    ("https://2gis.ru/voronezh/firm/70000001020861669", "Континенталь Воронеж — 2ГИС"),
    # Континенталь Воронеж — Яндекс Карты  
    ("https://yandex.ru/maps/org/kontinental/1275176666/", "Континенталь Воронеж — Яндекс Карты"),
    # Стройоптторг Вологда
    ("https://vologda.jsprav.ru/magazinyi-hoztovarov/strojopttorg3347/", "Стройоптторг Вологда — справочник"),
    # Аплайн Вологда — rubri.co
    ("https://vlg.rubri.co/c/aplayn/", "Аплайн Вологда — rubri.co"),
]

for url, name in urls:
    print(f"\n{'=' * 100}")
    print(f"{name}: {url}")
    print(f"{'=' * 100}")
    try:
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=15) as resp:
            html = resp.read().decode('utf-8', errors='ignore')
        
        text = re.sub(r'<script[^>]*>.*?</script>', '', html, flags=re.DOTALL)
        text = re.sub(r'<style[^>]*>.*?</style>', '', text, flags=re.DOTALL)
        text = re.sub(r'<[^>]+>', ' ', text)
        text = re.sub(r'\s+', ' ', text).strip()
        
        print(f"Длина: {len(text)} символов")
        
        # Извлечь ключевые данные
        for kw in ['телефон', 'Телефон', 'адрес', 'Адрес', 'сайт', 'email', 'Email',
                   'режим', 'часы', 'работает']:
            idx = 0
            count = 0
            while count < 2:
                idx = text.lower().find(kw.lower(), idx)
                if idx < 0:
                    break
                start = max(0, idx - 50)
                end = min(len(text), idx + 250)
                context = text[start:end].strip()
                if len(context) > 30:
                    print(f"  [{kw}]: ...{context}...")
                    count += 1
                idx += len(kw)
        
        # Также вывести первые 2000 символов
        print(f"\n--- Первые 2000 символов ---")
        print(text[:2000])
    except Exception as e:
        print(f"Error: {e}")
