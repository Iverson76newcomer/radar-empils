"""Получить детальные данные с найденных страниц."""
import urllib.request
import re

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml',
    'Accept-Language': 'ru-RU,ru;q=0.9,en;q=0.8',
}

urls_to_fetch = [
    # Континенталь Воронеж
    ("https://www.rusprofile.ru/id/4473683", "Континенталь Воронеж (ИНН 3663083053) - rusprofile"),
    ("https://www.audit-it.ru/contragent/1093668036571_ooo-kontinental", "Континенталь Воронеж (ИНН 3663078399) - audit-it"),
    # Аплайн Вологда
    ("https://apline35.ru/company/contacts/", "Аплайн Вологда - контакты"),
    ("https://apline35.ru/", "Аплайн Вологда - главная"),
    ("https://www.audit-it.ru/contragent/1093525017772_ooo-aplayn", "АПЛАЙН Вологда - audit-it"),
]

for url, name in urls_to_fetch:
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
        for kw in ['ИНН', 'ОГРН', 'директор', 'Директор', 'Генеральный', 'адрес', 'Адрес', 
                   'телефон', 'Телефон', 'сайт', 'Сайт', 'выручка', 'Выручка', 
                   'численность', 'Численность', 'учредитель', 'Учредитель',
                   'ОКВЭД', 'основной вид', 'статус', 'дата регистр']:
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
    except Exception as e:
        print(f"Error: {e}")
