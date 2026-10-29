"""Получить детальные данные с сайтов компаний и справочников."""
import urllib.request
import re

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml',
    'Accept-Language': 'ru-RU,ru;q=0.9,en;q=0.8',
}

urls_to_fetch = [
    # Стройбаза-Спектр Липецк
    ("https://spektr48.ru/", "Стройбаза-Спектр Липецк — сайт"),
    ("https://www.find-org.net/cli/6456831_ooo_strojjbaza_spektr", "Стройбаза-Спектр — find-org"),
    
    # Стройоптторг Вологда
    ("https://tdsot.ru/", "Стройоптторг Вологда — сайт"),
    ("https://tdsot.ru/customer/rekvizityi/", "Стройоптторг — реквизиты"),
    
    # Континенталь Воронеж
    ("https://krafor.com/", "Континенталь Воронеж — krafor.com"),
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
        for kw in ['ИНН', 'ОГРН', 'телефон', 'Телефон', 'адрес', 'Адрес', 
                   'директор', 'Директор', 'Генеральный', 'email', 'Email', 'E-mail',
                   'контакт', 'Контакт', 'режим работы', 'Часы работы']:
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
