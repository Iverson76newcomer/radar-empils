"""Получить детальную информацию с сайтов выставки."""
import urllib.request
import re

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml',
    'Accept-Language': 'ru-RU,ru;q=0.9,en;q=0.8',
}

urls = [
    ("https://centralasiacoatingshow.com/en/", "Официальный сайт"),
    ("https://centralasiacoatingshow.com/en/for-visitors/location-and-time-of-the-exhibition", "Место и время"),
    ("https://centralasiacoatingshow.com/en/for-visitors/list-of-exhibitors-2026", "Список участников 2026"),
    ("https://metallicheckiy-portal.ru/events/2026_central_asia_coating", "Металлический портал"),
    ("https://expolista.com/tradeshows/central-asia-coatings-show/2026", "Expolista"),
]

for url, name in urls:
    print(f"\n{'='*100}")
    print(f"{name}: {url}")
    print(f"{'='*100}")
    try:
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=20) as resp:
            html = resp.read().decode('utf-8', errors='ignore')
        
        text = re.sub(r'<script[^>]*>.*?</script>', '', html, flags=re.DOTALL)
        text = re.sub(r'<style[^>]*>.*?</style>', '', text, flags=re.DOTALL)
        text = re.sub(r'<[^>]+>', '\n', text)
        text = re.sub(r'\n\s*\n', '\n', text)
        text = re.sub(r'[ \t]+', ' ', text).strip()
        
        print(f"Длина: {len(text)} символов")
        
        # Ключевые слова
        for kw in ['дата', 'Дата', 'расписание', 'время', 'место', 'Место', 'адрес', 
                   'участник', 'Участник', 'экспонент', 'страна', 'павильон',
                   'январ', 'Январ', '2026', '2027', 'EXPO', 'Астана', 'Astana',
                   'лакокрасоч', 'краск', 'покрыт', 'сырьё', 'оборудование',
                   'registration', 'визит', 'вход', 'ticket', 'билет']:
            idx = 0
            count = 0
            while count < 1:
                idx = text.lower().find(kw.lower(), idx)
                if idx < 0:
                    break
                start = max(0, idx - 80)
                end = min(len(text), idx + 300)
                context = text[start:end].strip()
                if len(context) > 30:
                    print(f"  [{kw}]: ...{context[:350]}...")
                    count += 1
                idx += len(kw)
    except Exception as e:
        print(f"Error: {e}")
