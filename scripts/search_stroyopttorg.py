"""Поиск Стройоптторг Вологда — больше деталей."""
import urllib.request
import urllib.parse
import re

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml',
    'Accept-Language': 'ru-RU,ru;q=0.9,en;q=0.8',
}

queries = [
    "tdsot.ru реквизиты ИНН директор",
    "Стройоптторг Вологда ИНН 3525392787 директор",
    "ООО Стройоптторг Вологда Преображенского 18 телефон",
    "СтройОптТорг Вологда ЛКМ бренды VGT Tikkurila",
    "ООО Стройоптторг 3525298463 Вологодская обл",
]

for q in queries:
    print(f"\n=== {q} ===")
    encoded = urllib.parse.quote(q)
    search_url = f"https://html.duckduckgo.com/html/?q={encoded}"
    try:
        req = urllib.request.Request(search_url, headers=headers)
        with urllib.request.urlopen(req, timeout=15) as resp:
            html = resp.read().decode('utf-8', errors='ignore')
        
        match_results = re.findall(r'class="result__a"[^>]*href="([^"]+)"[^>]*>(.*?)</a>', html, re.DOTALL)
        snippets = re.findall(r'class="result__snippet"[^>]*>(.*?)</a>', html, re.DOTALL)
        
        for i, (url, title) in enumerate(match_results[:4], 1):
            title_clean = re.sub(r'<[^>]+>', '', title).replace('&amp;', '&')
            if 'uddg=' in url:
                m = re.search(r'uddg=([^&]+)', url)
                if m:
                    url = urllib.parse.unquote(m.group(1))
            snippet = ''
            if i-1 < len(snippets):
                snippet = re.sub(r'<[^>]+>', '', snippets[i-1]).replace('&amp;', '&')
            print(f"  {i}. {title_clean[:100]}")
            print(f"     URL: {url[:120]}")
            if snippet:
                print(f"     {snippet[:300]}")
    except Exception as e:
        print(f"  Error: {e}")
