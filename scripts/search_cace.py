"""Поиск информации о Central Asia Coating Exhibition в Астане."""
import urllib.request
import urllib.parse
import re

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml',
    'Accept-Language': 'ru-RU,ru;q=0.9,en;q=0.8',
}

queries = [
    "Central Asia Coating Exhibition Astana Kazakhstan 2026",
    "Central Asia Coating Exhibition Астана Казахстан выставка",
    "CACE Astana coating exhibition dates venue",
    "Central Asia Coating Exhibition 2026 участники",
    "Central Asia Coating Show Astana",
]

all_results = {}
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
        
        items = []
        for i, (url, title) in enumerate(match_results[:5], 1):
            title_clean = re.sub(r'<[^>]+>', '', title).replace('&amp;', '&')
            if 'uddg=' in url:
                m = re.search(r'uddg=([^&]+)', url)
                if m:
                    url = urllib.parse.unquote(m.group(1))
            snippet = ''
            if i-1 < len(snippets):
                snippet = re.sub(r'<[^>]+>', '', snippets[i-1]).replace('&amp;', '&')
            items.append({'title': title_clean, 'url': url, 'snippet': snippet})
            print(f"  {i}. {title_clean[:100]}")
            print(f"     URL: {url[:120]}")
            if snippet:
                print(f"     {snippet[:300]}")
        all_results[q] = items
    except Exception as e:
        print(f"  Error: {e}")

import json
with open('/home/z/my-project/scripts/cace_search.json', 'w', encoding='utf-8') as f:
    json.dump(all_results, f, ensure_ascii=False, indent=2)
print(f"\n=== Сохранено {sum(len(v) for v in all_results.values())} результатов ===")
