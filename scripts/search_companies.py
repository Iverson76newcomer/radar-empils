"""Поиск данных по 4 компаниям через интернет."""
import urllib.request
import urllib.parse
import re
import json

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml',
    'Accept-Language': 'ru-RU,ru;q=0.9,en;q=0.8',
}

companies = [
    {'name': 'ООО "СТРОЙБАЗА - СПЕКТР"', 'city': 'Липецк', 'region': 'Липецкая обл.', 'short': 'Стройбаза Спектр Липецк'},
    {'name': 'ООО Стройоптторг', 'city': 'Вологда', 'region': 'Вологодская обл.', 'short': 'Стройоптторг Вологда'},
    {'name': 'ООО "КОНТИНЕНТАЛЬ"', 'city': 'Воронеж', 'region': 'Воронежская обл.', 'short': 'Континенталь Воронеж ЛКМ'},
    {'name': 'ООО "Аплайн"', 'city': 'Вологда', 'region': 'Вологодская обл.', 'short': 'Аплайн Вологда ЛКМ краски'},
]

all_results = {}

for company in companies:
    print(f"\n{'=' * 100}")
    print(f"ПОИСК: {company['name']} ({company['city']})")
    print(f"{'=' * 100}")
    
    queries = [
        f"{company['short']} ИНН контакты",
        f"{company['short']} сайт телефон адрес",
        f"{company['short']} выручка директор",
    ]
    
    company_results = []
    for q in queries:
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
                company_results.append({
                    'title': title_clean[:150],
                    'url': url[:300],
                    'snippet': snippet[:400],
                })
                print(f"\n  {i}. {title_clean[:100]}")
                print(f"     URL: {url[:120]}")
                if snippet:
                    print(f"     {snippet[:300]}")
        except Exception as e:
            print(f"  Error: {e}")
    
    all_results[company['name']] = company_results

with open('/home/z/my-project/scripts/search_results_companies.json', 'w', encoding='utf-8') as f:
    json.dump(all_results, f, ensure_ascii=False, indent=2)
print(f"\n\n=== Сохранено {sum(len(v) for v in all_results.values())} результатов ===")
