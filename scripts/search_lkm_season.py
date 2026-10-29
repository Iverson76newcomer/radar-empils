"""Поиск данных о сезонности ЛКМ-рынка через открытые источники."""
import urllib.request
import urllib.parse
import re

# Try direct fetch from industry sources
urls_to_try = [
    "https://www.lkmport.ru/articles/sezonnost-prodazh-lkm",
    "https://krasko.ru/articles/sezonnost-lkm",
    "https://www.forhouse.ru/articles/sezonnost-pokraski",
]

# Try DuckDuckGo HTML (often accessible)
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'ru-RU,ru;q=0.9,en;q=0.8',
}

# Search query for LKM seasonality in Russia
query = urllib.parse.quote("сезонность продаж лакокрасочных материалов Россия по месяцам статистика")
search_url = f"https://html.duckduckgo.com/html/?q={query}"

try:
    req = urllib.request.Request(search_url, headers=headers)
    with urllib.request.urlopen(req, timeout=15) as resp:
        html = resp.read().decode('utf-8', errors='ignore')
    
    # Save full HTML for analysis
    with open('cash_gap_research/ddg_search.html', 'w', encoding='utf-8') as f:
        f.write(html)
    
    # Extract result snippets
    results = re.findall(r'<a[^>]+class="result__a"[^>]*href="([^"]+)"[^>]*>(.*?)</a>', html, re.DOTALL)
    snippets = re.findall(r'<a[^>]+class="result__snippet"[^>]*>(.*?)</a>', html, re.DOTALL)
    
    print(f"=== DuckDuckGo results: {len(results)} ===")
    for i, (url, title) in enumerate(results[:10], 1):
        title = re.sub(r'<[^>]+>', '', title)
        print(f"\n{i}. {title[:100]}")
        print(f"   URL: {url[:120]}")
        if i-1 < len(snippets):
            snippet = re.sub(r'<[^>]+>', '', snippets[i-1])
            print(f"   Snippet: {snippet[:300]}")
except Exception as e:
    print(f"Error: {e}")
