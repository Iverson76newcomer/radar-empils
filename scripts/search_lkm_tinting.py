"""Поиск информации о колеровочном оборудовании на сайтах производителей ЛКМ."""
import urllib.request
import urllib.parse
import re
import json

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml',
    'Accept-Language': 'ru-RU,ru;q=0.9,en;q=0.8',
}

queries = [
    "Tikkurila колеровочное оборудование для магазинов инструкция",
    "Dulux колеровочная машина для магазинов расстановка",
    "Beckers колеровочное оборудование retailer",
    "колеровочное оборудование в магазине ЛКМ где ставить",
    "Tikkurila Façade Color Planet оборудование",
    "колеровочная машина в магазине расстановка",
    "Corob колеровочное оборудование для магазинов",
    "Fast Fluid колеровочное оборудование для ЛКМ",
    "ColorMatic колеровочное оборудование розница",
    "Dromont колеровочное оборудование для ЛКМ",
    "Hero колеровочное оборудование ЛКМ",
    "CPS Color колеровочное оборудование для магазинов",
    "Protekt колеровочное оборудование для ЛКМ Россия",
    "колеровка краски в магазине оборудование типы",
    "колеровочная зона в магазине дизайн",
    "оформление зоны колеровки в строительном магазине",
    "планограмма колеровочной зоны ЛКМ",
    "retail tinting equipment guide LKM",
    "AkzoNobel tinting machine retail guide",
    "Sherwin-Williams ColorToGo equipment retail",
    "VTEX колеровочное оборудование",
    "Colpar колеровочное оборудование",
    "Hempel колеровочное оборудование retailer",
]

results = {}
for q in queries:
    print(f"\n=== Query: {q} ===")
    encoded = urllib.parse.quote(q)
    search_url = f"https://html.duckduckgo.com/html/?q={encoded}"
    try:
        req = urllib.request.Request(search_url, headers=headers)
        with urllib.request.urlopen(req, timeout=15) as resp:
            html = resp.read().decode('utf-8', errors='ignore')
        
        # Извлечь результаты
        match_results = re.findall(r'class="result__a"[^>]*href="([^"]+)"[^>]*>(.*?)</a>', html, re.DOTALL)
        snippets = re.findall(r'class="result__snippet"[^>]*>(.*?)</a>', html, re.DOTALL)
        
        items = []
        for i, (url, title) in enumerate(match_results[:6], 1):
            title_clean = re.sub(r'<[^>]+>', '', title).replace('&amp;', '&')
            if 'uddg=' in url:
                m = re.search(r'uddg=([^&]+)', url)
                if m:
                    url = urllib.parse.unquote(m.group(1))
            snippet = ''
            if i-1 < len(snippets):
                snippet = re.sub(r'<[^>]+>', '', snippets[i-1]).replace('&amp;', '&')
            items.append({'title': title_clean[:120], 'url': url[:200], 'snippet': snippet[:300]})
            print(f"  {i}. {title_clean[:90]}")
            print(f"     URL: {url[:100]}")
            if snippet:
                print(f"     {snippet[:200]}")
        results[q] = items
    except Exception as e:
        print(f"  Error: {e}")
        results[q] = []

# Сохранить
with open('tinting_research/search_results.json', 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)
print(f"\n\n=== Сохранено {sum(len(v) for v in results.values())} результатов ===")
