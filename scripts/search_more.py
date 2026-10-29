"""Дополнительный поиск — конкретно про расстановку и оформление зоны колеровки."""
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
    "планограмма колеровочной зоны ЛКМ расстановка оборудования",
    "как оформить зону колеровки в строительном магазине",
    "tint zone merchandising paint store layout",
    "color center paint store display merchandising",
    "point of sale tinting display design",
    "колеровочная машина в магазине дизайн места",
    "короб колеровочное оборудование официальный сайт",
    "fast fluid management колеровочное оборудование",
    "chromex колеровочное оборудование",
    "Speerit колеровочное оборудование для ЛКМ",
    "ColorMatic колеровочное оборудование Россия",
    "Hero LS колеровочное оборудование",
    "Protekt колеровочное оборудование Россия",
    "tecsa колеровочное оборудование",
    "Тиккурила колеровочное оборудование для дилеров",
    "колеровочное оборудование для строительного магазина выбор",
    "retail paint store color mixing station design",
    "paint store color corner merchandising",
    "in-store tinting equipment layout guide",
    "Dulux Color Center retail store equipment",
    "AkzoNobel Mixit color station dealer",
    "Benjamin Moore color station retail layout",
    "Sherwin Williams ProMar color center",
    "колеровочное оборудование где ставить в магазине",
    "колеровочная зона в магазине металла краски",
    "Bandit Fast Fluid equipment retailer brochure",
    "how to set up tinting zone in paint store",
    "wholesale tint equipment brochure PDF",
    "колеровочное оборудование буклет PDF",
    "tinting zone guide PDF brochure",
    "колеровочное оборудование текса типы",
    "Vandex колеровочное оборудование",
    "Rokolor колеровочное оборудование",
    "Коллеруем колеровочное оборудование производители",
    "колеровочное оборудование инструкция по установке",
    "Izmерительная колеровка оборудование розница",
    "диспенсер колоранта для магазина расстановка",
]

results = {}
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
        for i, (url, title) in enumerate(match_results[:6], 1):
            title_clean = re.sub(r'<[^>]+>', '', title).replace('&amp;', '&')
            if 'uddg=' in url:
                m = re.search(r'uddg=([^&]+)', url)
                if m:
                    url = urllib.parse.unquote(m.group(1))
            snippet = ''
            if i-1 < len(snippets):
                snippet = re.sub(r'<[^>]+>', '', snippets[i-1]).replace('&amp;', '&')
            items.append({'title': title_clean[:150], 'url': url[:300], 'snippet': snippet[:400]})
            print(f"  {i}. {title_clean[:90]}")
            print(f"     URL: {url[:120]}")
            if snippet:
                print(f"     {snippet[:250]}")
        results[q] = items
    except Exception as e:
        print(f"  Error: {e}")
        results[q] = []

with open('tinting_research/search_results_2.json', 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)
print(f"\n=== Сохранено {sum(len(v) for v in results.values())} результатов ===")
