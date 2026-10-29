"""Открыть ключевые сайты и собрать детальную информацию."""
import urllib.request
import re
import json

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml',
    'Accept-Language': 'ru-RU,ru;q=0.9,en;q=0.8',
}

# Ключевые сайты для детального изучения
urls_to_fetch = [
    # Сайты производителей колеровочного оборудования
    ("https://tintec.ru/", "Tintec — производитель колеровочного оборудования"),
    ("https://omnitint.ru/", "Omnitint — универсальная колеровочная система"),
    ("https://omnitint.ru/paint-shops", "Omnitint — страница для магазинов ЛКМ"),
    ("https://kolerovka.ru/", "Kolerovka.ru — каталог колеровочного оборудования"),
    ("https://kolerovka.com/", "Kolerovka.com — Corob, Fast&Fluid, Hero"),
    ("https://tecsa.ru/product-category/equipment/kolerovochnoe-oborudovanie/", "Tecsa — колеровочное оборудование"),
    ("https://lkm-kolerovka.ru/catalog/lkm/kolerovka", "LKM-kolerovka — виды оборудования"),
    ("https://www.santint-rus.ru/", "Santint — китайское оборудование"),
    ("https://vpscolor.ru/", "VPS Color — колеровочное оборудование"),
    ("https://palizh.com/catalog-industry/kolerovochnoe-oborudovanie/", "Palizh — колеровочное оборудование"),
    ("https://palizh.info/tinting/", "Palizh.info — преимущества колеровочного оборудования"),
    ("https://www.plasters.ru/dealer/kolerovochnoe-oborudovanie/", "Plasters — для дилеров"),
    ("https://www.gravihel-helios.com/ru/kolierovochnoie-oborudovaniie-i-koloristichieskiie-instrumienty/", "Gravihel — оборудование и инструменты"),
    ("http://ekokolor.ru/oborudovanie.php", "Эко Колор — оборудование"),
    # Производители ЛКМ про колеровку
    ("https://tikkurila.kz/dlya-professionalov/vdohnovenie/tikkurila-symphony-instrukciya-po-primeneniyu", "Tikkurila Symphony — инструкция"),
    ("https://www.dulux.ru/", "Dulux Россия"),
    ("https://dulux-color.ru/", "Dulux Color"),
    ("https://www.complexconstruction.ru/upload/iblock/793/Tikkurila_Catalog.pdf", "Tikkurila PDF каталог"),
    ("https://wagner.ru/professionalnoe/brands/tikkurila.html", "Wagner — Tikkurila"),
    ("https://gelingen.biz/promyshlennoe-oborudovanie/tikkurila/", "Gelingen — Tikkurila"),
    # Зарубежные
    ("https://www.corobgroup.com/", "Corob Group — официальный сайт"),
    ("https://www.fast-fluid.com/", "Fast & Fluid Management"),
    ("https://www.hero-srl.com/", "Hero SRL"),
    ("https://www.cpscolor.com/", "CPS Color"),
    ("https://www.dromont.com/", "Dromont"),
    ("https://www.chromaflo.com/", "Chromaflo"),
]

results = []
for url, name in urls_to_fetch:
    print(f"\n=== {name}: {url} ===")
    try:
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=15) as resp:
            html = resp.read().decode('utf-8', errors='ignore')
        
        # Извлечь текст
        text = re.sub(r'<script[^>]*>.*?</script>', '', html, flags=re.DOTALL)
        text = re.sub(r'<style[^>]*>.*?</style>', '', text, flags=re.DOTALL)
        text = re.sub(r'<[^>]+>', ' ', text)
        text = re.sub(r'\s+', ' ', text).strip()
        
        # Заголовок
        title_m = re.search(r'<title[^>]*>([^<]+)</title>', html)
        title = title_m.group(1).strip() if title_m else name
        
        # Поиск ключевых слов и контекста
        keywords = ['колеровочное оборудование', 'колеровочная зона', 'расстановка', 'магазин', 'дилер', 
                    'мерчендайзинг', 'планограмма', 'колеровочный центр', 'колеровочная машина', 
                    'дизайн места', 'колеровочный стенд', 'колеровочная станция',
                    'tint zone', 'color center', 'merchandising', 'point of sale', 'POS',
                    'retail', 'store', 'shop', 'equipment', 'dispenser', 'shaker', 'mixer']
        
        snippets = []
        for kw in keywords:
            idx = 0
            count = 0
            while count < 2:
                idx = text.lower().find(kw.lower(), idx)
                if idx < 0:
                    break
                start = max(0, idx - 100)
                end = min(len(text), idx + 350)
                context = text[start:end].strip()
                if len(context) > 50:
                    snippets.append({'keyword': kw, 'context': context})
                    count += 1
                idx += len(kw)
        
        results.append({
            'url': url,
            'name': name,
            'title': title,
            'length': len(text),
            'snippets': snippets[:6],  # топ-6 самых релевантных
        })
        
        print(f"  Title: {title}")
        print(f"  Length: {len(text)} chars")
        for s in snippets[:3]:
            print(f"  [{s['keyword']}]: {s['context'][:300]}")
    except Exception as e:
        print(f"  Error: {e}")
        results.append({'url': url, 'name': name, 'error': str(e)})

with open('tinting_research/detailed_content.json', 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)
print(f"\n=== Изучено {len(results)} сайтов ===")
