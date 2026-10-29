#!/usr/bin/env python3
"""
Анализ участников Kazbuild 2026 с точки зрения потенциала как дистрибьюторов
или партнёров Эмпилс по ЛКМ.
"""

import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

# === Стилевые константы ===
EMPILS_ORANGE = "F47629"
DARK = "313131"
WHITE = "FFFFFF"
LIGHT_ORANGE = "FDE6D5"

# Категории с цветами
CAT_COLORS = {
    "Производитель ЛКМ":        "FFCDD2",  # красный — конкурент
    "Дистрибьютор ЛКМ":         "C8E6C9",  # зелёный — цель
    "Строительная химия":       "FFE0B2",  # оранжевый — смежный
    "Герметики/Пены":           "FFE0B2",  # оранжевый — смежный
    "Сырьё для ЛКМ":            "BBDEFB",  # голубой — поставщик
    "Промышленный потребитель": "D1C4E9",  # фиолетовый — клиент
    "Смежные":                  "FFF9C4",  # жёлтый — слабая связь
    "Не связано с ЛКМ":         "E0E0E0",  # серый — не интересно
}

REC_COLORS = {
    "Высокий":   "C8E6C9",
    "Средний":   "FFE0B2",
    "Низкий":    "FFCDD2",
    "Не целевой": "E0E0E0",
}

# === Данные анализа ===
# (Компания, Страна, Категория, Информация, Сайт, Рекомендация, Обоснование)
COMPANIES = [
    ("ALIHAN COMPANY ALMATY", "Казахстан", "Дистрибьютор ЛКМ",
     "Оптовый продавец строительных материалов в Алматы. Широкий ассортимент, рейтинг 4.6 на Яндекс.Картах.",
     "yandex.com/maps/org/alihan_company/134843741634",
     "Средний", "Оптовик строительных материалов — может быть дистрибьютором Эмпилс по водно-дисперсионным краскам в Алматы"),

    ("ANADOLU CONSTRUCTION GROUP", "Узбекистан", "Не связано с ЛКМ",
     "Строительная компания, не производит и не дистрибьютирует ЛКМ.",
     "—", "Не целевой", "Не специализируется на ЛКМ"),

    ("ANHUI HUANHONG PLASTICS CO., LTD.", "Китай", "Не связано с ЛКМ",
     "Производитель пластиковых изделий.",
     "—", "Не целевой", "Пластиковая продукция, не связано с ЛКМ"),

    ("AOFENG", "Китай", "Герметики/Пены",
     "Pingdingshan Aofeng New Materials — производитель полиуретановых пен и силиконовых герметиков, 20+ лет опыта.",
     "aofengsealant.en.made-in-china.com",
     "Низкий", "Конкурент в нише монтажных пен/герметиков, не дистрибьютор красок"),

    ("BINZHOU LVBANG BOARD CO., LTD", "Китай", "Не связано с ЛКМ",
     "Производитель строительных плит.",
     "—", "Не целевой", "Панели и плиты, не ЛКМ"),

    ("CASALI SPA", "Италия", "Строительная химия",
     "Casali Group (Италия, с 1920-х) — производитель гидроизоляционных материалов и жидких покрытий. Часть Casali ABM.",
     "casali-group.com",
     "Низкий", "Конкурент в нише гидроизоляции. Возможен обмен технологиями, но не дистрибуция ЛКМ Эмпилс"),

    ("DIBA", "Иран", "Производитель ЛКМ",
     "Pars Diba Chemical Co. — бренд «Barad», крупнейший производитель декоративных и индустриальных красок в Иране, 45+ лет опыта.",
     "linkedin.com/company/barad-paint",
     "Низкий", "Прямой конкурент-производитель красок. Не дистрибьютор."),

    ("DUAYEN YAPI KIMYA VE MAKINE SAN.", "Турция", "Строительная химия",
     "Duayen Yapı A.Ş. (с 2001) — производитель промышленных клеёв, полиуретановых покрытий и наливных полов. 600 т/мес.",
     "duayen.com.tr/en",
     "Низкий", "Конкурент в нише промышленных покрытий и наливных полов."),

    ("DUOWEI UNION GROUP CO., LTD.", "Китай", "Не связано с ЛКМ",
     "Производитель строительных конструкций.",
     "—", "Не целевой", "Не связано с ЛКМ"),

    ("EWECA TRADING B.V.", "Нидерланды", "Смежные",
     "Eweca Trading B.V. — поставщик материалов и компонентов для строительства и обслуживания.",
     "galabau-messe.com/en/exhibitors/eweca-trading-bv-2571022",
     "Низкий", "Широкий профиль, не специализируется на ЛКМ"),

    ("FAB GROUP", "Азербайджан", "Производитель ЛКМ",
     "FAB Boya (с 2001) — производитель красок, лаков, эпоксидных смол и строительной химии в Баку. Лидер в Азербайджане.",
     "fabboya.az/en",
     "Низкий", "Прямой конкурент-производитель. Возможен обмен дистрибуцией на рынке Каспия."),

    ("FLOORPARTNER", "Китай", "Не связано с ЛКМ",
     "Производитель напольных покрытий.",
     "—", "Не целевой", "Напольные покрытия, не ЛКМ"),

    ("GUANGDONG KINGWINS NEW MATERIALS", "Китай", "Сырьё для ЛКМ",
     "Производитель новых материалов для строительства.",
     "—", "Низкий", "Возможный поставщик сырья, не дистрибьютор готовых ЛКМ"),

    ("GUANGZHOU DIYA INTERNATIONAL TRADING", "Китай", "Смежные",
     "Международная торговая компания строительными материалами.",
     "—", "Низкий", "Торговая компания широкого профиля"),

    ("GUANGZHOU HEXIN INDUSTRY CO., LTD", "Китай", "Смежные",
     "Промышленная компания строительных материалов.",
     "—", "Не целевой", "Недостаточно данных о связи с ЛКМ"),

    ("HAINING GENSIN PLASTIC SHEET", "Китай", "Не связано с ЛКМ",
     "Производитель пластиковых листов.",
     "—", "Не целевой", "Пластиковые листы"),

    ("HAIYAN HAOSHANGJI NEW MATERIALS", "Китай", "Сырьё для ЛКМ",
     "Производитель новых материалов.",
     "—", "Низкий", "Возможный поставщик сырья"),

    ("HENAN JIXIANG INDUSTRIAL", "Китай", "Не связано с ЛКМ",
     "Производитель промышленной продукции.",
     "—", "Не целевой", "Не связано с ЛКМ"),

    ("JIABOS", "Китай", "Не связано с ЛКМ",
     "Нет данных о связи с ЛКМ.",
     "—", "Не целевой", "Недостаточно информации"),

    ("JIANGSU YARET INTERNATIONAL TRADE", "Китай", "Смежные",
     "Международная торговля строительными материалами.",
     "—", "Низкий", "Широкий профиль, не специализирован на ЛКМ"),

    ("KAYSDAA", None, "Не связано с ЛКМ",
     "Нет данных.",
     "—", "Не целевой", "Недостаточно информации"),

    ("LANGFANG XINDUOLI NEW BUILDING MATERIALS", "Китай", "Не связано с ЛКМ",
     "Производитель строительных материалов.",
     "—", "Не целевой", "Не связано с ЛКМ"),

    ("LCP BUILDING PRODUCTS", "Сингапур", "Смежные",
     "Строительные материалы.",
     "—", "Низкий", "Широкий профиль"),

    ("LINYI KINSPIRE NEW MATERIAL", "Китай", "Сырьё для ЛКМ",
     "Новые материалы для строительства.",
     "—", "Низкий", "Возможный поставщик сырья"),

    ("LINYI OASIS NEW BUILDING MATERIALS", "Китай", "Не связано с ЛКМ",
     "Строительные материалы.",
     "—", "Не целевой", "Не ЛКМ"),

    ("MARMOX EGYPT", "Египет", "Не связано с ЛКМ",
     "Производитель мрамора и изделий из камня.",
     "—", "Не целевой", "Камень и мрамор, не ЛКМ"),

    ("MARRIONI DESIGN", "Казахстан", "Смежные",
     "Дизайн-студия интерьеров в Казахстане.",
     "—",
     "Средний", "Дизайн-студия — потенциальный продвиженец продуктов Эмпилс премиум-сегмента (AvantGarde) через спецификации в проектах"),

    ("MSD КОМПАНИЯ ТЕПЛОИЗОЛЯЦИОННЫХ ПОКРЫТИЙ", "Казахстан", "Смежные",
     "По данным поиска — производитель ПВХ-плёнок (MSD-PVF), а не теплоизоляционных красок. Название на выставке может отличаться от реального профиля.",
     "zjmsd.com",
     "Низкий", "Уточнить профиль на стенде. Если реально теплоизоляционные покрытия — потенциальный партнёр"),

    ("NANJING SZFC NEW MATERIALS", "Китай", "Сырьё для ЛКМ",
     "Производитель новых материалов.",
     "—", "Низкий", "Возможный поставщик сырья"),

    ("PUFAS", "Германия", "Производитель ЛКМ",
     "PUFAS — немецкий бренд (с 1920-х) красок, эмалей, грунтовок и строительной химии. Широко представлен в DIY-ритейле в Европе.",
     "pufas.de",
     "Низкий", "Прямой конкурент в премиум-сегменте. Не дистрибьютор."),

    ("QUANZHOU CITY SANLIAN MACHINERY", "Китай", "Не связано с ЛКМ",
     "Производитель оборудования.",
     "—", "Не целевой", "Машиностроение"),

    ("ROHDE KG", "Германия", "Смежные",
     "Rohde KG — производитель малярного инструмента (валики, кисти, лотки).",
     "rollingdogtools.com",
     "Низкий", "Поставщик малярного инструмента. Возможен кросс-маркетинг с Эмпилс, но не дистрибьютор ЛКМ"),

    ("SEHOME CHINA", "Китай", "Не связано с ЛКМ",
     "Строительные материалы.",
     "—", "Не целевой", "Недостаточно данных"),

    ("SHAANXI LARSEN NEW MATERIAL", "Китай", "Сырьё для ЛКМ",
     "Новые материалы.",
     "—", "Низкий", "Возможный поставщик сырья"),

    ("SHANDONG ALUMEDAL NEW MATERIAL", "Китай", "Не связано с ЛКМ",
     "Алюминиевые материалы.",
     "—", "Не целевой", "Алюминий, не ЛКМ"),

    ("SHANDONG DRUKEC NEW MATERIALS", "Китай", "Сырьё для ЛКМ",
     "Производитель эпоксидных смол и химических модификаторов для ЛКМ.",
     "landu-china.com",
     "Низкий", "Поставщик сырья (эпоксид, эфиры целлюлозы, VAE). Интересен как поставщик, не как дистрибьютор"),

    ("SHANDONG HUAJIAN ALUMINIUM GROUP", "Китай", "Не связано с ЛКМ",
     "Алюминиевый профиль.",
     "—", "Не целевой", "Алюминий, не ЛКМ"),

    ("SHANDONG JINING SANDIA WOOD", "Китай", "Промышленный потребитель",
     "Производитель древесных материалов.",
     "—",
     "Средний", "Покупатель ЛКМ по дереву (PROFIWOOD-аналог). Может потреблять грунты и лаки Эмпилс для своих изделий"),

    ("SHANDONG JUHUAN NEW MATERIAL", "Казахстан", "Герметики/Пены",
     "Shandong Juhuan New Material Technology — производитель PU пены, силиконовых герметиков, спрей-красок, жидких гвоздей. На Kazbuild зарегистрирован как Казахстан.",
     "juhuanpu.en.alibaba.com",
     "Низкий", "Конкурент в нише герметиков и аэрозольных красок. Не дистрибьютор ЛКМ"),

    ("SHANDONG KESHANG INTERNATIONAL TRADE", "Китай", "Сырьё для ЛКМ",
     "Международная торговля химической продукцией (Shandong Kanghong / Kesheng Chemistry).",
     "echemi.com/shop-us20210422173274120",
     "Низкий", "Трейдер химии. Поставщик сырья, не дистрибьютор готовых ЛКМ"),

    ("SHANDONG OYG NEW BUILDING MATERIALS", "Китай", "Не связано с ЛКМ",
     "Строительные материалы.",
     "—", "Не целевой", "Не ЛКМ"),

    ("SHANGHAI ALUBANG DECORATIVE MATERIAL", "Китай", "Не связано с ЛКМ",
     "Декоративные материалы.",
     "—", "Не целевой", "Декоративные панели, не ЛКМ"),

    ("SHENYANG ALISBA DECORATIVE MATERIALS", "Китай", "Не связано с ЛКМ",
     "Декоративные материалы.",
     "—", "Не целевой", "Не ЛКМ"),

    ("SI MOSAIC", "Россия", "Не связано с ЛКМ",
     "Мозаика.",
     "—", "Не целевой", "Мозаика, не ЛКМ"),

    ("SOMAFIX - SOMA GLOBAL", "Турция", "Герметики/Пены",
     "Somafix (Soma Kimya) — производитель монтажных пен, силиконовых герметиков, клеёв и строительной химии. Широкий экспорт.",
     "somafix.com.tr/en",
     "Низкий", "Конкурент в нише монтажной химии. Не дистрибьютор ЛКМ"),

    ("SUZHOU EVER-HAILING NEW MATERIALS", "Китай", "Сырьё для ЛКМ",
     "Новые материалы.",
     "—", "Низкий", "Возможный поставщик сырья"),

    ("TANGSHAN BAIJIA TRADING", "Китай", "Смежные",
     "Торговая компания.",
     "—", "Низкий", "Широкий профиль"),

    ("TAUDI", "Казахстан", "Производитель ЛКМ",
     "TAUDI — крупнейший отечественный производитель лакокрасочных материалов в Казахстане. Современное оборудование, широкая линейка продукции.",
     "taudi.kz",
     "Средний", "Главный конкурент Эмпилс в Казахстане. Не дистрибьютор. Можно изучить их ассортимент для отстройки"),

    ("TIMBERSTONE TOO", "Казахстан", "Промышленный потребитель",
     "Древесные материалы и изделия из дерева.",
     "—",
     "Средний", "Покупатель ЛКМ по дереву (грунты, лаки, антисептики). Потенциальный промышленный клиент Эмпилс"),

    ("UNITED ACADEMIC DESIGN", "Казахстан", "Смежные",
     "ТОО «United Academic Design», БИН 050140025047. Алматы, Алмалинский район. Руководитель: Лю И.",
     "kiberon.kz/search?lang=kz&search=050140025047",
     "Низкий", "По названию — дизайн-студия. Уточнить профиль на стенде"),

    ("WAYBOMALUMINUM", "Китай", "Не связано с ЛКМ",
     "Алюминиевая продукция.",
     "—", "Не целевой", "Алюминий"),

    ("XIANTAO MEI HOME DECORATION MATERIALS", "Китай", "Не связано с ЛКМ",
     "Декоративные материалы для дома.",
     "—", "Не целевой", "Декор, не ЛКМ"),

    ("XINXIANG YIBAIYUAN HIGH-TECH MATERIAL", "Китай", "Сырьё для ЛКМ",
     "Высокотехнологичные материалы.",
     "—", "Низкий", "Возможный поставщик сырья"),

    ("YIWU JIUFENG IMPORT AND EXPORT", "Китай", "Смежные",
     "Импортно-экспортная компания.",
     "—", "Низкий", "Широкий профиль"),

    ("YIWU JOCY IMPORT AND EXPORT", "Китай", "Смежные",
     "Импортно-экспортная компания.",
     "—", "Низкий", "Широкий профиль"),

    ("YONGFENG SELF-ADHESIVE", "Китай", "Смежные",
     "Wenzhou Yongfeng Self-Adhesive Materials — производитель самоклеящихся материалов (стикеры, наклейки на окна/стены).",
     "prettgold.en.alibaba.com",
     "Низкий", "Самоклеящиеся материалы. Возможен кросс-маркетинг, но не дистрибьютор ЛКМ"),

    ("ZHEJIANG ECOMAX DECORATIVE MATERIAL", "Китай", "Не связано с ЛКМ",
     "Декоративные материалы.",
     "—", "Не целевой", "Декор"),

    ("ZHEJIANG LOBEL WALLPAPER", "Китай", "Не связано с ЛКМ",
     "Обои.",
     "—", "Не целевой", "Обои, не ЛКМ"),

    ("ZHEJIANG QIDE NEW MATERIALS", "Китай", "Сырьё для ЛКМ",
     "Новые материалы.",
     "—", "Низкий", "Возможный поставщик сырья"),

    ("АЛЮМИНИЕВАЯ КОМПОЗИТНАЯ ПАНЕЛЬ ALUSIGNPANEL", "Китай", "Не связано с ЛКМ",
     "Алюминиевые композитные панели.",
     "—", "Не целевой", "Композитные панели, не ЛКМ"),

    ("АПОЛЛО, ООО", "Китай", "Смежные",
     "По названию — неясный профиль. Зарегистрирован в списке как китайская компания с русским ООО.",
     "—", "Не целевой", "Уточнить профиль на стенде"),

    ("Белка", "Россия", "Смежные",
     "Точное соответствие не найдено. Возможно, бренд строительных материалов или химии.",
     "—", "Низкий", "Уточнить профиль на стенде. Название не позволяет однозначно идентифицировать"),

    ("Грамм Лента", "ОАЭ", "Не связано с ЛКМ",
     "Ленты и клейкие материалы.",
     "—", "Не целевой", "Не ЛКМ"),

    ("ДЕ-БАГЕТ, БЕЛЛО ДЕКО", "Россия", "Не связано с ЛКМ",
     "Багет и декоративные элементы.",
     "—", "Не целевой", "Декор, не ЛКМ"),

    ("ДХКЗ ТОО", "Казахстан", "Дистрибьютор ЛКМ",
     "ТОО «ДХКЗ», г. Астана, БИН 230440021866. ОКЭД 46734 — «Оптовая торговля лакокрасочной продукцией». Молодая компания (зарегистрирована в 2023 г.).",
     "ba.prg.kz/710000000-astana/230440021866-too-dkhkz",
     "Высокий", "Прямая цель! Зарегистрирована как оптовый продавец ЛКМ в Астане. Идеальный кандидат в дистрибьюторы Эмпилс на столичный регион"),

    ("КАЗФЛОР ИНЖИНИРИНГ ПРОМЫШЛЕННЫЕ ПОЛЫ", "Казахстан", "Промышленный потребитель",
     "ТОО «ТД Промышленные Полы» — устройство и поставка промышленных наливных полов в Казахстане.",
     "poly-prom.kz/about_us",
     "Средний", "Промышленный потребитель эпоксидных и полиуретановых грунтов/эмалей Эмпилс для устройства полов"),

    ("КОМАНДА МАТЕРИАЛОВ KOOHI", "Иран", "Смежные",
     "Поставщик цемента и строительных материалов из Ирана.",
     "—", "Низкий", "Цемент и строительные материалы, не ЛКМ"),

    ("МАЛАЗИЙСКИЙ СОВЕТ ПО ДРЕВЕСИНЕ", "Малайзия", "Промышленный потребитель",
     "Малазийский совет по древесине — отраслевая ассоциация поставщиков древесных материалов.",
     "—",
     "Низкий", "Не конечный покупатель, но может помочь выйти на производителей мебели/древесины — потребителей ЛКМ по дереву"),

    ("МАХИНА ТСТ КАЗАХСТАН", "Казахстан", "Не связано с ЛКМ",
     "ТСТ Казахстан — insufficient data, по названию — машиностроение.",
     "—", "Не целевой", "Уточнить профиль"),

    ("МТС-СЕРВИС ГРУПП", "Казахстан", "Строительная химия",
     "ТОО «МТС-SERVICE GROUP», Алматы, ул. Стасова, 11. Теплоизоляция для стен, под штукатурку, гидроизоляция.",
     "mtc-sg.com",
     "Средний", "Поставщик строительной химии и теплоизоляции. Может быть интересен как смежный дистрибьютор ЛКМ Эмпилс (органика) в Алматы"),

    ("ПОЛИМЕРМЕТАЛЛТ", "Россия", "Промышленный потребитель",
     "Производитель порошковых красок и оборудования для полимерных порошковых покрытий. Логичный партнёр/клиент по промышленным ЛКМ Эмпилс.",
     "polimerkraska.ru",
     "Средний", "Поставщик порошковых красок — смежная ниша. Возможен обмен клиентами в промышленном сегменте"),

    ("Профиль 35", "Россия", "Промышленный потребитель",
     "Производитель профлиста с полимерным покрытием (НС-35, МП-35). Крупный промышленный потребитель грунт-эмалей по металлу.",
     "almaty.metallprofil.kz",
     "Высокий", "Производитель оцинкованного профлиста с полимерным покрытием — крупный потенциальный промышленный клиент Эмпилс по грунт-эмалям (TITAN)"),

    ("Силиконовый герметик HAOHONG", "Китай", "Герметики/Пены",
     "Zhaoqing Haohong New Material — производитель кислотных, нейтральных и структурных силиконовых герметиков.",
     "gd-haohong.com",
     "Низкий", "Конкурент в нише силиконовых герметиков. Не дистрибьютор ЛКМ"),

    ("СТАРК АЛЬФА ТОО", "Казахстан", "Герметики/Пены",
     "Stark Alpha — первый и единственный производитель двухкомпонентного герметика для стеклопакетов в Казахстане. Связан с Pulver Kazakhstan.",
     "starka.nl/about",
     "Низкий", "Узкая ниша (герметики для стеклопакетов). Не дистрибьютор ЛКМ, но может быть партнёром по комплексным поставкам"),

    ("СТРОЙМАРТ", "Россия", "Дистрибьютор ЛКМ",
     "По данным поиска — поставщик лакокрасочных материалов (нужна проверка, общее название).",
     "—",
     "Средний", "По названию — строительный магазин. Возможный дистрибьютор DIY-сегмента Эмпилс. Требует уточнения на стенде"),

    ("ТССП КАЗАХСТАН", "Казахстан", "Дистрибьютор ЛКМ",
     "По данным поиска — торговля лакокрасочными материалами в Алматы (Первомайская промзона). Возможно, несколько компаний под этим брендом.",
     "sheber-shop.kz/g7966737-lakokrasochnye-materialy",
     "Высокий", "Торговля ЛКМ в Алматы — прямой кандидат в дистрибьюторы Эмпилс. Уточнить на стенде"),

    ("ЭКОПЛИТ", "Узбекистан", "Не связано с ЛКМ",
     "Ecoplit Uzb — производитель строительных материалов (бамбуковые плиты). Ташкент.",
     "facebook.com/ecoplit.uz",
     "Не целевой", "Строительные панели, не ЛКМ"),
]

# === Создание Excel ===

wb = Workbook()
wb.properties.creator = "Z.ai · Анализ Kazbuild 2026"

# === Лист 1: Полный список с анализом ===
ws = wb.active
ws.title = "Анализ участников"

# Заголовок
ws.merge_cells("A1:G1")
ws["A1"] = "Kazbuild 2026 — анализ участников на потенциал дистрибуции ЛКМ Эмпилс"
ws["A1"].font = Font(name="Calibri", size=16, bold=True, color=WHITE)
ws["A1"].fill = PatternFill("solid", fgColor=EMPILS_ORANGE)
ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[1].height = 36

# Подзаголовок
ws.merge_cells("A2:G2")
ws["A2"] = f"Всего компаний в списке: {len(COMPANIES)} · Источник: kazbuild.kz · Анализ от {os.popen('date +%Y-%m-%d').read().strip()}"
ws["A2"].font = Font(name="Calibri", size=10, italic=True, color=DARK)
ws["A2"].fill = PatternFill("solid", fgColor=LIGHT_ORANGE)
ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[2].height = 22

# Заголовки таблицы
HEADERS = ["№", "Наименование компании", "Страна", "Категория", "Информация о компании", "Сайт", "Рекомендация"]
header_row = 4
for col, h in enumerate(HEADERS, 1):
    cell = ws.cell(row=header_row, column=col, value=h)
    cell.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
    cell.fill = PatternFill("solid", fgColor=DARK)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = Border(
        top=Side(style="thin", color="FFFFFF"),
        bottom=Side(style="thin", color="FFFFFF"),
        left=Side(style="thin", color="FFFFFF"),
        right=Side(style="thin", color="FFFFFF"),
    )
ws.row_dimensions[header_row].height = 32

# Данные
thin_border = Border(
    top=Side(style="thin", color="DDD0C8"),
    bottom=Side(style="thin", color="DDD0C8"),
    left=Side(style="thin", color="DDD0C8"),
    right=Side(style="thin", color="DDD0C8"),
)

# Сортируем: Высокий → Средний → Низкий → Не целевой, внутри — по алфавиту
rec_order = {"Высокий": 0, "Средний": 1, "Низкий": 2, "Не целевой": 3}
sorted_companies = sorted(COMPANIES, key=lambda x: (rec_order.get(x[5], 9), x[0]))

for idx, (name, country, cat, info, site, rec, reason) in enumerate(sorted_companies, 1):
    row = header_row + idx
    ws.cell(row=row, column=1, value=idx)
    ws.cell(row=row, column=2, value=name)
    ws.cell(row=row, column=3, value=country or "—")
    ws.cell(row=row, column=4, value=cat)
    ws.cell(row=row, column=5, value=info)
    ws.cell(row=row, column=6, value=site)
    ws.cell(row=row, column=7, value=rec)

    # Стиль строки
    for col in range(1, 8):
        cell = ws.cell(row=row, column=col)
        cell.font = Font(name="Calibri", size=10, color=DARK,
                         bold=(col == 2 or col == 7))
        cell.alignment = Alignment(horizontal="left" if col != 1 and col != 3 else "center",
                                   vertical="center", wrap_text=True)
        cell.border = thin_border

    # Цвет рекомендации
    rec_color = REC_COLORS.get(rec, "FFFFFF")
    ws.cell(row=row, column=7).fill = PatternFill("solid", fgColor=rec_color)
    ws.cell(row=row, column=7).alignment = Alignment(horizontal="center", vertical="center")

    # Цвет категории
    cat_color = CAT_COLORS.get(cat, "FFFFFF")
    ws.cell(row=row, column=4).fill = PatternFill("solid", fgColor=cat_color)
    ws.cell(row=row, column=4).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Высота строки
    ws.row_dimensions[row].height = 50

# Ширины колонок
widths = [5, 35, 14, 22, 55, 35, 14]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

# Freeze panes
ws.freeze_panes = "A5"

# === Лист 2: Легенда ===
ws2 = wb.create_sheet("Легенда")

ws2.merge_cells("A1:C1")
ws2["A1"] = "Легенда категорий и рекомендаций"
ws2["A1"].font = Font(name="Calibri", size=14, bold=True, color=WHITE)
ws2["A1"].fill = PatternFill("solid", fgColor=EMPILS_ORANGE)
ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws2.row_dimensions[1].height = 32

# Категории
ws2["A3"] = "КАТЕГОРИИ КОМПАНИЙ"
ws2["A3"].font = Font(name="Calibri", size=12, bold=True, color=EMPILS_ORANGE)
ws2.merge_cells("A3:C3")

cat_headers = ["Категория", "Цвет", "Что значит"]
for col, h in enumerate(cat_headers, 1):
    c = ws2.cell(row=4, column=col, value=h)
    c.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
    c.fill = PatternFill("solid", fgColor=DARK)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = thin_border
ws2.row_dimensions[4].height = 24

cat_explanations = [
    ("Производитель ЛКМ", "Конкурент Эмпилс — собственное производство красок"),
    ("Дистрибьютор ЛКМ", "Прямая цель — оптовая торговля ЛКМ, может стать каналом сбыта Эмпилс"),
    ("Строительная химия", "Смежная ниша — теплоизоляция, гидроизоляция, наливные полы. Возможен кросс-маркетинг"),
    ("Герметики/Пены", "Конкурент в нише монтажной химии. Обычно не дистрибьюторы ЛКМ"),
    ("Сырьё для ЛКМ", "Поставщики смол, эфиров целлюлозы, модификаторов. Интересны как поставщики, не как каналы сбыта"),
    ("Промышленный потребитель", "Производители профлиста, мебели, древесины — крупные клиенты для промышленных ЛКМ Эмпилс (TITAN, грунт-эмали)"),
    ("Смежные", "Дизайн-студии, торговые компании. Возможны точечные партнёрства"),
    ("Не связано с ЛКМ", "Профиль не пересекается с ЛКМ — алюминий, камень, пластик, обои и т.д."),
]

for i, (cat, desc) in enumerate(cat_explanations, 5):
    ws2.cell(row=i, column=1, value=cat).font = Font(name="Calibri", size=10, bold=True)
    ws2.cell(row=i, column=1).fill = PatternFill("solid", fgColor=CAT_COLORS[cat])
    ws2.cell(row=i, column=2, value="").fill = PatternFill("solid", fgColor=CAT_COLORS[cat])
    ws2.cell(row=i, column=3, value=desc).font = Font(name="Calibri", size=10)
    ws2.cell(row=i, column=3).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    for col in range(1, 4):
        ws2.cell(row=i, column=col).border = thin_border
        ws2.cell(row=i, column=col).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws2.row_dimensions[i].height = 28

# Рекомендации
ws2["A15"] = "УРОВНИ РЕКОМЕНДАЦИИ"
ws2["A15"].font = Font(name="Calibri", size=12, bold=True, color=EMPILS_ORANGE)
ws2.merge_cells("A15:C15")

rec_headers = ["Уровень", "Цвет", "Действие"]
for col, h in enumerate(rec_headers, 1):
    c = ws2.cell(row=16, column=col, value=h)
    c.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
    c.fill = PatternFill("solid", fgColor=DARK)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = thin_border
ws2.row_dimensions[16].height = 24

rec_explanations = [
    ("Высокий", "Связаться на выставке, назначить встречу, подготовить коммерческое предложение"),
    ("Средний", "Посетить стенд, собрать контакты, оценить потенциал после выставки"),
    ("Низкий", "Мониторинг. Конкуренты или нецелевые — проходить мимо, фиксировать для анализа рынка"),
    ("Не целевой", "Не тратить время. Профиль не пересекается с ЛКМ"),
]

for i, (rec, desc) in enumerate(rec_explanations, 17):
    ws2.cell(row=i, column=1, value=rec).font = Font(name="Calibri", size=10, bold=True)
    ws2.cell(row=i, column=1).fill = PatternFill("solid", fgColor=REC_COLORS[rec])
    ws2.cell(row=i, column=2, value="").fill = PatternFill("solid", fgColor=REC_COLORS[rec])
    ws2.cell(row=i, column=3, value=desc).font = Font(name="Calibri", size=10)
    for col in range(1, 4):
        ws2.cell(row=i, column=col).border = thin_border
        ws2.cell(row=i, column=col).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws2.row_dimensions[i].height = 28

# Ширины
ws2.column_dimensions["A"].width = 25
ws2.column_dimensions["B"].width = 10
ws2.column_dimensions["C"].width = 70

# === Лист 3: Резюме и приоритетные цели ===
ws3 = wb.create_sheet("Резюме и приоритеты")

ws3.merge_cells("A1:C1")
ws3["A1"] = "Резюме анализа Kazbuild 2026"
ws3["A1"].font = Font(name="Calibri", size=16, bold=True, color=WHITE)
ws3["A1"].fill = PatternFill("solid", fgColor=EMPILS_ORANGE)
ws3["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws3.row_dimensions[1].height = 36

# Статистика
ws3["A3"] = "Распределение по категориям"
ws3["A3"].font = Font(name="Calibri", size=12, bold=True, color=EMPILS_ORANGE)
ws3.merge_cells("A3:C3")

# Считаем статистику
from collections import Counter
cat_stats = Counter(c[2] for c in COMPANIES)
rec_stats = Counter(c[5] for c in COMPANIES)

# Таблица категорий
ws3.cell(row=4, column=1, value="Категория").font = Font(bold=True, color=WHITE)
ws3.cell(row=4, column=2, value="Кол-во").font = Font(bold=True, color=WHITE)
ws3.cell(row=4, column=3, value="% от списка").font = Font(bold=True, color=WHITE)
for col in range(1, 4):
    ws3.cell(row=4, column=col).fill = PatternFill("solid", fgColor=DARK)
    ws3.cell(row=4, column=col).alignment = Alignment(horizontal="center", vertical="center")
    ws3.cell(row=4, column=col).border = thin_border

total = len(COMPANIES)
cat_order = ["Дистрибьютор ЛКМ", "Производитель ЛКМ", "Строительная химия", "Герметики/Пены",
             "Сырьё для ЛКМ", "Промышленный потребитель", "Смежные", "Не связано с ЛКМ"]
for i, cat in enumerate(cat_order, 5):
    count = cat_stats.get(cat, 0)
    if count == 0:
        continue
    ws3.cell(row=i, column=1, value=cat).font = Font(name="Calibri", size=10)
    ws3.cell(row=i, column=1).fill = PatternFill("solid", fgColor=CAT_COLORS[cat])
    ws3.cell(row=i, column=2, value=count).font = Font(name="Calibri", size=10, bold=True)
    ws3.cell(row=i, column=2).alignment = Alignment(horizontal="center")
    ws3.cell(row=i, column=3, value=f"{count/total*100:.1f}%").font = Font(name="Calibri", size=10)
    ws3.cell(row=i, column=3).alignment = Alignment(horizontal="center")
    for col in range(1, 4):
        ws3.cell(row=i, column=col).border = thin_border
    ws3.row_dimensions[i].height = 24

# Таблица рекомендаций
start_row = 5 + len([c for c in cat_order if cat_stats.get(c, 0) > 0]) + 1
ws3.cell(row=start_row, column=1, value="Распределение по уровню рекомендации").font = Font(name="Calibri", size=12, bold=True, color=EMPILS_ORANGE)
ws3.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=3)

ws3.cell(row=start_row+1, column=1, value="Уровень").font = Font(bold=True, color=WHITE)
ws3.cell(row=start_row+1, column=2, value="Кол-во").font = Font(bold=True, color=WHITE)
ws3.cell(row=start_row+1, column=3, value="Действие").font = Font(bold=True, color=WHITE)
for col in range(1, 4):
    ws3.cell(row=start_row+1, column=col).fill = PatternFill("solid", fgColor=DARK)
    ws3.cell(row=start_row+1, column=col).alignment = Alignment(horizontal="center", vertical="center")
    ws3.cell(row=start_row+1, column=col).border = thin_border

rec_action = {
    "Высокий": "Назначить встречу на Kazbuild",
    "Средний": "Посетить стенд, собрать контакты",
    "Низкий": "Мониторинг для анализа рынка",
    "Не целевой": "Не тратить время",
}
for i, rec in enumerate(["Высокий", "Средний", "Низкий", "Не целевой"], start_row+2):
    count = rec_stats.get(rec, 0)
    ws3.cell(row=i, column=1, value=rec).font = Font(name="Calibri", size=10, bold=True)
    ws3.cell(row=i, column=1).fill = PatternFill("solid", fgColor=REC_COLORS[rec])
    ws3.cell(row=i, column=2, value=count).font = Font(name="Calibri", size=10, bold=True)
    ws3.cell(row=i, column=2).alignment = Alignment(horizontal="center")
    ws3.cell(row=i, column=3, value=rec_action[rec]).font = Font(name="Calibri", size=10)
    for col in range(1, 4):
        ws3.cell(row=i, column=col).border = thin_border
        ws3.cell(row=i, column=col).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws3.row_dimensions[i].height = 24

# Приоритетные цели
target_row = start_row + 7
ws3.cell(row=target_row, column=1, value="🎯 ПРИОРИТЕТНЫЕ ЦЕЛИ — посетить на Kazbuild 2026").font = Font(name="Calibri", size=12, bold=True, color=EMPILS_ORANGE)
ws3.merge_cells(start_row=target_row, start_column=1, end_row=target_row, end_column=3)

priority_companies = [c for c in sorted_companies if c[5] in ("Высокий", "Средний")]

ws3.cell(row=target_row+1, column=1, value="Компания").font = Font(bold=True, color=WHITE)
ws3.cell(row=target_row+1, column=2, value="Страна").font = Font(bold=True, color=WHITE)
ws3.cell(row=target_row+1, column=3, value="Обоснование").font = Font(bold=True, color=WHITE)
for col in range(1, 4):
    ws3.cell(row=target_row+1, column=col).fill = PatternFill("solid", fgColor=DARK)
    ws3.cell(row=target_row+1, column=col).alignment = Alignment(horizontal="center", vertical="center")
    ws3.cell(row=target_row+1, column=col).border = thin_border

for i, (name, country, cat, info, site, rec, reason) in enumerate(priority_companies, target_row+2):
    ws3.cell(row=i, column=1, value=name).font = Font(name="Calibri", size=10, bold=True)
    ws3.cell(row=i, column=2, value=country or "—").font = Font(name="Calibri", size=10)
    ws3.cell(row=i, column=2).alignment = Alignment(horizontal="center")
    ws3.cell(row=i, column=3, value=reason).font = Font(name="Calibri", size=10)
    ws3.cell(row=i, column=3).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws3.cell(row=i, column=1).fill = PatternFill("solid", fgColor=REC_COLORS[rec])
    for col in range(1, 4):
        ws3.cell(row=i, column=col).border = thin_border
    ws3.row_dimensions[i].height = 38

# Ширины
ws3.column_dimensions["A"].width = 40
ws3.column_dimensions["B"].width = 14
ws3.column_dimensions["C"].width = 65

# === Сохранение ===
OUT = "/home/z/my-project/download/Kazbuild_2026_Анализ_участников.xlsx"
wb.save(OUT)
print(f"✅ Файл сохранён: {OUT}")
print(f"   Размер: {os.path.getsize(OUT):,} байт")
print(f"   Компаний в списке: {len(COMPANIES)}")
print(f"   Приоритетных целей (Высокий + Средний): {len(priority_companies)}")
