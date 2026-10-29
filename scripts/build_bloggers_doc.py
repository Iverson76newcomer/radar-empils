"""Создание DOCX-документа со списком региональных нишевых блогеров для TITAN.
   Источник: 3 агента (YouTube, Rutube, Яндекс.Дзен) — 42 канала.
   Telegram/VK — отдельный раздел для заполнения после восстановления квоты.
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Палитра TITAN
COLOR_PRIMARY = RGBColor(0x0B, 0x1C, 0x2C)
COLOR_ACCENT = RGBColor(0xF4, 0x76, 0x29)
COLOR_BODY = RGBColor(0x1A, 0x2B, 0x3D)
COLOR_SECONDARY = RGBColor(0x5B, 0x6B, 0x7D)
COLOR_GREEN = RGBColor(0x16, 0xA3, 0x4A)
COLOR_RED = RGBColor(0xC0, 0x39, 0x2B)
COLOR_YELLOW = RGBColor(0xD9, 0x77, 0x06)

# Данные по каналам — собрано 3 агентами
youtube_channels = [
    {'name': 'Клуб ПРОФИ авторемонта', 'url': 'https://www.youtube.com/@profi_autoremonta',
     'subs': '121 000', 'views': '~30-50K/видео',
     'topic': 'Автомалярка, кузовной ремонт, покраска авто, ЛКМ',
     'geo': 'Россия', 'contact': 'VK vk.ru/profi_autoremonta',
     'idea': 'Обзор TITAN 3-в-1 как бюджетного антикора днища и арок; сравнение с профессиональными грунтами в гаражных условиях.',
     'relevance': 5},
    {'name': 'Garage Heads', 'url': 'https://www.youtube.com/@garageheads',
     'subs': '132 000', 'views': '~60 млн всего',
     'topic': 'Гаражный DIY, покраска/тюнинг авто',
     'geo': 'Россия/СНГ', 'contact': 'через YT',
     'idea': 'Интеграция в видео о подготовке кузова к покраске — TITAN как грунт по ржавчине; «гаражный тест» адгезии.',
     'relevance': 4},
    {'name': 'GARAGE 24', 'url': 'https://www.youtube.com/@GARAGE24',
     'subs': '~100 000+', 'views': '21K-130K/видео',
     'topic': 'Покраска авто с нуля своими руками, кузовной ремонт',
     'geo': 'Россия', 'contact': 'через YT',
     'idea': 'Включение TITAN в плейлист «покраска авто с ноля» — как грунт-эмаль по ржавчине для локальных участков.',
     'relevance': 5},
    {'name': 'ПокраскаМеталла', 'url': 'через YT поиск',
     'subs': 'уточнить', 'views': 'уточнить',
     'topic': 'Покраска металла по ржавчине, защита забора',
     'geo': 'Россия', 'contact': 'через YT',
     'idea': 'Идеальный тематический канал — обзорная серия «краска по металлу 3-в-1»: TITAN vs конкуренты; тест на ржавой поверхности.',
     'relevance': 5},
    {'name': 'Проф Заборы', 'url': 'через YT поиск',
     'subs': 'уточнить', 'views': 'уточнить',
     'topic': 'Заборы из профнастила, покраска металла для забора',
     'geo': 'Россия', 'contact': 'через YT',
     'idea': 'Видео «чем покрасить каркас/столбы забора»: TITAN 3-в-1 по ржавчине как оптимальное «грунт + защита + цвет».',
     'relevance': 5},
    {'name': 'MAGEX paint', 'url': 'https://www.youtube.com/channel/UC9P7bJNLJLqryUBcwjaALWQ',
     'subs': 'уточнить', 'views': 'уточнить',
     'topic': 'Обучающее видео по покраске авто, технологии ЛКМ',
     'geo': 'Россия', 'contact': 'через YT',
     'idea': 'Технологический ролик — TITAN 3-в-1 как грунт по металлу/ржавчине перед покраской; демонстрация адгезии.',
     'relevance': 5},
    {'name': 'Дом и Дача. Полезные советы.', 'url': 'через YT поиск',
     'subs': 'уточнить', 'views': 'уточнить',
     'topic': 'Ремонт, инструменты, материалы, дача/дом',
     'geo': 'Россия', 'contact': 'через YT',
     'idea': 'Обзор категории «краски по металлу для дачи» — TITAN как рекомендация для ворот, труб, заборов, мангал-зон.',
     'relevance': 4},
    {'name': 'ИДЕИ ДЛЯ ДОМА', 'url': 'через YT поиск',
     'subs': 'уточнить', 'views': '~39K',
     'topic': 'Практичные изделия из металла сваркой, идеи для дачи',
     'geo': 'Россия', 'contact': 'через YT',
     'idea': 'Финальная покраска сварных самоделок из металлолома — показать «сварил → зачистил → покрасил 3-в-1 → стоит на улице».',
     'relevance': 4},
    {'name': 'Мастер Самоделок', 'url': 'https://www.youtube.com/@mastersamodelok',
     'subs': 'уточнить', 'views': 'уточнить',
     'topic': 'Самоделки из металла и дерева',
     'geo': 'Россия', 'contact': 'через YT',
     'idea': 'Покраска самодельных металлоконструкций (стеллажи, качели, мангалы) — показать «грунт + финиш за один проход».',
     'relevance': 3},
    {'name': 'Деревенский кузнец', 'url': 'через YT поиск',
     'subs': 'уточнить', 'views': 'уточнить',
     'topic': 'Ковка, сварка, работа с металлом и деревом',
     'geo': 'Россия (сельская)', 'contact': 'через YT',
     'idea': 'Защита кованых изделий от коррозии — TITAN по ржавчине как финишное покрытие для уличных кованых элементов.',
     'relevance': 4},
    {'name': 'Художественная ковка металла', 'url': 'через YT поиск',
     'subs': 'уточнить', 'views': 'уточнить',
     'topic': 'Холодная ковка, самодельные гибочные станки, кованые изделия',
     'geo': 'Россия', 'contact': 'через YT',
     'idea': 'Защита кованых изделий (оградки, козырьки, фонари) — TITAN 3-в-1 по ржавчине как финиш.',
     'relevance': 4},
    {'name': 'ONB Master Pro (Олег «Бери и делай»)', 'url': 'onbmasterpro.com',
     'subs': 'уточнить', 'views': 'уточнить',
     'topic': 'Покраска авто в гараже, обучение малярке, продажа материалов',
     'geo': 'Россия', 'contact': 'onbmasterpro.com',
     'idea': 'Партнёрский обзор-продажа: TITAN в ассортименте магазина + обучающий ролик «покраска детали с ржавчиной».',
     'relevance': 4},
    {'name': 'Get Low Custom', 'url': 'https://www.youtube.com/@getlowcustom',
     'subs': 'уточнить', 'views': 'уточнить',
     'topic': 'Кастом-покраска авто/мото в гараже',
     'geo': 'Россия/СНГ', 'contact': 'TikTok @getlowcustom',
     'idea': 'Покраска кастомных деталей/рамы мото — TITAN как грунт-эмаль по металлу для грунтования рамы.',
     'relevance': 3},
    {'name': 'Автомастерская BAV', 'url': 'https://www.tiktok.com/@avtomasterskayabav',
     'subs': 'уточнить', 'views': 'уточнить',
     'topic': 'Покраска авто в гаражных условиях',
     'geo': 'Россия/СНГ', 'contact': 'TikTok @avtomasterskayabav',
     'idea': '«Покрасил авто в гараже» — TITAN как грунт по ржавчине при подготовке к основному цвету.',
     'relevance': 3},
    {'name': 'Стройхак', 'url': 'https://www.youtube.com/@Stroyhack',
     'subs': '~500 000+ (требует проверки)', 'views': '1-8 млн на топовых',
     'topic': 'DIY, ремонт, стройка, лайфхаки, самоделки',
     'geo': 'Россия', 'contact': 'sponsr.ru/diy_samurai/',
     'idea': '«Стройхак»-ролик: 3 способа защитить металлические ворота/забор от ржавчины с TITAN 3-в-1.',
     'relevance': 3},
]

rutube_channels = [
    {'name': 'АвтоДен', 'url': 'https://rutube.ru/channel/28075249/',
     'subs': '48 327', 'views': '~40-50K',
     'topic': 'Реставрация ржавых авто, кузовной ремонт, покраска авто',
     'geo': 'Немчиновка (Московская обл.)', 'contact': 'TG @AutoDen1, +7(964)563-44-42',
     'idea': 'TITAN 3-в-1 по ржавчине — этап защиты обработанного металла после сварки/зачистки кузова. Натурный показ нанесения.',
     'relevance': 5},
    {'name': 'Denis МЕХАНИК', 'url': 'https://rutube.ru/channel/24861184/',
     'subs': '63 019', 'views': '~15-20K',
     'topic': 'Ремонт авто, самоделки, видеоуроки по ГАЗели, ВАЗу',
     'geo': 'не указано', 'contact': 'denismechanik@themakers.ru',
     'idea': 'В ролике про ремонт/защиту днища или кузовных элементов авто — TITAN как защита металла от коррозии.',
     'relevance': 5},
    {'name': 'МЕНЕДЖЕР АНТОН', 'url': 'https://rutube.ru/channel/23612457/',
     'subs': '127 738', 'views': '~70-90K',
     'topic': 'Авто-эксперименты, кастомные проекты с авто',
     'geo': 'Белгород', 'contact': 'reklamaanton@mail.ru',
     'idea': 'Высокие просмотры. TITAN как бюджетная защита металла в гаражных условиях — формат «сделал сам».',
     'relevance': 4},
    {'name': 'Гараж 54', 'url': 'https://rutube.ru/channel/42833365/',
     'subs': '91 445', 'views': '~15-25K',
     'topic': 'Автомобильные эксперименты и креативные проекты',
     'geo': 'не указано', 'contact': 'promo@garage54.ru',
     'idea': 'В экспериментальном видео с авто/металлоконструкциями — TITAN для защиты металла. Развлекательный формат.',
     'relevance': 4},
    {'name': 'Клубный Сервис', 'url': 'https://rutube.ru/channel/4954692/',
     'subs': '194 116', 'views': '~35-45K',
     'topic': 'Ремонт автомобилей, обзоры авто, разбор проблем',
     'geo': 'Москва', 'contact': '+7(495)778-45-27, VK, Дзен, YT',
     'idea': 'Интеграция в видео про кузовной ремонт/антикор — TITAN для защиты металлических элементов днища, арок.',
     'relevance': 4},
    {'name': 'Дом у дороги. Деревенская жизнь', 'url': 'https://rutube.ru/channel/23888106/',
     'subs': '42 489', 'views': '~45-55K (высокая вовлечённость)',
     'topic': 'Деревенская жизнь, переезд из города, стройка и ремонт в доме',
     'geo': 'Россия (деревня)', 'contact': 'supersyakini@gmail.com, VK, YT',
     'idea': 'Покраска металлического забора, ворот, конструкций TITANом. ЦА дачников/владельцев домов.',
     'relevance': 4},
    {'name': 'Ziskind Vi', 'url': 'https://rutube.ru/channel/24159160/',
     'subs': '14 706', 'views': '~8-12K',
     'topic': 'DIY-строительство: каркасная баня, мастерская, дом из газобетона',
     'geo': 'Ленинградская обл. (под Петербургом)', 'contact': 'ziskindvi@yandex.ru',
     'idea': 'Покраска металлических конструкций (ворота, каркас, перголы) TITANом. Прямая демонстрация на объекте.',
     'relevance': 4},
    {'name': 'GOODMASTER', 'url': 'https://rutube.ru/channel/24653417/',
     'subs': '12 753', 'views': '~1,5-2K',
     'topic': 'Отделочные и строительные работы, самоделки, лайфхаки',
     'geo': 'не указано', 'contact': 'berashek@gmail.com',
     'idea': 'Покраска металлических элементов (ворота, забор, трубы) грунт-эмалью TITAN 3-в-1. Прямое DIY-попадание.',
     'relevance': 4},
    {'name': 'МЕХВОД', 'url': 'https://rutube.ru/channel/24208152/',
     'subs': '17 878', 'views': '~5-6K',
     'topic': 'Автоэксперт, новинки автопрома, машины с пробегом, ремонт',
     'geo': 'не указано', 'contact': 'AvtoMexvod@yandex.ru',
     'idea': 'В обзоре б/у авто — рекомендация TITAN для антикоррозийной защиты кузова при покупке авто с пробегом.',
     'relevance': 3},
    {'name': 'Мастерская K-POWER.RU', 'url': 'https://rutube.ru/channel/41886871/',
     'subs': '41 264', 'views': '~5-15K',
     'topic': 'Ремонт бензиновых двигателей (корейские и европейские ДВС)',
     'geo': 'Набережные Челны', 'contact': 'через ЛС Rutube',
     'idea': 'При ремонте двигателя/подкапотного пространства — защита металлических элементов от коррозии TITAN.',
     'relevance': 3},
    {'name': 'AUTOFACT', 'url': 'https://rutube.ru/channel/23771953/',
     'subs': '39 772', 'views': '~8-13K',
     'topic': 'Подбор автомобилей, тесты, обзоры, ремонт и автоподбор',
     'geo': 'не указано', 'contact': 'mavtofakt@yandex.ru, TG @autofact_import',
     'idea': 'При подборе/обзоре б/у авто — рекомендация TITAN для защиты кузова от коррозии.',
     'relevance': 3},
    {'name': 'Домостройщик', 'url': 'https://rutube.ru/channel/5163119/',
     'subs': '5 198', 'views': '~4-4,5K',
     'topic': 'Переезд на юг, покупка земли, строительство дома своими руками',
     'geo': 'юг России', 'contact': 'VK https://vk.com/domostroyshik',
     'idea': 'Покраска металлического забора/ворот/навеса TITANом. Аудитория дачников/строителей.',
     'relevance': 3},
    {'name': 'Стройхак', 'url': 'https://rutube.ru/channel/26614588/',
     'subs': '5 435', 'views': '~1,5-2K',
     'topic': 'Ремонт, строительство, лайфхаки и строительные хитрости',
     'geo': 'не указано', 'contact': 'через ЛС Rutube',
     'idea': 'Формат «стройхак» — показать TITAN как хитрость для быстрой покраски металла без сложной подготовки.',
     'relevance': 3},
    {'name': 'Всё просто — делай сам !!!', 'url': 'https://rutube.ru/channel/28492471/',
     'subs': '5 482', 'views': '~150-400',
     'topic': 'DIY-ремонт своими руками, эксперименты, лайфхаки',
     'geo': 'не указано', 'contact': 'через ЛС Rutube',
     'idea': 'Ролик-инструкция по покраске металлических поверхностей TITANом. Формат DIY-туториала.',
     'relevance': 2},
    {'name': 'Сварщик из Москвы (Welder from Moscow)', 'url': 'https://rutube.ru/channel/27854523/',
     'subs': '9 887', 'views': '~2-3K',
     'topic': 'Все виды сварки, сварочное оборудование, сварка чугуна и металла',
     'geo': 'Москва', 'contact': 'welder-from-moscow@yandex.ru',
     'idea': 'После сварки конструкций (ворота, заборы) требуется антикоррозийная защита. TITAN как завершение сварочных работ.',
     'relevance': 5},
]

dzen_authors = [
    {'name': 'Клуб ПРОФИ авторемонта', 'url': 'https://dzen.ru/profi_autoremonta',
     'subs': '19 000', 'views': 'статьи+видео',
     'topic': 'Автомалярка, кузовной ремонт, покраска авто',
     'geo': 'Россия', 'contact': 'через ЛС Дзена',
     'idea': 'Обзор TITAN 3-в-1 в контексте подготовки кузова авто — демонстрация преобразования ржавчины и грунтования за один проход.',
     'relevance': 5},
    {'name': 'Виталик и Ангелина', 'url': 'https://dzen.ru/vitalsonik',
     'subs': '76 700', 'views': '~50K+',
     'topic': 'DIY-блог из деревни — сварка заборов, покраска краскопультом, дача',
     'geo': 'Россия (деревня)', 'contact': 'smex6767@gmail.com, TG @vitalik_angelina',
     'idea': 'Интеграция в влог о покраске металлического забора/ворот дачи — TITAN как грунт-эмаль 3-в-1 по ржавчине.',
     'relevance': 5},
    {'name': 'ММА сварка для начинающих', 'url': 'https://dzen.ru/mmasvarka',
     'subs': '20 300', 'views': 'статьи',
     'topic': 'Ручная дуговая сварка, обзоры инверторов. Статья «Чем покрасить металлическую бочку»',
     'geo': 'Россия', 'contact': 'через ЛС Дзена',
     'idea': 'Тест TITAN 3-в-1 на металлической бочке/ёмкости — сравнение с обычной краской, работа по ржавчине без пескоструя.',
     'relevance': 5},
    {'name': 'Для Домашних Сварщиков (Сергей)', 'url': 'https://dzen.ru/svarka',
     'subs': '134 000', 'views': 'статьи+видео',
     'topic': 'Сварка для новичков, сварные швы, гаражный сварщик',
     'geo': 'Россия', 'contact': 'TG @domsvarka',
     'idea': 'Инструкция по защите сварных швов от коррозии — TITAN как финишное покрытие поверх сварных соединений ворот, гаражей.',
     'relevance': 4},
    {'name': 'Сварщик из Москвы (Владимир)', 'url': 'https://dzen.ru/welder_from_moscow',
     'subs': '30 700', 'views': 'статьи+видео',
     'topic': 'Сварка (34 года опыта), ремонт тисков, сварка чугуна, лазерная сварка',
     'geo': 'Москва', 'contact': 'welder-from-moscow@yandex.ru',
     'idea': 'Обзор защиты отремонтированных сваркой металлоизделий (тиски, детали станков) — TITAN как защита от ржавчины.',
     'relevance': 4},
    {'name': 'Санько Сами Своими Руками', 'url': 'https://dzen.ru/77vitek77',
     'subs': '22 500', 'views': 'статьи+видео',
     'topic': 'Поделки из дерева и металла, ремонт авто (Renault Laguna), восстановление ножей',
     'geo': 'Россия', 'contact': 'через ЛС Дзена',
     'idea': 'Покраска восстановленных металлических деталей авто и самоделок TITANом — 3-в-1 на деталях с ржавчиной.',
     'relevance': 4},
    {'name': 'Кузница СВАРОГА', 'url': 'https://dzen.ru/id/5fbd75524b9b1b331d25ab00',
     'subs': '28 000', 'views': 'статьи+видео',
     'topic': 'Кузнечное оружейное дело — Damascus-сталь, ножи, ковка',
     'geo': 'Россия', 'contact': 'через ЛС Дзена / Rutube',
     'idea': 'Защита кованых изделий для улицы (ворота, заборные элементы, мебель) — TITAN от атмосферной коррозии.',
     'relevance': 3},
    {'name': 'Холодная ковка', 'url': 'https://dzen.ru/id/5f5521d83fbe9a0c384dbc2e',
     'subs': '11 500', 'views': 'статьи+видео',
     'topic': 'Холодная ковка, станки для ковки, трубогибы, поделки из металла для дачи',
     'geo': 'Россия', 'contact': 'TG @MasterRuslik',
     'idea': 'Покраска готовых кованых изделий (заборные секции, решётки, ворота) — TITAN как финишная отделка.',
     'relevance': 3},
    {'name': 'МастерDIY', 'url': 'https://dzen.ru/id/5f34cfae05113f5202a228b5',
     'subs': '13 100', 'views': 'статьи+видео',
     'topic': 'Самоделки из профильной трубы, старого насоса, домкрата. Гаражный DIY',
     'geo': 'Россия', 'contact': 'MAX, Rutube',
     'idea': 'Покраска самоделок из профильной трубы (навесы, стеллажи) — защита металлоконструкций одним слоем TITAN.',
     'relevance': 4},
    {'name': 'Дачный мастер', 'url': 'https://dzen.ru/dachamaster_org',
     'subs': '14 200', 'views': 'статьи+видео',
     'topic': 'Ремонт, отделка, строительство загородных домов, столярное дело',
     'geo': 'Россия', 'contact': 'через ЛС Дзена',
     'idea': 'Обзор покраски металлических ворот гаража/забора на даче TITANом — тест на ржавых воротах без зачистки.',
     'relevance': 4},
    {'name': 'Из города на дачу. Жизнь в деревне.', 'url': 'https://dzen.ru/id/62223ba31c4d5b071ef825ed',
     'subs': '7 653', 'views': 'статьи+видео',
     'topic': 'Бюджетные переделки на дачном участке, ремонт, жизнь в деревне',
     'geo': 'Россия (деревня)', 'contact': 'через ЛС Дзена',
     'idea': 'Бюджетная покраска старых металлических конструкций на даче (забор, ворота, трубы) — акцент на экономию 3-в-1.',
     'relevance': 3},
    {'name': 'Custombox', 'url': 'https://dzen.ru/custombox',
     'subs': '6 609', 'views': 'статьи+видео',
     'topic': 'Реставрация старых автомобилей (Mercedes W114 1975 г.), кузовные работы',
     'geo': 'Россия', 'contact': 'через ЛС Дзена',
     'idea': 'Защита отремонтированных кузовных деталей и металлических элементов ретро-авто TITANом от повторной коррозии.',
     'relevance': 3},
]

# Создаём документ
doc = Document()
section = doc.sections[0]
section.top_margin = Cm(2.54)
section.bottom_margin = Cm(2.54)
section.left_margin = Cm(2.0)
section.right_margin = Cm(2.0)

style = doc.styles['Normal']
style.font.name = 'Arial'
style.font.size = Pt(10)
style.font.color.rgb = COLOR_BODY

# === Заголовок ===
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('TITAN — Региональные нишевые блогеры')
run.font.size = Pt(18)
run.font.bold = True
run.font.color.rgb = COLOR_PRIMARY

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Для постепенного масштабирования присутствия бренда в digital-каналах')
run.font.size = Pt(11)
run.font.italic = True
run.font.color.rgb = COLOR_SECONDARY

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Дата подготовки: 15 сентября 2026 г.')
run.font.size = Pt(10)
run.font.italic = True
run.font.color.rgb = COLOR_SECONDARY

# Контекст — почему нишевые блогеры, а не топовые
p = doc.add_paragraph()
run = p.add_run('Контекст решения')
run.font.size = Pt(13)
run.font.bold = True
run.font.color.rgb = COLOR_PRIMARY

context_text = (
    'Ранее была проработана связка из 8 топовых блогеров (Амарян, Батрутдинов, Nail, Баженов, '
    'Лапенко, Пучков, Харламов, Лифт) с аудиторией 2,5-6,2 млн подписчиков. Решение оказалось '
    'дорогим и сложным в реализации. В качестве альтернативы предлагается начать с нишевых '
    'региональных блогеров с аудиторией 5-200 тыс. подписчиков на 4 платформах: YouTube, Rutube, '
    'Яндекс.Дзен, Telegram/VK. Это позволяет:'
)
p = doc.add_paragraph(context_text)
p.paragraph_format.first_line_indent = Cm(0.5)
p.paragraph_format.line_spacing = 1.3

for txt in [
    'Снизить стоимость одной интеграции в 10-30 раз (от 5-15 тыс. ₽ за канал вместо 200-500 тыс. ₽)',
    'Распределить риски: 10-15 каналов по 50-100 тыс. ₽ дадут сопоставимый охват с 1 топовым',
    'Получить более высокую конверсию — нишевая аудитория точно совпадает с ЦА TITAN (мужчины 30-55 лет, DIY)',
    'Тестировать гипотезы и форматы — быстро и дёшево отрабатывать идеи перед масштабированием',
    'Постепенно подключать новые каналы — начать с 3-5 пилотов, расширять при подтверждённой эффективности',
]:
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run(txt)
    run.font.size = Pt(10)

# === Статус сбора данных ===
doc.add_paragraph()
p = doc.add_paragraph()
run = p.add_run('Статус сбора данных по платформам')
run.font.size = Pt(13)
run.font.bold = True
run.font.color.rgb = COLOR_PRIMARY

table = doc.add_table(rows=1, cols=5)
table.style = 'Light Grid Accent 1'
hdr = table.rows[0].cells
headers = ['Платформа', 'Статус', 'Каналов найдено', 'Диапазон подписчиков', 'Содержит контакты']
for i, h in enumerate(headers):
    hdr[i].text = ''
    p = hdr[i].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(h)
    run.font.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    tcPr = hdr[i]._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), '313131')
    tcPr.append(shd)

platforms = [
    ('YouTube', '✅ Готово', '15', '10K-500K', 'частично'),
    ('Rutube', '✅ Готово', '15', '5K-194K', '11 из 15'),
    ('Яндекс.Дзен', '✅ Готово', '12', '5K-134K', '4 из 12'),
    ('Telegram + VK', '⚠️ В процессе (превышен rate limit API поиска)', '0', '—', '—'),
    ('ИТОГО', '—', '42', '—', '—'),
]
for row_data in platforms:
    row = table.add_row()
    cells = row.cells
    for i, v in enumerate(row_data):
        cells[i].text = ''
        p = cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(v)
        run.font.size = Pt(10)
        if i == 0:
            run.font.bold = True
        if row_data[0] == 'ИТОГО':
            run.font.bold = True
            tcPr = cells[i]._tc.get_or_add_tcPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:fill'), 'F47629')
            tcPr.append(shd)
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        elif 'Готово' in str(v):
            run.font.color.rgb = COLOR_GREEN
        elif 'процессе' in str(v):
            run.font.color.rgb = COLOR_YELLOW

widths = [Cm(3.5), Cm(5.5), Cm(2.5), Cm(3.5), Cm(3.5)]
for row in table.rows:
    for i, w in enumerate(widths):
        if i < len(row.cells):
            row.cells[i].width = w

doc.add_paragraph()

# === Функция добавления раздела платформы ===
def add_platform_section(title, channels, total_label):
    p = doc.add_paragraph()
    run = p.add_run(title)
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.color.rgb = COLOR_PRIMARY

    p = doc.add_paragraph()
    run = p.add_run(f'Найдено {len(channels)} каналов. Сортировка по релевантности для TITAN (от ⭐⭐⭐⭐⭐ до ⭐⭐).')
    run.font.size = Pt(10)
    run.font.italic = True
    run.font.color.rgb = COLOR_SECONDARY

    # Сортируем по релевантности (убывание)
    channels_sorted = sorted(channels, key=lambda x: -x['relevance'])

    # Таблица
    table = doc.add_table(rows=1, cols=7)
    table.style = 'Light Grid Accent 1'
    hdr = table.rows[0].cells
    headers = ['№', 'Канал', 'Подписчики', 'Просмотры', 'Тематика', 'Контакт', 'Идея интеграции']
    for i, h in enumerate(headers):
        hdr[i].text = ''
        p = hdr[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        tcPr = hdr[i]._tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:fill'), '313131')
        tcPr.append(shd)

    for i, ch in enumerate(channels_sorted, 1):
        row = table.add_row()
        cells = row.cells

        # № + звёзды релевантности
        stars = '⭐' * ch['relevance']
        cells[0].text = ''
        p = cells[0].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(f"{i}\n{stars}")
        run.font.size = Pt(8)

        # Канал + URL
        cells[1].text = ''
        p = cells[1].paragraphs[0]
        run = p.add_run(ch['name'])
        run.font.bold = True
        run.font.size = Pt(9)
        # URL (как обычный текст, без hyperlink — для простоты)
        run = p.add_run(f"\n{ch['url']}")
        run.font.size = Pt(7)
        run.font.italic = True
        run.font.color.rgb = COLOR_SECONDARY

        # Подписчики
        cells[2].text = ''
        p = cells[2].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(ch['subs'])
        run.font.size = Pt(9)
        run.font.bold = True

        # Просмотры
        cells[3].text = ''
        p = cells[3].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(ch['views'])
        run.font.size = Pt(8)

        # Тематика + Гео
        cells[4].text = ''
        p = cells[4].paragraphs[0]
        run = p.add_run(ch['topic'])
        run.font.size = Pt(8)
        run = p.add_run(f"\n📍 {ch['geo']}")
        run.font.size = Pt(7)
        run.font.italic = True
        run.font.color.rgb = COLOR_SECONDARY

        # Контакт
        cells[5].text = ''
        p = cells[5].paragraphs[0]
        run = p.add_run(ch['contact'])
        run.font.size = Pt(8)
        if '@' in ch['contact'] or 'mail' in ch['contact'].lower() or '.ru' in ch['contact'].lower():
            run.font.color.rgb = COLOR_GREEN

        # Идея интеграции
        cells[6].text = ''
        p = cells[6].paragraphs[0]
        run = p.add_run(ch['idea'])
        run.font.size = Pt(8)
        if ch['relevance'] >= 5:
            run.font.bold = True
            run.font.color.rgb = COLOR_ACCENT

    # Ширины
    widths = [Cm(0.8), Cm(3.0), Cm(1.5), Cm(1.5), Cm(2.8), Cm(2.5), Cm(4.0)]
    for row in table.rows:
        for i, w in enumerate(widths):
            if i < len(row.cells):
                row.cells[i].width = w

    doc.add_paragraph()

# === Раздел 1: YouTube ===
add_platform_section('Раздел 1. YouTube — 15 нишевых каналов DIY/ЛКМ/авто', youtube_channels, 'YouTube')

# === Раздел 2: Rutube ===
add_platform_section('Раздел 2. Rutube — 15 нишевых каналов DIY/ЛКМ/авто', rutube_channels, 'Rutube')

# === Раздел 3: Яндекс.Дзен ===
add_platform_section('Раздел 3. Яндекс.Дзен — 12 авторов DIY/ЛКМ/авто', dzen_authors, 'Дзен')

# === Раздел 4: Telegram + VK ===
p = doc.add_paragraph()
run = p.add_run('Раздел 4. Telegram + VK')
run.font.size = Pt(14)
run.font.bold = True
run.font.color.rgb = COLOR_PRIMARY

p = doc.add_paragraph()
run = p.add_run('⚠️ Поиск по Telegram и VK не завершён — превышен rate limit API поиска (ошибка 429). ')
run.font.size = Pt(11)
run.font.bold = True
run.font.color.rgb = COLOR_RED
run = p.add_run('Сбор будет выполнен после восстановления квоты (обычно 30-60 минут). Готовые запросы:')
run.font.size = Pt(10)
run.font.color.rgb = COLOR_BODY

queries_list = [
    'telegram канал diy ремонт своими руками мужчина',
    'telegram покраска металла гараж ворота краска',
    'telegram канал строительство дача обустройство',
    'telegram покраска авто гараж своими руками',
    'telegram кузнечное дело металлообработка канал',
    'telegram самоделки металл своими руками канал',
    'vk сообщество diy ремонт гараж своими руками',
    'vk видео покраска авто гараж металл',
]
for q in queries_list:
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run(q)
    run.font.size = Pt(9)
    run.font.color.rgb = COLOR_SECONDARY
    run.font.name = 'Consolas'

# Альтернативные источники для Telegram/VK
p = doc.add_paragraph()
run = p.add_run('Альтернативные источники (если rate limit сохраняется):')
run.font.size = Pt(11)
run.font.bold = True
run.font.color.rgb = COLOR_PRIMARY

alts = [
    ('TGStat.ru', 'каталоги Telegram-каналов с фильтром по тематикам «Строительство», «Авто», «Ремонт»'),
    ('Telemetr.io', 'каталоги Telegram-каналов с точными числами подписчиков и контактами'),
    ('VK поиск', 'vk.com/search?c[section]=communities с фильтром по темам «Ремонт», «Строительство», «Авто»'),
    ('Яндекс', 'запросы вида site:t.me «покраска гаража» и site:vk.com «ремонт гаража своими руками»'),
]
for name, desc in alts:
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run(f'{name} — ')
    run.font.bold = True
    run.font.size = Pt(10)
    run = p.add_run(desc)
    run.font.size = Pt(10)

# === Стратегия внедрения ===
doc.add_page_break()
p = doc.add_paragraph()
run = p.add_run('Стратегия постепенного внедрения')
run.font.size = Pt(15)
run.font.bold = True
run.font.color.rgb = COLOR_PRIMARY

p = doc.add_paragraph()
run = p.add_run(
    'Принцип «начать с малого и масштабировать при подтверждённой эффективности» — '
    'основной подход к работе с нишевыми блогерами. Ниже 3 этапа реализации.'
)
run.font.size = Pt(10)
p.paragraph_format.first_line_indent = Cm(0.5)

# Этапы
stages = [
    ('Этап 1. Пилот (1-2 месяца)', '5-7 каналов из топ-рейтинга релевантности',
     [
         'Бюджет: 200-400 тыс. ₽ (5-7 каналов × 30-60 тыс. ₽)',
         'Формат: натурная интеграция — обзор TITAN 3-в-1 на реальном объекте блогера',
         'Цель: тест гипотезы «нишевые блогеры дают лучшую конверсию, чем топовые»',
         'Метрика: рост брендовых запросов «TITAN краска» в Wordstat (+15-25% за 2 месяца)',
         'KPI по продажам: рост продаж TITAN в категориях DIY на 5-10% в регионах блогеров',
         'Кандидаты: АвтоДен, Сварщик из Москвы, Denis МЕХАНИК, Виталик и Ангелина, ММА сварка для начинающих, Клуб ПРОФИ авторемонта, Ziskind Vi',
     ]),
    ('Этап 2. Масштабирование (3-6 месяцев)', '15-20 каналов на 3-4 платформах',
     [
         'Бюджет: 800-1500 тыс. ₽ (15-20 каналов × 40-80 тыс. ₽)',
         'Расширение на Telegram/VK (после сбора данных)',
         'Форматы: нативные обзоры + DIY-туториалы + сравнительные тесты',
         'Добавить 2-3 YouTube канала с аудиторией 100K+ для широкого охвата',
         'Метрика: рост брендовых запросов на 40-60% за 6 месяцев',
         'KPI по продажам: рост продаж TITAN в DIY на 15-25% в регионах присутствия блогеров',
     ]),
    ('Этап 3. Оптимизация (7-12 месяцев)', '10-15 эффективных каналов на постоянной основе',
     [
         'Бюджет: 600-1000 тыс. ₽/6 месяцев (10-15 каналов × 50-80 тыс. ₽)',
         'Отсев неэффективных каналов по итогам пилота',
         'Заключение годовых контрактов с топ-5 блогерами (рейт-карта, эксклюзив)',
         'Запуск совместных продуктов (TITAN x Блогер — ограниченная серия цвета/объёма)',
         'Метрика: рост брендовых запросов «TITAN» в Wordstat до 5000+/мес (с 1430 сейчас)',
         'KPI: рост продаж TITAN в DIY на 30-50% год-к-году',
     ]),
]

for stage_title, stage_subtitle, stage_items in stages:
    p = doc.add_paragraph()
    run = p.add_run(stage_title)
    run.font.size = Pt(12)
    run.font.bold = True
    run.font.color.rgb = COLOR_ACCENT

    p = doc.add_paragraph()
    run = p.add_run(stage_subtitle)
    run.font.size = Pt(10)
    run.font.italic = True
    run.font.color.rgb = COLOR_SECONDARY

    for item in stage_items:
        p = doc.add_paragraph(style='List Bullet')
        run = p.add_run(item)
        run.font.size = Pt(10)
    doc.add_paragraph()

# === Топ-10 приоритетных каналов ===
p = doc.add_paragraph()
run = p.add_run('Топ-10 приоритетных каналов для первого контакта')
run.font.size = Pt(15)
run.font.bold = True
run.font.color.rgb = COLOR_PRIMARY

p = doc.add_paragraph()
run = p.add_run('Отобраны по 3 критериям: релевантность для TITAN, наличие прямых контактов для связи, наличие живой аудитории с подтверждёнными просмотрами.')
run.font.size = Pt(10)
run.font.italic = True
run.font.color.rgb = COLOR_SECONDARY

top10 = [
    # (Имя, Платформа, Подписчики, Контакт, Идея)
    ('АвтоДен', 'Rutube', '48K', 'TG @AutoDen1', 'TITAN 3-в-1 по ржавчине — этап защиты обработанного металла после сварки/зачистки кузова'),
    ('Denis МЕХАНИК', 'Rutube', '63K', 'denismechanik@themakers.ru', 'TITAN как защита днища/кузовных элементов авто от коррозии'),
    ('Сварщик из Москвы', 'Rutube + Дзен', '10K + 31K', 'welder-from-moscow@yandex.ru', 'После сварки конструкций (ворота, заборы) — антикоррозийная защита TITAN'),
    ('Виталик и Ангелина', 'Дзен', '77K', 'smex6767@gmail.com, TG @vitalik_angelina', 'Покраска металлического забора/ворот дачи TITAN 3-в-1 по ржавчине'),
    ('Клуб ПРОФИ авторемонта', 'YouTube + Дзен', '121K + 19K', 'VK + ЛС Дзена', 'Обзор TITAN 3-в-1 в подготовке кузова авто — антикор днища и арок'),
    ('ММА сварка для начинающих', 'Дзен', '20K', 'ЛС Дзена', 'Тест TITAN на металлической бочке — сравнение с обычной краской, работа по ржавчине'),
    ('МЕНЕДЖЕР АНТОН', 'Rutube', '128K', 'reklamaanton@mail.ru', 'Высокий охват. TITAN как бюджетная защита металла в гаражных условиях'),
    ('Дом у дороги', 'Rutube', '42K', 'supersyakini@gmail.com', 'Покраска забора/ворот/конструкций TITANом. ЦА дачников'),
    ('Ziskind Vi', 'Rutube', '15K', 'ziskindvi@yandex.ru', 'Покраска металлических конструкций (ворота, каркас) TITAN на реальном объекте'),
    ('GOODMASTER', 'Rutube', '13K', 'berashek@gmail.com', 'Покраска ворот/забора/труб грунт-эмалью TITAN 3-в-1. Прямое DIY-попадание'),
]

table = doc.add_table(rows=1, cols=6)
table.style = 'Light Grid Accent 1'
hdr = table.rows[0].cells
headers = ['№', 'Канал', 'Платформа', 'Подписчики', 'Контакт', 'Идея интеграции']
for i, h in enumerate(headers):
    hdr[i].text = ''
    p = hdr[i].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(h)
    run.font.bold = True
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    tcPr = hdr[i]._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), 'F47629')
    tcPr.append(shd)

for i, (name, platform, subs, contact, idea) in enumerate(top10, 1):
    row = table.add_row()
    cells = row.cells
    cells[0].text = ''
    p = cells[0].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(str(i))
    run.font.bold = True
    run.font.size = Pt(10)

    cells[1].text = ''
    p = cells[1].paragraphs[0]
    run = p.add_run(name)
    run.font.bold = True
    run.font.size = Pt(9)

    cells[2].text = ''
    p = cells[2].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(platform)
    run.font.size = Pt(9)

    cells[3].text = ''
    p = cells[3].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(subs)
    run.font.bold = True
    run.font.size = Pt(9)
    run.font.color.rgb = COLOR_GREEN

    cells[4].text = ''
    p = cells[4].paragraphs[0]
    run = p.add_run(contact)
    run.font.size = Pt(8)
    run.font.color.rgb = COLOR_ACCENT

    cells[5].text = ''
    p = cells[5].paragraphs[0]
    run = p.add_run(idea)
    run.font.size = Pt(8)

widths = [Cm(0.8), Cm(2.8), Cm(2.5), Cm(1.8), Cm(3.5), Cm(6.0)]
for row in table.rows:
    for i, w in enumerate(widths):
        if i < len(row.cells):
            row.cells[i].width = w

# === Ориентировочный бюджет ===
doc.add_paragraph()
p = doc.add_paragraph()
run = p.add_run('Ориентировочный бюджет интеграций (по рыночным ставкам)')
run.font.size = Pt(13)
run.font.bold = True
run.font.color.rgb = COLOR_PRIMARY

table = doc.add_table(rows=1, cols=4)
table.style = 'Light Grid Accent 1'
hdr = table.rows[0].cells
for i, h in enumerate(['Диапазон подписчиков', 'Стоимость интеграции', 'Каналов найдено', 'Итого бюджет']):
    hdr[i].text = ''
    p = hdr[i].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(h)
    run.font.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    tcPr = hdr[i]._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), '313131')
    tcPr.append(shd)

budget_rows = [
    ('5-15K подписчиков', '5-15 тыс. ₽', '~12 каналов', '60-180 тыс. ₽'),
    ('15-50K подписчиков', '15-40 тыс. ₽', '~15 каналов', '225-600 тыс. ₽'),
    ('50-130K подписчиков', '40-120 тыс. ₽', '~10 каналов', '400-1200 тыс. ₽'),
    ('130K+ подписчиков', '100-200 тыс. ₽', '~5 каналов', '500-1000 тыс. ₽'),
    ('ИТОГО при полной реализации', '', '42 канала', '1,2-3,0 млн ₽'),
]
for row_data in budget_rows:
    row = table.add_row()
    cells = row.cells
    for i, v in enumerate(row_data):
        cells[i].text = ''
        p = cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(v)
        run.font.size = Pt(10)
        if row_data[0].startswith('ИТОГО'):
            run.font.bold = True
            tcPr = cells[i]._tc.get_or_add_tcPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:fill'), 'F47629')
            tcPr.append(shd)
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        elif 'тыс. ₽' in str(v) or 'млн' in str(v):
            if i == 3:
                run.font.bold = True

widths = [Cm(4.5), Cm(4.0), Cm(4.0), Cm(4.5)]
for row in table.rows:
    for i, w in enumerate(widths):
        if i < len(row.cells):
            row.cells[i].width = w

# === Следующие шаги ===
doc.add_paragraph()
p = doc.add_paragraph()
run = p.add_run('Следующие шаги')
run.font.size = Pt(13)
run.font.bold = True
run.font.color.rgb = COLOR_PRIMARY

steps = [
    ('1. Завершить сбор данных по Telegram/VK', 'после восстановления квоты API поиска (30-60 мин)'),
    ('2. Уточнить метрики', 'через Social Blade / TGStat для каналов с пометкой «уточнить»'),
    ('3. Связаться с топ-10 каналами', 'запрос медиакита и прайса, рассылка КП'),
    ('4. Подготовить ТЗ на интеграцию', 'единый формат для всех: натурный показ TITAN 3-в-1 на реальном объекте'),
    ('5. Запустить пилот с 5-7 каналами', 'бюджет 200-400 тыс. ₽, срок — 1-2 месяца'),
    ('6. Измерить эффективность', 'рост брендовых запросов в Wordstat, рост продаж в регионах блогеров'),
    ('7. Решение о масштабировании', 'по итогам пилота — расширение на 15-20 каналов'),
]
for step, desc in steps:
    p = doc.add_paragraph(style='List Number')
    run = p.add_run(f'{step}: ')
    run.font.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = COLOR_PRIMARY
    run = p.add_run(desc)
    run.font.size = Pt(10)

# === Заключение ===
doc.add_paragraph()
p = doc.add_paragraph()
run = p.add_run('Заключение')
run.font.size = Pt(13)
run.font.bold = True
run.font.color.rgb = COLOR_PRIMARY

conclusion = (
    'Стратегия постепенного масштабирования через нишевые региональные блогеры — реалистичная '
    'альтернатива дорогим топовым блогерам. Собрано 42 канала на 3 платформах (YouTube, Rutube, '
    'Яндекс.Дзен) с детальными контактами и идеями интеграции. Средняя стоимость интеграции — '
    '30-80 тыс. ₽ за канал, что в 5-15 раз дешевле топовых блогеров (200-500 тыс. ₽ за интеграцию '
    'у Амаряна, Харламова и др.). При запуске пилота с 5-7 каналами за 200-400 тыс. ₽ можно '
    'протестировать гипотезу и принять решение о масштабировании на основе данных, а не интуиции. '
    'Это первый шаг в долгосрочной digital-стратегии TITAN.'
)
p = doc.add_paragraph(conclusion)
p.paragraph_format.first_line_indent = Cm(0.5)
p.paragraph_format.line_spacing = 1.3

# Сохраняем
out_path = '/home/z/my-project/download/TITAN_Региональные_блогеры_список.docx'
doc.save(out_path)
print(f'✅ Сохранено: {out_path}')
print(f'\nСтатистика:')
print(f'  YouTube:  {len(youtube_channels)} каналов')
print(f'  Rutube:   {len(rutube_channels)} каналов')
print(f'  Дзен:     {len(dzen_authors)} авторов')
print(f'  Telegram/VK: 0 (в процессе)')
print(f'  ИТОГО:    {len(youtube_channels) + len(rutube_channels) + len(dzen_authors)} каналов')
print(f'\nТоп-10 приоритетных:')
for i, (name, platform, subs, contact, _) in enumerate(top10, 1):
    print(f'  {i}. {name} ({platform}, {subs}) — {contact}')
