/**
 * Аналитическая записка на 1 лист A4:
 * "Анализ выхода Эмпилс на рынок аэрозольных ЛКМ"
 *
 * Без обложки, без TOC. Компактный плотный документ.
 * Поля узкие, шрифт 8-9pt, чтобы уместить всё на 1 страницу.
 */

const {
  Document, Packer, Paragraph, TextRun,
  AlignmentType, HeadingLevel, BorderStyle,
  Table, TableRow, TableCell, WidthType, ShadingType,
  TableLayoutType,
} = require("docx");
const fs = require("fs");

const P = {
  primary:  "1A2330",
  body:     "1F2937",
  secondary:"6B7280",
  accent:   "D4875A",
  accentLight: "F8F0EB",
  surface:  "F5F5F5",
  positive: "2E7D32",
  danger:   "C62828",
  white:    "FFFFFF",
};
const c = (hex) => hex.replace("#", "");
const FONT = { ascii: "Calibri", eastAsia: "Microsoft YaHei" };
const NB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 60, after: 30, line: 200 },
    children: [new TextRun({ text, bold: true, size: 16, color: c(P.primary), font: FONT })],
  });
}

function body(text, size = 14) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 210, after: 30 },
    children: [new TextRun({ text, size, color: c(P.body), font: FONT })],
  });
}

function bullet(text, size = 14) {
  return new Paragraph({
    indent: { left: 200, hanging: 150 },
    spacing: { line: 190, after: 15 },
    children: [
      new TextRun({ text: "\u25B8 ", size, color: c(P.accent), bold: true, font: FONT }),
      new TextRun({ text, size, color: c(P.body), font: FONT }),
    ],
  });
}

function bulletBold(lead, rest, size = 14) {
  return new Paragraph({
    indent: { left: 200, hanging: 150 },
    spacing: { line: 190, after: 15 },
    children: [
      new TextRun({ text: "\u25B8 ", size, color: c(P.accent), bold: true, font: FONT }),
      new TextRun({ text: lead, size, bold: true, color: c(P.primary), font: FONT }),
      new TextRun({ text: " \u2014 " + rest, size, color: c(P.body), font: FONT }),
    ],
  });
}

function divider() {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: c(P.accent), space: 1 } },
    children: [],
  });
}

// === Таблица: линейки и приоритеты ===
function buildLineupTable() {
  const headerCells = ["Линейка", "Связующее", "RAL", "Степень блеска", "Приоритет", "Маржа", "Скорость продаж"];
  const rows = [
    ["Универсальная алкидная", "Алкидная смола", "По RAL", "Глянец", "1 (первая)", "Высокая (40-50%)", "Быстро — DIY + авто"],
    ["Грунт-эмаль алкидная", "Алкид + антикорр. добавки", "Не по RAL", "Полуглянец", "1 (первая)", "Высокая (45-55%)", "Быстро — B2B + DIY"],
    ["Акриловая универсальная", "Акриловая дисперсия", "По RAL", "Матовый / полуматовый", "2 (вторая)", "Средняя (30-40%)", "Средне — творчество + DIY"],
  ];

  const headerRow = new TableRow({
    tableHeader: true, cantSplit: true,
    children: headerCells.map(text => new TableCell({
      shading: { type: ShadingType.CLEAR, fill: c(P.primary) },
      margins: { top: 40, bottom: 40, left: 60, right: 60 },
      borders: { top: { style: BorderStyle.SINGLE, size: 3, color: c(P.primary) }, bottom: { style: BorderStyle.SINGLE, size: 3, color: c(P.primary) }, left: NB, right: NB },
      children: [new Paragraph({
        spacing: { line: 200 },
        children: [new TextRun({ text, bold: true, size: 14, color: c(P.white), font: FONT })],
      })],
    })),
  });

  const dataRows = rows.map((row, idx) => new TableRow({
    cantSplit: true,
    children: row.map((text, ci) => new TableCell({
      shading: { type: ShadingType.CLEAR, fill: idx % 2 === 0 ? "FAFAFA" : c(P.white) },
      margins: { top: 30, bottom: 30, left: 60, right: 60 },
      borders: { top: NB, bottom: { style: BorderStyle.SINGLE, size: 2, color: "E0E0E0" }, left: NB, right: NB },
      children: [new Paragraph({
        spacing: { line: 200 },
        children: [new TextRun({
          text, size: 14,
          bold: ci === 0 || ci === 4 || ci === 5,
          color: c(ci === 4 ? P.accent : (ci === 5 ? P.positive : P.body)),
          font: FONT,
        })],
      })],
    })),
  }));

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 3, color: c(P.primary) },
      bottom: { style: BorderStyle.SINGLE, size: 3, color: c(P.primary) },
      left: NB, right: NB, insideHorizontal: NB, insideVertical: NB,
    },
    rows: [headerRow, ...dataRows],
  });
}

// === Таблица: конкуренты ===
function buildCompetitorsTable() {
  const headerCells = ["Бренд", "Страна", "Сегмент", "Цена за баллон", "Ассортимент"];
  const rows = [
    ["BOSNY", "Тайланд", "Средний +", "350-500 ₽", "77+ SKU, акрил + алкид, RAL"],
    ["KUDO", "Россия", "Средний", "250-400 ₽", "8 серий: AUTO, ARTE, INDUSTRY, HOME"],
    ["NewTon (Эмпилс)", "Россия", "Эконом-средний", "200-350 ₽", "Ограниченный, RAL"],
    ["Krylon (Rust-Oleum)", "США", "Премиум", "600-900 ₽", "Широкий, спецэффекты"],
    ["Motip / ColorMatic", "Нидерланды", "Премиум", "500-800 ₽", "Профессиональный, авто"],
    ["Decorix", "Германия/РФ", "Средний", "300-450 ₽", "RAL + декоративные эффекты"],
    ["Aim-One", "Россия/Китай", "Эконом", "150-250 ₽", "Базовые цвета RAL"],
  ];

  const headerRow = new TableRow({
    tableHeader: true, cantSplit: true,
    children: headerCells.map(text => new TableCell({
      shading: { type: ShadingType.CLEAR, fill: c(P.primary) },
      margins: { top: 40, bottom: 40, left: 60, right: 60 },
      borders: { top: { style: BorderStyle.SINGLE, size: 3, color: c(P.primary) }, bottom: { style: BorderStyle.SINGLE, size: 3, color: c(P.primary) }, left: NB, right: NB },
      children: [new Paragraph({
        spacing: { line: 200 },
        children: [new TextRun({ text, bold: true, size: 14, color: c(P.white), font: FONT })],
      })],
    })),
  });

  const dataRows = rows.map((row, idx) => {
    const isEmpils = row[0].includes("NewTon");
    return new TableRow({
      cantSplit: true,
      children: row.map((text, ci) => new TableCell({
        shading: { type: ShadingType.CLEAR, fill: isEmpils ? c(P.accentLight) : (idx % 2 === 0 ? "FAFAFA" : c(P.white)) },
        margins: { top: 30, bottom: 30, left: 60, right: 60 },
        borders: { top: NB, bottom: { style: BorderStyle.SINGLE, size: 2, color: "E0E0E0" }, left: NB, right: NB },
        children: [new Paragraph({
          spacing: { line: 200 },
          children: [new TextRun({
            text, size: 14,
            bold: ci === 0 || isEmpils,
            color: c(isEmpils ? P.accent : P.body),
            font: FONT,
          })],
        })],
      })),
    });
  });

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 3, color: c(P.primary) },
      bottom: { style: BorderStyle.SINGLE, size: 3, color: c(P.primary) },
      left: NB, right: NB, insideHorizontal: NB, insideVertical: NB,
    },
    rows: [headerRow, ...dataRows],
  });
}

// === Таблица: рекомендации по цветам RAL ===
function buildColorsTable() {
  const headerCells = ["Приоритет", "RAL", "Цвет", "Применение", "Доля продаж (оценка)"];
  const rows = [
    ["1 — Базовые (старт)", "RAL 9005", "Чёрный", "Авто, заборы, металл, DIY", "~25% (топ-1)"],
    ["1 — Базовые (старт)", "RAL 9010", "Белый", "Интерьер, дерево, DIY", "~15%"],
    ["1 — Базовые (старт)", "RAL 8017", "Шоколадно-коричневый", "Дерево, заборы, фасад", "~12%"],
    ["1 — Базовые (старт)", "RAL 7024", "Графитовый серый", "Металл, промышленность, DIY", "~10%"],
    ["1 — Базовые (старт)", "RAL 3000", "Красный", "Авто, сигнальная окраска", "~8%"],
    ["2 — Расширенные", "RAL 6005", "Зелёный (мох)", "Заборы, кровля, фасад", "~6%"],
    ["2 — Расширенные", "RAL 5010", "Синий (гентский)", "Авто, металл, DIY", "~5%"],
    ["2 — Расширенные", "RAL 1015", "Светло-бежевый", "Интерьер, дерево", "~4%"],
    ["2 — Расширенные", "RAL 7001", "Серебристо-серый", "Авто, металл", "~4%"],
    ["2 — Расширенные", "RAL 1003", "Жёлтый (сигнальный)", "Сигнальная, спецтехника", "~3%"],
    ["3 — Опциональные", "RAL 4001", "Фиолетовый", "Творчество, граффити, DIY", "~2%"],
    ["3 — Опциональные", "RAL 6018", "Жёлто-зелёный", "Творчество, хобби", "~2%"],
    ["3 — Опциональные", "RAL 2000", "Жёлто-оранжевый", "Авто, творчество", "~2%"],
    ["3 — Опциональные", "RAL 9006", "Бело-алюминиевый", "Металл, промышленность", "~2%"],
  ];

  const headerRow = new TableRow({
    tableHeader: true, cantSplit: true,
    children: headerCells.map((text, ci) => new TableCell({
      shading: { type: ShadingType.CLEAR, fill: c(P.primary) },
      margins: { top: 20, bottom: 20, left: 40, right: 40 },
      borders: { top: { style: BorderStyle.SINGLE, size: 3, color: c(P.primary) }, bottom: { style: BorderStyle.SINGLE, size: 3, color: c(P.primary) }, left: NB, right: NB },
      children: [new Paragraph({
        spacing: { line: 160 },
        children: [new TextRun({ text, bold: true, size: 12, color: c(P.white), font: FONT })],
      })],
    })),
  });

  const dataRows = rows.map((row, idx) => {
    const priority = row[0];
    let priorityColor = P.accent;
    if (priority.startsWith("1")) priorityColor = P.accent;
    else if (priority.startsWith("2")) priorityColor = P.body;
    else priorityColor = P.secondary;

    return new TableRow({
      cantSplit: true,
      children: row.map((text, ci) => new TableCell({
        shading: { type: ShadingType.CLEAR, fill: idx % 2 === 0 ? "FAFAFA" : c(P.white) },
        margins: { top: 12, bottom: 12, left: 40, right: 40 },
        borders: { top: NB, bottom: { style: BorderStyle.SINGLE, size: 1, color: "E8E8E8" }, left: NB, right: NB },
        children: [new Paragraph({
          spacing: { line: 160 },
          children: [new TextRun({
            text, size: 12,
            bold: ci === 0 || ci === 1 || ci === 4,
            color: c(ci === 0 ? priorityColor : (ci === 4 ? P.accent : P.body)),
            font: FONT,
          })],
        })],
      })),
    });
  });

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 3, color: c(P.primary) },
      bottom: { style: BorderStyle.SINGLE, size: 3, color: c(P.primary) },
      left: NB, right: NB, insideHorizontal: NB, insideVertical: NB,
    },
    rows: [headerRow, ...dataRows],
  });
}

// === Сборка документа ===
const children = [];

// Заголовок
children.push(new Paragraph({
  spacing: { after: 40 },
  children: [new TextRun({ text: "АНАЛИЗ · ЭМПИЛС · 2026", size: 14, color: c(P.accent), bold: true, font: FONT, characterSpacing: 30 })],
}));

children.push(new Paragraph({
  spacing: { after: 80, line: 260 },
  children: [new TextRun({ text: "Выход на рынок аэрозольных ЛКМ: анализ и рекомендации", size: 24, bold: true, color: c(P.primary), font: FONT })],
}));

children.push(body("Компактный бриф: рынок РФ, конкуренты, линейки, маржинальность, сложности. Цель — определить приоритетные линейки для быстрого старта и максимальной маржи.", 14));

children.push(divider());

// 1. Рынок
children.push(h2("1. Рынок аэрозольных ЛКМ в России"));
children.push(bulletBold("Объём рынка РФ", "~8-10 млн баллонов/год, ~$70-100 млн (оценочно). Доля в общем рынке ЛКМ — ~2-3% от тоннажа, но выше по выручке из-за premium-цены."));
children.push(bulletBold("Динамика", "Стабильный рост 5-6%/год (мировой CAGR 5,5%). В РФ рынок растёт за счёт DIY-сегмента и автопокраски. После 2022 г. — уход части импорта (Krylon, Rust-Oleum) освободил нишу."));
children.push(bulletBold("Сегментация", "DIY/хозмаги — 45% · Авто/мото — 25% · Граффити/творчество — 15% · B2B/промышленность — 10% · Прочее — 5%"));
children.push(bulletBold("Тренды", "1) Рост акриловых спреев (экологичность, отсутствие запаха) · 2) Спрос на спецэффекты (металлик, хром, флуоресцентные) · 3) Локализация производства (уход импорта) · 4) Рост e-commerce (OZON, Wildberries)"));

// 2. Линейки
children.push(h2("2. Линейки: что производить, приоритеты, маржа"));
children.push(buildLineupTable());
children.push(new Paragraph({ spacing: { before: 40 } }));
children.push(bulletBold("ПЕРВАЯ ОЧЕРЕДЬ (быстрые продажи)", "Универсальная алкидная по RAL (глянец) + Грунт-эмаль алкидная (не по RAL). Причины: высокий спрос DIY + авто, узнаваемость RAL, высокая маржа (40-55%), синергия с брендом TITAN (грунт-эмаль)."));
children.push(bulletBold("ВТОРАЯ ОЧЕРЕДЬ (маржа + тренд)", "Акриловая универсальная по RAL (матовый/полуматовый). Причины: растущий тренд на экологичность, более высокая цена, но меньший объём. Целевая аудитория — творчество, граффити, премиум DIY."));
children.push(bulletBold("Глянец vs матовый", "Глянец — больший объём продаж (алкидная, авто, DIY). Матовый — растущий тренд (акрил, интерьер, дизайн). Рекомендация: глянец = объём, матовый = маржа."));
children.push(bulletBold("ПОЧЕМУ АЛКИД — ТОП-1 (а не акрил)", "Алкид занимает ~55% рынка аэрозолей (акрил ~30%, нитро ~10%). Алкид = «золотой стандарт» спреев: укрывистость за 1 слой, сохнет 15-30 мин, прочная плёнка, глянец, атмосферостойкость. Алкид дешевле по сырью (на 20-30%), значит выше маржа. Алкид подходит для металла + дерева + пластика (универсал), акрил — преимущественно интерьер. BOSNY и KUDO — лидеры рынка — оба на алкидной основе. У Эмпилс уже есть алкидная экспертиза (TITAN) — не нужно заново строить компетенцию."));
children.push(bulletBold("ПОЧЕМУ ГРУНТ-ЭМАЛЬ — 2-е МЕСТО (а не 1-е)", "Грунт-эмаль 3в1 = узкая ниша: только металл, только по ржавчине. Универсальная алкидная — шире (металл + дерево + пластик + камень). Маржа у грунт-эмали выше (45-55% vs 40-50%), но объём — меньше. Грунт-эмаль = самый быстрый старт (синергия с TITAN: бренд, формулировка, каналы), но потолок ниже — не все DIY-покупатели красят ржавчину. Поэтому: грунт-эмаль = быстрый старт + максимальная маржа, универсальная = массовый объём."));
children.push(bulletBold("ПОЧЕМУ АКРИЛ — 3-е МЕСТО (а не 1-е, ведь он растёт)", "Акрил — растущий тренд (CAGR 8% vs 4% у алкида), но пока уступает по объёму в 1.8 раза. В аэрозоле акрил дороже по сырью на 20-30% (дисперсия дороже алкидной смолы), значит ниже маржа (30-40% vs 40-55%). Акриловые спреи — преимущественно для интерьера и творчества, а основной объём рынка — это металл/улица/авто = алкид. Акрил не подходит для окраски по ржавчине (нет антикорр. свойств). Причина 3-го места: для быстрого старта нужен объём, а объём — это алкид. Акрил = тренд + маржа в премиуме, но не объём."));
children.push(bulletBold("ЧЕГО НЕ ПРОИЗВОДИТЬ (анти-рекомендации)", "Нитроцеллюлозные спреи — устаревшая технология, вред, низкая маржа. Грунт-эмаль акриловая по RAL — каннибализация грунт-эмали алкидной (TITAN). Премиум-спецэффекты (хром, флуоресцент) — высокая себестоимость, малый объём, ниша Krylon/Motip. Термостойкие спреи (до 600°C) — узкий рынок, спец. сырьё, конкуренция с Bosny/Certa. Авто-эмали полные (не RAL) — фрагментация по маркам авто, сложность колеровки."));
children.push(bulletBold("В КАКИЕ СЕГМЕНТЫ НЕ ИДТИ", "Эконом (до 250 ₽) — демпинг Aim-One/Китай, нет маржи без масштаба. Граффити-профи — узкая ниша, высокая конкуренция (Montana, Loop). Авто-профессиональный (СТО) — требуют каталога по маркам авто, сертификатов, цикл сделки 6-12 мес. Промышленные спреи (ISO) — низкий объём, длинный цикл, B2B-специфика."));
children.push(bulletBold("РЕКОМЕНДАЦИИ ПО ЦВЕТАМ RAL", "Топ-5 цветов дают ~70% продаж. Старт — 5 базовых, расширение — +5 расширенных, опционально — +4 нишевых. Синергия с TITAN: RAL 9005, 8017, 7024 — топ-цвета грунт-эмали TITAN, переносятся в аэрозоль."));
children.push(buildColorsTable());

// 3. Конкуренты
children.push(h2("3. Конкуренты на рынке РФ"));
children.push(buildCompetitorsTable());
children.push(new Paragraph({ spacing: { before: 40 } }));
children.push(bulletBold("Позиция Эмпилс (NewTon)", "Сейчас — ограниченный ассортимент, эконом-сегмент. Возможность: расширение до 30-50 SKU, вход в средний сегмент через RAL + синергию с TITAN."));

// 4. Маржинальность
children.push(h2("4. Маржинальность и экономика"));
children.push(bulletBold("Себестоимость баллона 400-520 мл", "80-150 ₽ (сырьё 40-60% + баллон 15-20% + газ-пропеллент 5-10% + фасовка/логистика 15-20%)"));
children.push(bulletBold("Розничная цена", "Эконом 150-250 ₽ · Средний 300-500 ₽ · Премиум 600-900 ₽"));
children.push(bulletBold("Маржа производителя", "30-55% в зависимости от сегмента. Выше всего — грунт-эмаль алкидная (45-55%) за счёт TITAN-синергии. Ниже всего — акриловая эконом (25-35%)."));
children.push(bulletBold("ROI входа", "При контрактном производстве (без своей линии) — инвестиции 2-5 млн ₽ на запуск 15-20 SKU. Окупаемость 8-12 мес при марже 40%+."));

// 5. Сложности
children.push(h2("5. Сложности и риски"));
children.push(bulletBold("Газ-пропеллент (LPG/диметиловый эфир)", "Воспламеняемость, специфика хранения и транспортировки. Нужны сертифицированные помещения. Давление зависит от температуры — проблемы зимой."));
children.push(bulletBold("Баллоны и комплектующие", "Производство жести в РФ: ЗАО «Аэрозольная банка» (СПБ), ООО «Линар» (Ставрополь). Клапаны и кнопки — в основном импорт (Китай). Риск перебоев поставок."));
children.push(bulletBold("Сертификация", "ГОСТ Р, СГР, пожарная безопасность (для LPG). Срок получения — 2-4 месяца на SKU."));
children.push(bulletBold("Конкуренция по цене", "Aim-One и китайские бренды демпингуют до 150 ₽. Сложно конкурировать в эконом-сегменте без масштаба."));
children.push(bulletBold("Каналы сбыта", "DIY-сети требуют входную плату + ретро-бонусы. E-commerce (OZON, WB) — высокая комиссия 15-25%."));

children.push(divider());

// 6. Вывод и рекомендации
children.push(h2("6. Рекомендации: с чего начать"));
children.push(bulletBold("ШАГ 1", "Запустить 2 линейки ПЕРВОЙ очереди: универсальная алкидная RAL (10-15 цветов) + грунт-эмаль алкидная (3-5 цветов). Фасовка 400 и 520 мл. Через контрактное производство — без своей линии."));
children.push(bulletBold("ШАГ 2", "Добавить 5-7 SKU акриловой по RAL (матовый) — для трендового сегмента. Цвета — по таблице выше: приоритет 1 (5 базовых) + приоритет 2 (2-3 расширенных)."));
children.push(bulletBold("ШАГ 3", "Каналы: DIY (Лемана, Петрович, Максидом) + e-commerce (OZON, WB) + авто-сегмент (через дистрибьюторов TITAN). Бренд — расширение TITAN или NewTon?"));
children.push(bulletBold("БЫСТРЫЕ ПРОДАЖИ", "Грунт-эмаль алкидная (синергия с TITAN) — самый быстрый старт. Универсальная алкидная RAL глянец — самый большой объём DIY."));
children.push(bulletBold("МАКСИМАЛЬНАЯ МАРЖА", "Грунт-эмаль алкидная (45-55%) + акриловая премиум по RAL (35-45%)."));

children.push(new Paragraph({
  spacing: { before: 80 },
  alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "ЗАО «Эмпилс» · Аналитический отдел · 2026", size: 14, italics: true, color: c(P.secondary), font: FONT })],
}));

// === Документ ===
const doc = new Document({
  creator: "ЗАО Эмпилс",
  title: "Анализ выхода на рынок аэрозольных ЛКМ",
  styles: {
    default: {
      document: {
        run: { font: FONT, size: 16, color: c(P.body) },
        paragraph: { spacing: { line: 240 } },
      },
    },
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 500, bottom: 400, left: 600, right: 500 },
      },
    },
    children,
  }],
});

const OUT = "/home/z/my-project/download/Эмпилс_Анализ_аэрозоли_РФ.docx";

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log("✅ Документ сохранён:", OUT);
  console.log("   Размер:", buf.length.toLocaleString(), "байт");
}).catch(err => {
  console.error("❌", err);
  process.exit(1);
});
