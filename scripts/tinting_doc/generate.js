// Справка: колеровочное оборудование в магазинах ЛКМ — анализ конкурентов
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, ShadingType, WidthType,
  PageBreak, ExternalHyperlink,
} = require("docx");
const fs = require("fs");

const P = {
  primary: "#1F4E79",
  accent: "#C00000",
  body: "#1A1A1A",
  secondary: "#595959",
  muted: "#808080",
  positive: "#375623",
  warn: "#BF8F00",
};
const c = (hex) => hex.replace("#", "");

const FONT = { ascii: "Calibri", eastAsia: "Calibri" };

// === Утилиты ===
function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 180, line: 312 },
    children: [new TextRun({ text, bold: true, size: 32, color: c(P.primary) })],
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 140, line: 312 },
    children: [new TextRun({ text, bold: true, size: 28, color: c(P.primary) })],
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 220, after: 120, line: 312 },
    children: [new TextRun({ text, bold: true, size: 24, color: c(P.primary) })],
  });
}

function p(text, opts = {}) {
  const runs = Array.isArray(text) ? text : [new TextRun({
    text, size: 22, color: c(P.body), font: FONT,
  })];
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 312, after: 140 },
    indent: { firstLine: 420 },
    ...opts,
    children: runs,
  });
}

function pRich(parts, opts = {}) {
  const runs = parts.map(part => {
    if (typeof part === "string") {
      return new TextRun({ text: part, size: 22, color: c(P.body), font: FONT });
    }
    const [t, props = {}] = Array.isArray(part) ? part : [part.text, part];
    return new TextRun({
      text: t, size: props.size || 22,
      color: c(props.color || P.body),
      bold: !!props.bold,
      italics: !!props.italic,
      font: FONT,
    });
  });
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 312, after: 140 },
    indent: { firstLine: 420 },
    ...opts,
    children: runs,
  });
}

function li(text, opts = {}) {
  const runs = Array.isArray(text) ? text : [new TextRun({
    text, size: 22, color: c(P.body), font: FONT,
  })];
  return new Paragraph({
    spacing: { line: 312, after: 80 },
    bullet: { level: 0 },
    ...opts,
    children: runs,
  });
}

function liBold(label, text) {
  return li([
    new TextRun({ text: label, bold: true, size: 22, color: c(P.primary), font: FONT }),
    new TextRun({ text: " — ", size: 22, color: c(P.body), font: FONT }),
    new TextRun({ text: text, size: 22, color: c(P.body), font: FONT }),
  ]);
}

function empty() {
  return new Paragraph({
    spacing: { line: 240, after: 0 },
    children: [new TextRun({ text: "" })],
  });
}

function hr() {
  return new Paragraph({
    spacing: { before: 120, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: c(P.primary) } },
    children: [new TextRun({ text: "", size: 2 })],
  });
}

function caption(text) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 80, after: 200, line: 280 },
    children: [new TextRun({
      text, italics: true, size: 18, color: c(P.secondary), font: FONT,
    })],
  });
}

// Таблицы
function tCellHeader(text, widthPct) {
  return new TableCell({
    width: { size: widthPct, type: WidthType.PERCENTAGE },
    shading: { type: ShadingType.CLEAR, fill: c(P.primary) },
    margins: { top: 100, bottom: 100, left: 140, right: 140 },
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { line: 280 },
      children: [new TextRun({
        text, bold: true, size: 22, color: "FFFFFF", font: FONT,
      })],
    })],
  });
}

function tCell(text, widthPct, opts = {}) {
  const isBold = opts.bold || false;
  const color = opts.color ? c(opts.color) : c(P.body);
  const align = opts.align || AlignmentType.LEFT;
  const fill = opts.fill;
  const fontSize = opts.size || 22;

  const cellProps = {
    width: { size: widthPct, type: WidthType.PERCENTAGE },
    margins: { top: 80, bottom: 80, left: 140, right: 140 },
    children: [new Paragraph({
      alignment: align,
      spacing: { line: 280 },
      children: [new TextRun({
        text: String(text || '—'), bold: isBold, size: fontSize, color, font: FONT,
      })],
    })],
  };
  if (fill) {
    cellProps.shading = { type: ShadingType.CLEAR, fill: c(fill) };
  }
  return new TableCell(cellProps);
}

function tCellMulti(lines, widthPct, opts = {}) {
  const align = opts.align || AlignmentType.LEFT;
  const fill = opts.fill;
  const fontSize = opts.size || 22;
  const color = opts.color ? c(opts.color) : c(P.body);
  const isBold = opts.bold || false;

  const paras = lines.map(line => new Paragraph({
    alignment: align,
    spacing: { line: 280, after: 40 },
    children: [new TextRun({
      text: String(line || '—'), bold: isBold, size: fontSize, color, font: FONT,
    })],
  }));

  const cellProps = {
    width: { size: widthPct, type: WidthType.PERCENTAGE },
    margins: { top: 80, bottom: 80, left: 140, right: 140 },
    children: paras,
  };
  if (fill) {
    cellProps.shading = { type: ShadingType.CLEAR, fill: c(fill) };
  }
  return new TableCell(cellProps);
}

const TB = {
  top: { style: BorderStyle.SINGLE, size: 4, color: c(P.primary) },
  bottom: { style: BorderStyle.SINGLE, size: 4, color: c(P.primary) },
  left: { style: BorderStyle.SINGLE, size: 2, color: "BFBFBF" },
  right: { style: BorderStyle.SINGLE, size: 2, color: "BFBFBF" },
  insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "D0D0D0" },
  insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "D0D0D0" },
};

// =================== СОДЕРЖАНИЕ ===================
const children = [];

// === ВВЕДЕНИЕ ===
children.push(new Paragraph({
  alignment: AlignmentType.LEFT,
  spacing: { before: 0, after: 180, line: 280 },
  children: [new TextRun({
    text: "Анализ конкурентов по колеровочному оборудованию в рознице ЛКМ",
    italics: true, size: 20, color: c(P.secondary),
  })],
}));

children.push(new Paragraph({
  alignment: AlignmentType.LEFT,
  spacing: { before: 0, after: 240, line: 280 },
  children: [new TextRun({
    text: "Цель: собрать информацию о том, как производители ЛКМ и производители колеровочного оборудования рекомендуют размещать и оформлять зону колеровки в розничном магазине.",
    italics: true, size: 20, color: c(P.secondary),
  })],
}));

children.push(hr());

// === КРАТКАЯ СПРАВКА ===
children.push(h1("1. Краткая справка: что нашли"));

children.push(p([
  "Исследование проведено в два этапа: ",
  ["Этап 1", { bold: true, color: P.primary }],
  " — анализ сайтов производителей ЛКМ (Tikkurila, Dulux, AkzoNobel, Benjamin Moore, Sherwin-Williams) на предмет информации о колеровочном оборудовании для розницы; ",
  ["Этап 2", { bold: true, color: P.primary }],
  " — анализ сайтов производителей колеровочного оборудования (Corob, Fast&Fluid, Hero, Dromont, CPS Color, Santint, VPS Color, Tintec, Omnitint, Tecsa, Palizh).",
]));

children.push(p([
  "Главный вывод: ",
  ["специализированных буклетов с детальными планограммами зоны колеровки в открытом доступе нет ни у одного производителя", { bold: true, color: P.accent }],
  ". Производители дают общие рекомендации по типам оборудования и комплекту POS-материалов, но визуальное оформление «колеровочного острова» каждый ритейлер проектирует самостоятельно или по индивидуальному заказу у производителя.",
]));

children.push(p([
  "Лучший публичный материал по теме — ",
  ["PDF-каталог «Тиккурила Симфония»", { bold: true, color: P.accent }],
  " (complexconstruction.ru/upload/iblock/793/Tikkurila_Catalog.pdf), где есть разделы «Колеровочное оборудование», «Маркетинговые материалы» и «Демонстрационные стенды» с конкретными рекомендациями по оснащению торговых точек разного типа. На основе этого каталога и собрана существенная часть рекомендаций ниже.",
]));

children.push(p([
  "Российский рынок также предлагает универсальное решение Omnitint (omnitint.ru) — ",
  ["независимую колеровочную систему", { bold: true, color: P.primary }],
  ", которая позволяет магазину работать с разными брендами красок на одной машине. Это снимает привязку к одному производителю ЛКМ и упрощает организацию зоны колеровки.",
]));

// === ЭТАП 1: ПРОИЗВОДИТЕЛИ ЛКМ ===
children.push(h1("2. Этап 1: Производители ЛКМ и их подходы"));

children.push(p(
  "Крупные мировые производители ЛКМ строят свои колеровочные системы по принципу замкнутого цикла: «базовая краска + колорант + дозатор + ПО + каталог цветов — всё от одного производителя». Это даёт точность попадания в цвет, но привязывает магазин к одному бренду. Ниже — что каждый из них публикует о колеровочном оборудовании."
));

children.push(h2("2.1. Tikkurila — система «Симфония»"));

children.push(p([
  "Tikkurila — ",
  ["пионер технологии колеровки", { bold: true }],
  " с 1970-х годов. Система «Тиккурила Симфония» включает 6 взаимоувязанных компонентов: базисные краски, пигментные пасты, рецептуры и ПО, дозирующее оборудование, смесители и маркетинговые материалы. База — 2 436 фирменных оттенков + 15 000 цветов в базе данных программы Colour Composer."
]));

children.push(pRich([
  ["Где искать информацию:", { bold: true, color: P.primary }],
  [" PDF-каталог «Тиккурила Симфония»", { italic: true }],
  [" на сайте complexconstruction.ru. Это самый подробный из публично доступных материалов по теме.",
   {}],
]));

children.push(p("Tikkurila даёт рекомендации по типу торговой точки — три уровня оборудования:"));

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TB,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tCellHeader("Тип магазина", 30),
      tCellHeader("Рекомендуемое оборудование", 40),
      tCellHeader("Пример модели Tikkurila", 30),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Маленький магазин «сделай сам» (DIY)", 30, { bold: true, fill: "F2F7FC" }),
      tCell("Компактный ручной дозатор на 16 канистр по 3 л. Минимальная доза 0,077 мл.", 40, { fill: "F2F7FC" }),
      tCell("Tikkurila Rondo", 30, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Магазин средних размеров", 30, { bold: true }),
      tCell("Автоматический дозатор на 16-20 канистр, вместимость канистр 3-5 л. ПО Colour Composer.", 40),
      tCell("Tikkurila Hydro", 30),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Крупный оптовый склад / DIY-гипермаркет", 30, { bold: true, fill: "F2F7FC" }),
      tCell("Высокопроизводительный автоматический дозатор с расширенным набором канистр, спектрофотометр для измерения цвета по образцу.", 40, { fill: "F2F7FC" }),
      tCell("Tikkurila OptiColor + спектрофотометр", 30, { fill: "F2F7FC" }),
    ]}),
  ],
}));

children.push(caption("Источник: PDF-каталог «Тиккурила Симфония», раздел «Колеровочное оборудование»."));

children.push(h3("Маркетинговые материалы Tikkurila"));

children.push(p("Tikkurila предоставляет дилерам полный набор POS-материалов, формирующих фирменный стиль зоны колеровки:"));

children.push(liBold("Стенд «Grande Bello»", "флагманский демонстрационный стенд: 4 элемента-шкафа из ламината алюминиевого цвета, шкаф-элемент «Тайка», тыльный освещённый белый парус с логотипом Tikkurila, световой шкаф с 3 типами освещения (дневной свет, люминесцентная, накаливания — для оценки цвета при разном свете)."));
children.push(liBold("Выкраски и чипсы", "образцы цветов колерной системы «Симфония» для демонстрации."));
children.push(liBold("Модуль «Красивый Дом»", "карманы с листами для интерьеров и фасадов — 60 + 40 листов подобранных цветовых комбинаций."));
children.push(liBold("Образцы накраски Тайка", "А4, перламутровые краски и лазури."));
children.push(liBold("Передвижные каретки", "с проспектами и колерными картами."));
children.push(liBold("Веера цветов", "Опус 1 (яркие), Опус 2 (пастельные), «Симфония Интерьер» (720 цветов — основной инструмент продавца-консультанта), «Красивый Дом» для интерьеров и фасадов."));
children.push(liBold("Колерные карты", "буклеты для лаков и антисептиков."));

children.push(p([
  "Ключевой подход Tikkurila к оформлению зоны колеровки: ",
  ["серый цвет алюминия как нейтральный фон", { bold: true, color: P.primary }],
  ", который подчёркивает яркость выкрасок и скрывает пыль/грязь. Белый парус с логотипом создаёт контраст. Каждый стенд имеет подсветку — это критично, потому что покупатель выбирает цвет при разном освещении.",
]));

children.push(h2("2.2. Dulux (AkzoNobel)"));

children.push(p(
  "Dulux работает по модели «Color Centre» — фирменная колеровочная зона в магазине. Российский сайт dulux.ru информации по оборудованию не раскрывает; основная информация доступна на сайтах дилеров Dulux Color (dulux-color.ru)."
));

children.push(p([
  "В России Dulux использует колеровочную систему ",
  ["AkzoNobel Mixit", { bold: true }],
  " с автоматическими дозаторами. В мире AkzoNobel работает с оборудованием Fast&Fluid Management и Corob — двух крупнейших производителей дозаторов для розницы."
]));

children.push(liBold("Фирменная концепция", "Dulux Decorator Centre (в Великобритании) — это полноценные шоу-румы с колеровочной зоной, обучающими площадками и демонстрационными стендами. В России аналог — точки Dulux Color в Москве."));
children.push(liBold("Что Dulux предоставляет дилерам", "колеровочное оборудование (по договору), цветовые веера Dulux Colour Palette, RAL, NCS, POS-материалы с брендингом Dulux."));

children.push(h2("2.3. Другие производители ЛКМ"));

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TB,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tCellHeader("Производитель", 25),
      tCellHeader("Колеровочная система", 25),
      tCellHeader("Что предоставляет дилерам", 35),
      tCellHeader("Информация в открытом доступе", 15),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Beckers", 25, { bold: true, fill: "F2F7FC" }),
      tCell("Beckers Color Studio", 25, { fill: "F2F7FC" }),
      tCell("Фирменное оборудование + POS-материалы. Адресная программа для профессиональных дилеров.", 35, { fill: "F2F7FC" }),
      tCell("минимум", 15, { align: AlignmentType.CENTER, fill: "F2F7FC", color: P.warn }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Hempel", 25, { bold: true }),
      tCell("Hempacolor", 25),
      tCell("Колеровочное оборудование для дилеров морских и защитных ЛКМ.", 35),
      tCell("минимум", 15, { align: AlignmentType.CENTER, color: P.warn }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Benjamin Moore (США)", 25, { bold: true, fill: "F2F7FC" }),
      tCell("Benjamin Moore Color Center", 25, { fill: "F2F7FC" }),
      tCell("Фирменные стенды, оборудование Gennex Color System, обучение персонала. Сильная программа в США.", 35, { fill: "F2F7FC" }),
      tCell("средне", 15, { align: AlignmentType.CENTER, fill: "F2F7FC", color: P.warn }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Sherwin-Williams", 25, { bold: true }),
      tCell("Sher-Color + ColorToGo", 25),
      tCell("Полноценная концепция «розничного центра красок» с фирменными стендами, образцами выкрасок и услугой ColorToGo (банки-пробники, покупатель берёт домой попробовать).", 35),
      tCell("средне", 15, { align: AlignmentType.CENTER, color: P.warn }),
    ]}),
  ],
}));

children.push(caption("Источники: dulux.ru, dulux-color.ru, tikkurila.kz, complexconstruction.ru (PDF), benjaminmoore.com, sherwin-williams.com."));

children.push(p([
  "Сводный вывод по Этапу 1: ",
  ["производители ЛКМ держат информацию о колеровочном оборудовании закрытой", { bold: true, color: P.accent }],
  " — это часть их коммерческой модели привязки дилера. Полноценные буклеты с планограммами зоны колеровки не публикуются. Исключение — Tikkurila, чей PDF-каталог «Симфония» является самым подробным публичным материалом по теме."
]));

// === ЭТАП 2: ПРОИЗВОДИТЕЛИ КОЛЕРОВОЧНОГО ОБОРУДОВАНИЯ ===
children.push(h1("3. Этап 2: Производители колеровочного оборудования"));

children.push(p(
  "В отличие от производителей ЛКМ, производители колеровочного оборудования более открыты: их цель — продать оборудование любому дилеру ЛКМ или DIY-сети, поэтому они публикуют каталоги, технические спецификации и инструкции. Однако специализированных буклетов «как оформить зону колеровки в магазине» нет — производители продают отдельные машины и ПО, а не дизайн-проект зоны."
));

children.push(h2("3.1. Мировые производители"));

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TB,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tCellHeader("Производитель", 18),
      tCellHeader("Страна", 12),
      tCellHeader("Ключевые продукты", 30),
      tCellHeader("Доступность в РФ", 18),
      tCellHeader("Информация в открытом доступе", 22),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Corob Group", 18, { bold: true, fill: "F2F7FC" }),
      tCell("Италия", 12, { fill: "F2F7FC" }),
      tCell("Дозаторы, смесители, ПО. Один из мировых лидеров.", 30, { fill: "F2F7FC" }),
      tCell("Через дилеров (kolerovka.com)", 18, { fill: "F2F7FC" }),
      tCell("Каталоги продуктов, спецификации. Буклетов по оформлению зоны нет.", 22, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Fast & Fluid Management", 18, { bold: true }),
      tCell("Нидерланды", 12),
      tCell("Дозаторы TI Series (TI Element 2S, TI Lite), ПО ColorTone. Партнёр AkzoNobel/Dulux.", 30),
      tCell("✗ С 2022 не поставляется в РФ (через Tecsa — в другие страны СНГ)", 18, { color: P.accent }),
      tCell("fast-fluid.com — каталог продуктов, спецификации.", 22),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Hero SRL", 18, { bold: true, fill: "F2F7FC" }),
      tCell("Италия", 12, { fill: "F2F7FC" }),
      tCell("Дозаторы и миксеры для ЛКМ, компактные решения для розницы.", 30, { fill: "F2F7FC" }),
      tCell("Через lkm-kolerovka.ru", 18, { fill: "F2F7FC" }),
      tCell("Спецификации на оборудование.", 22, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Dromont", 18, { bold: true }),
      tCell("Италия", 12),
      tCell("Retail Tinting Machines (Perfecta), Factory Dispensing Systems, ПО IBSM. Полный цикл от розницы до производства.", 30),
      tCell("Через дилеров", 18),
      tCell("dromont.com — раздел «Retail Tinting Machines» с описанием решений.", 22),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("CPS Color", 18, { bold: true, fill: "F2F7FC" }),
      tCell("Швеция/США", 12, { fill: "F2F7FC" }),
      tCell("Дозаторы и краскораспределительные системы. Исторический лидер рынка.", 30, { fill: "F2F7FC" }),
      tCell("Через дилеров", 18, { fill: "F2F7FC" }),
      tCell("cpscolor.com — сайт недоступен из РФ без VPN.", 22, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Santint", 18, { bold: true }),
      tCell("Китай", 12),
      tCell("Дозаторы (модели AO480, AO320), шейкеры, смесительные установки. Самый доступный сегмент.", 30),
      tCell("ООО «Тинтэко» — santint-rus.ru", 18),
      tCell("Полные спецификации на русском.", 22),
    ]}),
  ],
}));

children.push(caption("Источники: официальные сайты производителей, kolerovka.com, kolerovka.ru, lkm-kolerovka.ru, santint-rus.ru, tecsa.ru, tintec.ru, omnitint.ru."));

children.push(h2("3.2. Российские игроки и интеграторы"));

children.push(p([
  "На российском рынке после 2022 года сформировалась новая ситуация: ",
  ["Fast&Fluid (Нидерланды) и Corob (Италия) напрямую не поставляются", { bold: true, color: P.accent }],
  ", но доступны через официальных дилеров (Tecsa — для Fast&Fluid в страны СНГ). Активно развиваются китайский Santint и отечественные интеграторы (Omnitint, Tintec, VPS Color, Palizh)."
]));

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TB,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tCellHeader("Компания", 18),
      tCellHeader("Что предлагает", 38),
      tCellHeader("Что рассказывает про организацию зоны колеровки", 44),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Omnitint", 18, { bold: true, color: P.primary, fill: "F2F7FC" }),
      tCellMulti([
        "Универсальная колеровочная система:",
        "• 16 универсальных пигментных паст",
        "• ПО tinTECH Studio",
        "• Дозаторы (любые бренды)",
        "• Лаборатория адаптации формул",
        "• 1 год сервисного обслуживания бесплатно",
      ], 38, { fill: "F2F7FC" }),
      tCellMulti([
        "Самая подробная информация в РФ. Ключевая идея:",
        "• одна машина на весь ассортимент магазина (а не по машине на бренд)",
        "• освобождается торговая площадь",
        "• сокращение товарного запаса колорантов",
        "• свобода смены брендов без замены системы",
        "• полный запуск: 3-6 недель (включая обучение)",
      ], 44, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Tecsa", 18, { bold: true }),
      tCellMulti([
        "Официальный дилер Fast&Fluid Management 25 лет.",
        "• Автоматические дозаторы",
        "• Шейкеры",
        "• ПО",
        "• Сервис во всех регионах РФ",
      ], 38),
      tCellMulti([
        "Дилерская модель: продают оборудование + обслуживание.",
        "Важно: с 2022 поставки Fast&Fluid в РФ и РБ прекращены,",
        "но возможны в другие страны СНГ по прямому контракту.",
        "Буклетов по оформлению зоны нет.",
      ], 44),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Tintec", 18, { bold: true, fill: "F2F7FC" }),
      tCellMulti([
        "Подбор системы под задачу:",
        "• для магазинов",
        "• для колеровочных центров",
        "• для производственных задач",
      ], 38, { fill: "F2F7FC" }),
      tCellMulti([
        "Подход «система под вашу задачу» — консультация по выбору.",
        "Специализированных буклетов по оформлению зоны нет.",
      ], 44, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Palizh (Палиж)", 18, { bold: true }),
      tCellMulti([
        "Комплекс колеровочного оборудования:",
        "• Дозаторы, миксеры, шейкеры",
        "• Доставка и установка",
        "• Онлайн-консультации",
        "• Электронная база рецептур",
      ], 38),
      tCellMulti([
        "Указывают типы оборудования и форму тары (круглая/квадратная/овальная).",
        "Рекомендуют шейкеры для лабораторий, оптовых и розничных магазинов, складов, цехов.",
        "Планограмм зоны нет.",
      ], 44),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("VPS Color", 18, { bold: true, fill: "F2F7FC" }),
      tCellMulti([
        "Колеровочные системы:",
        "• Колоранты (универсальные, водные, органорастворимые)",
        "• Дозаторы, миксеры",
        "• Колориметры",
        "• Цветовые веера",
      ], 38, { fill: "F2F7FC" }),
      tCell("Позиционируют как «производительные и современные дозаторы», «простые и надёжные миксеры». Без деталей оформления зоны.", 44, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Plasters.ru", 18, { bold: true }),
      tCellMulti([
        "Дилер колеровочного оборудования:",
        "• Доставка на точку продаж",
        "• Подробная инструкция по сборке и запуску",
        "• Техническая поддержка",
      ], 38),
      tCellMulti([
        "Интегратор, который закрывает вопрос «под ключ».",
        "Сами планограммы не публикуют.",
      ], 44),
    ]}),
  ],
}));

children.push(caption("Источники: omnitint.ru, tecsa.ru, tintec.ru, palizh.com, palizh.info, vpscolor.ru, plasters.ru."));

// === РАЗДЕЛ 4: ПРАКТИЧЕСКИЕ РЕКОМЕНДАЦИИ ===
children.push(h1("4. Сводные рекомендации: как оформить зону колеровки"));

children.push(p([
  "Поскольку специализированных буклетов по оформлению зоны колеровки в открытом доступе нет, ниже — ",
  ["свод рекомендаций, собранных из разных источников", { bold: true, color: P.primary }],
  " (прежде всего из PDF-каталога Tikkurila «Симфония» и рекомендаций Omnitint):"
]));

children.push(h2("4.1. Где ставить колеровочное оборудование в магазине"));

children.push(p(
  "Зависит от типа магазина и его площади. Tikkurila классифицирует 3 типа торговых точек:"
));

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TB,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tCellHeader("Тип магазина", 25),
      tCellHeader("Рекомендуемое место зоны колеровки", 40),
      tCellHeader("Оборудование", 35),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Маленький магазин DIY", 25, { bold: true, fill: "F2F7FC" }),
      tCell("У дальней стены, в углу. Компактная стойка 1×1 м. Зона видна с входа, но не загромождает проход.", 40, { fill: "F2F7FC" }),
      tCell("Ручной дозатор на 16 канистр (как Tikkurila Rondo). Миксер для одной банки.", 35, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Магазин средних размеров", 25, { bold: true }),
      tCell("Отдельный «колеровочный остров» 2×2 м в центре зала или вдоль стены. Зона выделяется светом и фирменными стендами.", 40),
      tCell("Автоматический дозатор, шейкер, компьютер с ПО, стенд с выкрасками, веера цветов.", 35),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Крупный DIY-гипермаркет", 25, { bold: true, fill: "F2F7FC" }),
      tCell("Отдельная зона 4×4 м с несколькими машинами (если работают разные бренды) или одна универсальная (Omnitint). Спектрофотометр для подбора цвета по образцу клиента.", 40, { fill: "F2F7FC" }),
      tCell("Полный комплект: дозатор + спектрофотометр + шейкер + стенд Grande Bello + ноутбук + принтер этикеток.", 35, { fill: "F2F7FC" }),
    ]}),
  ],
}));

children.push(h2("4.2. Что должно входить в зону колеровки"));

children.push(p(
  "По сводным данным Tikkurila, Omnitint и других производителей, минимальный комплект зоны колеровки:"
));

children.push(liBold("Дозирующее устройство (дозатор)", "автоматическое или ручное. Дозирует колоранты в базовую краску по формуле. Минимальная доза — 0,077 мл (как Tikkurila Rondo)."));
children.push(liBold("Смеситель (шейкер)", "перемешивает краску с колорантами до однородной массы. Типы: вибрационный, гиромиксер. Форма тары: круглая, квадратная, овальная — важна для выбора шейкера."));
children.push(liBold("Компьютер с ПО", "управляет дозатором, содержит базу рецептур (15 000+ цветов), ведёт статистику колеровок, печатает этикетки. Примеры: Tikkurila Colour Composer, Omnitint tinTECH Studio."));
children.push(liBold("Каталог цветов и веера", "для выбора цвета покупателем. Форматы: веера (для продавца), колерные карты (для покупателя), А4-стенды с готовыми выкрасками."));
children.push(liBold("Демонстрационный стенд", "с образцами выкрасок и колерной картой. У Tikkurila — стенд «Grande Bello» из 4 элементов алюминиевого цвета."));
children.push(liBold("Принтер этикеток", "печатает наклейки с названием краски, кодом цвета, объёмом тары, ценой, штрих-кодом."));
children.push(liBold("Спектрофотометр (опционально)", "для измерения цвета по образцу клиента и подбора собственной рецептуры. Необходим в крупных DIY-гипермаркетах."));

children.push(h2("4.3. Дизайн зоны колеровки — ключевые принципы"));

children.push(p("Из анализа Tikkurila (каталог «Симфония», раздел «Маркетинговые материалы»):"));

children.push(liBold("Серый цвет алюминия как фон", "нейтральный фон подчёркивает яркость выкрасок и скрывает пыль/грязь. Это базовый принцип оформления стенда колеровки."));
children.push(liBold("Белый парус с логотипом", "тыльная сторона стенда — белый, с фирменным логотипом. Создаёт контраст и фокусирует внимание."));
children.push(liBold("Подсветка с 3 типами света", "критично для выбора цвета: лампа дневного света, люминесцентная, лампа накаливания. Покупатель должен видеть, как цвет будет выглядеть при разном освещении дома."));
children.push(liBold("Шкафчики для хранения", "в выпуклой нижней части стенда — двухдверные шкафчики. Там хранят проспекты, колерные карты, расходные материалы."));
children.push(liBold("Передвижные каретки", "мобильные стойки с проспектами и колерными картами. Можно поднести к покупателю, чтобы он выбрал цвет."));
children.push(liBold("Башня «Тайка»", "у Tikkurila — отдельный модуль с образцами накрасок перламутровых красок и лазурей А4-формата."));

children.push(h2("4.4. Универсальное решение Omnitint"));

children.push(p([
  "Omnitint (omnitint.ru) — российская универсальная колеровочная система, ",
  ["решающая ключевую проблему розницы", { bold: true, color: P.accent }],
  ": когда магазин вынужден держать несколько колеровочных машин под разные бренды красок и дублирующие комплекты колорантов на складе."
]));

children.push(p("Omnitint предлагает:"));
children.push(liBold("Одна машина на весь ассортимент", "колеруйте все марки красок, представленные в магазине, на одном дозаторе. Освобождается торговая площадь."));
children.push(liBold("Один комплект колорантов", "16 универсальных пигментных паст вместо нескольких полных наборов под каждый бренд."));
children.push(liBold("Свобода смены брендов", "ротируйте хитовые позиции разных производителей без замены системы колеровки."));
children.push(liBold("Маркетинговые материалы в ПО", "баннеры, иконки, описания продуктов под ваш магазин."));
children.push(liBold("Сервисное обслуживание 1 год бесплатно", "установка, калибровка, ремонт оборудования всех производителей."));
children.push(liBold("Сроки запуска", "лабораторная адаптация под ассортимент — до 1 месяца. Полный запуск — 3-6 недель, включая поставку, обучение, настройку."));

// === РАЗДЕЛ 5: ИСТОЧНИКИ ===
children.push(h1("5. Источники"));

children.push(h2("5.1. Производители ЛКМ"));
children.push(li("Tikkurila — PDF-каталог «Тиккурила Симфония»: https://www.complexconstruction.ru/upload/iblock/793/Tikkurila_Catalog.pdf"));
children.push(li("Tikkurila — инструкция по применению Symphony: https://tikkurila.kz/dlya-professionalov/vdohnovenie/tikkurila-symphony-instrukciya-po-primeneniyu"));
children.push(li("Dulux Россия: https://www.dulux.ru/"));
children.push(li("Dulux Color: https://dulux-color.ru/"));
children.push(li("Wagner — Tikkurila: https://wagner.ru/professionalnoe/brands/tikkurila.html"));
children.push(li("Gelingen — Tikkurila: https://gelingen.biz/promyshlennoe-oborudovanie/tikkurila/"));
children.push(li("ManualsPlus — Tikkurila руководства: https://manuals.plus/category/tikkurila"));
children.push(li("Dulux колеровка по каталогу: https://www.kraski-online.ru/tinting/dulux_colour_pallete/"));

children.push(h2("5.2. Производители колеровочного оборудования (мировые)"));
children.push(li("Fast & Fluid Management (Нидерланды): https://www.fast-fluid.com/"));
children.push(li("Dromont (Италия): https://www.dromont.com/"));
children.push(li("Corob Group (Италия) — через дилеров"));
children.push(li("Hero SRL (Италия) — через lkm-kolerovka.ru"));
children.push(li("CPS Color (Швеция/США) — сайт недоступен из РФ без VPN"));
children.push(li("Santint (Китай) — представитель в РФ: ООО «Тинтэко», https://www.santint-rus.ru/"));
children.push(li("Chromaflo / Vibrantz Technologies: https://www.chromaflo.com/"));

children.push(h2("5.3. Российские интеграторы и дилеры"));
children.push(li("Omnitint — универсальная колеровочная система: https://omnitint.ru/ и https://omnitint.ru/paint-shops"));
children.push(li("Tecsa — официальный дилер Fast&Fluid Management: https://tecsa.ru/product-category/equipment/kolerovochnoe-oborudovanie/"));
children.push(li("Tintec — система колеровки под задачу: https://tintec.ru/"));
children.push(li("Kolerovka.ru — каталог оборудования и сервис: https://kolerovka.ru/"));
children.push(li("Kolerovka.com — Corob, Fast&Fluid, Hero сервис: https://kolerovka.com/"));
children.push(li("LKM-kolerovka.ru — каталог Hero, Iris, Green, Datacolor: https://lkm-kolerovka.ru/"));
children.push(li("VPS Color — колеровочные системы: https://vpscolor.ru/"));
children.push(li("Palizh — колеровочное оборудование: https://palizh.com/catalog-industry/kolerovochnoe-oborudovanie/ и https://palizh.info/tinting/"));
children.push(li("Plasters.ru — дилер колеровочного оборудования: https://www.plasters.ru/dealer/kolerovochnoe-oborudovanie/"));
children.push(li("Kovali — автоматическая колеровочная система КОВАКОЛОР: https://kovali.ru/kolerochnaya-sistema-kovacolor/"));
children.push(li("Gravihel — оборудование и инструменты: https://www.gravihel-helios.com/"));
children.push(li("Эко Колор — оборудование: http://ekokolor.ru/oborudovanie.php"));

children.push(h2("5.4. Полезные материалы по мерчендайзингу"));
children.push(li("Академия мерчендайзинга Богачёвой — расстановка оборудования: http://merchbook.com/services/equipment"));

// === ВЫВОДЫ ===
children.push(h1("6. Выводы и рекомендации"));

children.push(p([
  "1. ",
  ["Специализированных буклетов по оформлению зоны колеровки в открытом доступе нет", { bold: true, color: P.accent }],
  ". Производители ЛКМ держат эту информацию закрытой — это часть коммерческой модели. Производители колеровочного оборудования публикуют спецификации машин, но не дизайн-проекты зоны."
]));

children.push(p([
  "2. ",
  ["Лучший публичный материал — PDF-каталог «Тиккурила Симфония»", { bold: true, color: P.positive }],
  ". Это единственный документ, где есть конкретные рекомендации по типам торговых точек (DIY / средний / оптовый), типам оборудования и набору маркетинговых материалов с описанием стенда «Grande Bello»."
]));

children.push(p([
  "3. ",
  ["Для разработки собственного буклета по зоне колеровки", { bold: true, color: P.primary }],
  " можно использовать структуру каталога Tikkurila «Симфония» как образец: разделы «Колеровочное оборудование» (типы дозаторов по типам магазинов), «Маркетинговые материалы» (стенды, веера, колерные карты) и «Демонстрационные стенды» (фото, спецификации, комплектация)."
]));

children.push(p([
  "4. ",
  ["На российском рынке 2026 года", { bold: true, color: P.primary }],
  " ключевой тренд — универсальные системы (Omnitint, VPS Color), которые позволяют магазину работать с разными брендами красок на одной машине. Это снижает зависимость от одного производителя ЛКМ и упрощает организацию зоны колеровки — особенно для DIY-сетей с большим ассортиментом."
]));

children.push(p([
  "5. ",
  ["Для Эмпилс рекомендация", { bold: true, color: P.accent }],
  ": разработать собственный буклет для дистрибьюторов по оформлению зоны колеровки на основе лучших практик Tikkurila. Включить в него: 3 типа торговых точек, рекомендуемое оборудование под каждый тип, комплект POS-материалов (стенды с выкрасками Эмпилс, веера цветов TITAN/PROFIWOOD/ОРЕОЛ/AvantGarde/NEWTONE, колерные карты), инструкцию по подключению к ПО (если будет собственная система) или адаптацию под Omnitint. Это даст конкурентное преимущество перед производителями, которые такой буклет не предоставляют."
]));

// === ДОКУМЕНТ ===
const doc = new Document({
  creator: "Z.ai",
  title: "Анализ конкурентов по колеровочному оборудованию",
  description: "Справка по производителям ЛКМ и колеровочного оборудования",
  styles: {
    default: {
      document: {
        run: { font: { ascii: "Calibri", eastAsia: "Calibri" }, size: 22, color: c(P.body) },
        paragraph: { spacing: { line: 312 } },
      },
    },
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, bottom: 1440, left: 1701, right: 1417 },
      },
    },
    headers: {
      default: new (require("docx").Header)({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { line: 240 },
          children: [new TextRun({
            text: "Колеровочное оборудование в рознице ЛКМ — анализ конкурентов",
            size: 18, color: c(P.muted), italics: true, font: FONT,
          })],
        })],
      }),
    },
    footers: {
      default: new (require("docx").Footer)({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { line: 240 },
          children: [
            new TextRun({ text: "Стр. ", size: 18, color: c(P.muted), font: FONT }),
            new TextRun({ children: [require("docx").PageNumber.CURRENT], size: 18, color: c(P.muted), font: FONT }),
            new TextRun({ text: " из ", size: 18, color: c(P.muted), font: FONT }),
            new TextRun({ children: [require("docx").PageNumber.TOTAL_PAGES], size: 18, color: c(P.muted), font: FONT }),
          ],
        })],
      }),
    },
    children: children,
  }],
});

Packer.toBuffer(doc).then(buffer => {
  const outPath = "/home/z/my-project/download/Колеровочное_оборудование_анализ_конкурентов.docx";
  fs.writeFileSync(outPath, buffer);
  console.log(`✓ Документ сохранён: ${outPath}`);
  console.log(`  Размер: ${buffer.length.toLocaleString()} байт`);
});
