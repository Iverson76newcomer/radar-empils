// Буклет Эмпилс по колеровочному оборудованию для дистрибьюторов
// Подробный, детальный, практический (не концептуальный)

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, PageNumber, AlignmentType, HeadingLevel,
  WidthType, BorderStyle, ShadingType, PageBreak, LevelFormat,
  TabStopType, TabStopPosition,
} = require("docx");
const fs = require("fs");

// === ПАЛИТРА ЭМПИЛС ===
const P = {
  primary: "#1F4E79",      // тёмно-синий — основной фирменный
  primaryLight: "#D6E4F0", // светлый синий — фон таблиц
  accent: "#C00000",        // красный — акценты, важное
  positive: "#375623",     // зелёный — позитив
  warn: "#BF8F00",         // жёлтый — предупреждения
  body: "#1A1A1A",         // основной текст
  secondary: "#595959",    // подписи
  muted: "#808080",         // колонтитулы
  tableHead: "#1F4E79",
  tableRowAlt: "#F2F7FC",
  highlight: "#FFF2CC",   // выделение
};
const c = (hex) => hex.replace("#", "");
const FONT = { ascii: "Calibri", eastAsia: "Calibri" };

// === УТИЛИТЫ ===
function safeText(value, placeholder) {
  if (value === undefined || value === null || value === "" ||
      String(value) === "NaN" || String(value) === "undefined") {
    return placeholder || "—";
  }
  return String(value);
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 180, line: 312 },
    pageBreakBefore: true,
    children: [new TextRun({
      text, bold: true, size: 32, color: c(P.primary), font: FONT,
    })],
  });
}

function h1NoBreak(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 180, line: 312 },
    children: [new TextRun({
      text, bold: true, size: 32, color: c(P.primary), font: FONT,
    })],
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 140, line: 312 },
    children: [new TextRun({
      text, bold: true, size: 28, color: c(P.primary), font: FONT,
    })],
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 220, after: 120, line: 312 },
    children: [new TextRun({
      text, bold: true, size: 24, color: c(P.primary), font: FONT,
    })],
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

function pNoIndent(text, opts = {}) {
  const runs = Array.isArray(text) ? text : [new TextRun({
    text, size: 22, color: c(P.body), font: FONT,
  })];
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 312, after: 140 },
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

function nli(text, opts = {}) {
  const runs = Array.isArray(text) ? text : [new TextRun({
    text, size: 22, color: c(P.body), font: FONT,
  })];
  return new Paragraph({
    spacing: { line: 312, after: 80 },
    numbering: { reference: "main-num", level: 0 },
    ...opts,
    children: runs,
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

function hr() {
  return new Paragraph({
    spacing: { before: 120, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: c(P.primary) } },
    children: [new TextRun({ text: "", size: 2 })],
  });
}

// === ВЫДЕЛЕННЫЙ БЛОК (infobox) ===
function infobox(title, lines, opts = {}) {
  const fillColor = opts.fill || "F2F7FC";
  const titleColor = opts.titleColor || P.primary;
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 8, color: c(titleColor) },
      bottom: { style: BorderStyle.SINGLE, size: 8, color: c(titleColor) },
      left: { style: BorderStyle.SINGLE, size: 2, color: c(titleColor) },
      right: { style: BorderStyle.SINGLE, size: 2, color: c(titleColor) },
      insideHorizontal: { style: BorderStyle.NONE },
      insideVertical: { style: BorderStyle.NONE },
    },
    rows: [new TableRow({
      cantSplit: true,
      children: [new TableCell({
        shading: { type: ShadingType.CLEAR, fill: fillColor },
        margins: { top: 200, bottom: 200, left: 280, right: 280 },
        children: [
          new Paragraph({
            spacing: { line: 312, after: 100 },
            children: [new TextRun({
              text: title, bold: true, size: 24, color: c(titleColor), font: FONT,
            })],
          }),
          ...lines.map(line => {
            if (typeof line === "string") {
              return new Paragraph({
                spacing: { line: 312, after: 60 },
                children: [new TextRun({ text: line, size: 22, color: c(P.body), font: FONT })],
              });
            }
            // line = {text, bold, color}
            return new Paragraph({
              spacing: { line: 312, after: 60 },
              children: [new TextRun({
                text: line.text, size: 22, color: c(line.color || P.body), font: FONT,
                bold: !!line.bold,
              })],
            });
          }),
        ],
      })],
    })],
  });
}

// === ТАБЛИЦА ===
const TB = {
  top: { style: BorderStyle.SINGLE, size: 4, color: c(P.primary) },
  bottom: { style: BorderStyle.SINGLE, size: 4, color: c(P.primary) },
  left: { style: BorderStyle.SINGLE, size: 2, color: "BFBFBF" },
  right: { style: BorderStyle.SINGLE, size: 2, color: "BFBFBF" },
  insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "D0D0D0" },
  insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "D0D0D0" },
};

function tCellHeader(text, widthPct, opts = {}) {
  return new TableCell({
    width: { size: widthPct, type: WidthType.PERCENTAGE },
    shading: { type: ShadingType.CLEAR, fill: c(P.tableHead) },
    margins: { top: 100, bottom: 100, left: 140, right: 140 },
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.CENTER,
      spacing: { line: 280 },
      children: [new TextRun({
        text: safeText(text), bold: true, size: 22, color: "FFFFFF", font: FONT,
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
        text: safeText(text), bold: isBold, size: fontSize, color, font: FONT,
      })],
    })],
  };
  if (fill) cellProps.shading = { type: ShadingType.CLEAR, fill: c(fill) };
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
      text: safeText(line), bold: isBold, size: fontSize, color, font: FONT,
    })],
  }));
  const cellProps = {
    width: { size: widthPct, type: WidthType.PERCENTAGE },
    margins: { top: 80, bottom: 80, left: 140, right: 140 },
    children: paras,
  };
  if (fill) cellProps.shading = { type: ShadingType.CLEAR, fill: c(fill) };
  return new TableCell(cellProps);
}

// ====================== СОДЕРЖАНИЕ БУКЛЕТА ======================
const children = [];

// === РАЗДЕЛ 1. ОБЛОЖКА ===
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 2400, after: 240 },
  children: [new TextRun({
    text: "ЗАО «ЭМПИЛС»",
    bold: true, size: 36, color: c(P.primary), font: FONT,
  })],
}));

children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 1200 },
  children: [new TextRun({
    text: "Ведущий российский производитель ЛКМ",
    italics: true, size: 22, color: c(P.secondary), font: FONT,
  })],
}));

children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 240 },
  border: {
    top: { style: BorderStyle.SINGLE, size: 12, color: c(P.primary) },
    bottom: { style: BorderStyle.SINGLE, size: 12, color: c(P.primary) },
  },
  children: [new TextRun({
    text: "КОЛЕРОВОЧНОЕ ОБОРУДОВАНИЕ",
    bold: true, size: 44, color: c(P.primary), font: FONT,
  })],
}));

children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 1200 },
  children: [new TextRun({
    text: "для дистрибьюторов",
    size: 28, color: c(P.secondary), italics: true, font: FONT,
  })],
}));

children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 200 },
  children: [new TextRun({
    text: "Полный комплект оборудования, колорантов и сопровождения",
    size: 22, color: c(P.body), font: FONT,
  })],
}));

children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 2000 },
  children: [new TextRun({
    text: "Издание 2026 · Версия 1.0",
    size: 20, color: c(P.muted), italics: true, font: FONT,
  })],
}));

children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 1200, after: 100 },
  children: [new TextRun({
    text: "8 (800) 200-49-11",
    bold: true, size: 28, color: c(P.primary), font: FONT,
  })],
}));

children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 100 },
  children: [new TextRun({
    text: "www.empils.ru",
    size: 22, color: c(P.primary), font: FONT,
  })],
}));

children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 100 },
  children: [new TextRun({
    text: "Ростов-на-Дону · 1-й Машиностроительный пер., 21",
    size: 18, color: c(P.muted), font: FONT,
  })],
}));

// === РАЗДЕЛ 2. СОДЕРЖАНИЕ (ОГЛАВЛЕНИЕ) ===
children.push(h1NoBreak("Содержание"));

const toc = [
  ["1.", "Введение: Эмпилс — ведущий российский производитель ЛКМ", "3"],
  ["2.", "Что такое колеровка и зачем она нужна", "4"],
  ["3.", "Компоненты колеровочной системы Эмпилс", "5"],
  ["4.", "Базовые продукты под колеровку", "6"],
  ["5.", "Колоранты MONICOLOR PRO (16 цветов)", "8"],
  ["6.", "Колеровочное оборудование: COROB D180 + MultiSpeed CABO", "10"],
  ["7.", "Программное обеспечение Innovatint", "13"],
  ["8.", "Цветовые веера и каталоги", "14"],
  ["9.", "Маркетинговые материалы", "15"],
  ["10.", "Где ставить оборудование в магазине", "16"],
  ["11.", "Оформление зоны колеровки", "17"],
  ["12.", "Требования к помещению", "18"],
  ["13.", "Этапы установки и пуско-наладки", "19"],
  ["14.", "Условия предоставления оборудования", "20"],
  ["15.", "Эксплуатация и техобслуживание", "21"],
  ["16.", "Сервисная поддержка и контакты", "23"],
];

for (const [num, title, page] of toc) {
  children.push(new Paragraph({
    spacing: { line: 312, after: 80 },
    tabStops: [
      { type: TabStopType.RIGHT, position: 9000 },
    ],
    children: [
      new TextRun({ text: num + " ", bold: true, size: 22, color: c(P.primary), font: FONT }),
      new TextRun({ text: title, size: 22, color: c(P.body), font: FONT }),
      new TextRun({ text: "\t" + page, size: 22, color: c(P.secondary), font: FONT }),
    ],
  }));
}

children.push(new Paragraph({
  spacing: { before: 240, after: 0 },
  children: [new TextRun({
    text: "Как пользоваться буклетом:",
    bold: true, size: 22, color: c(P.primary), font: FONT,
  })],
}));

children.push(pNoIndent([
  "Разделы 4–9 — для менеджеров по закупкам: что закупать и в каком объёме. Разделы 10–13 — для мерчендайзеров и технических специалистов: где и как ставить. Разделы 14–16 — коммерческие условия и сервис.",
]));

// === РАЗДЕЛ 3. ВВЕДЕНИЕ ===
children.push(h1("1. Введение"));

children.push(h2("Эмпилс — ведущий российский производитель ЛКМ"));

children.push(p(
  "ЗАО «Эмпилс» — один из ведущих российских производителей лакокрасочных материалов. История химического производства на площадке в Ростове-на-Дону ведётся с 1920 года, бренд «ЭМПИЛС» существует более 35 лет. Сегодня Эмпилс производит полный спектр ЛКМ — от бытовых красок до промышленных эмалей — на собственном заводе в Ростове-на-Дону (1-й Машиностроительный пер., 21)."
));

children.push(p(
  "Современный портфель брендов Эмпилс включает пять линеек, покрывающих все основные сегменты рынка ЛКМ: AvantGarde (премиум-сегмент: интерьерные и фасадные краски), NEWTONE (средний сегмент: интерьерные моющиеся, фасадные силиконовые краски), ОРЕОЛ (массовый сегмент: интерьерные, фасадные краски, эмали), РАСЦВЕТ (алкидные эмали ПФ-115), ПРОСТОКРАШЕНО (бюджетные эмали ПФ-115). Все пять линеек имеют базовые продукты под автоматическую колеровку — это позволяет дистрибьютору получить единый колеровочный сервис на весь ассортимент."
));

children.push(p([
  "Колеровка в Эмпилс — это не отдельная услуга, а ",
  ["часть цельной системы", { bold: true, color: P.primary }],
  ", объединяющей базовые краски, пигментные пасты (колоранты), рецептуры, программное обеспечение, дозирующее оборудование, смесители, цветовые каталоги и маркетинговые материалы. Все эти составляющие взаимосвязаны и тщательно контролируются на заводе, что гарантирует точное попадание в выбранный цвет от партии к партии."
]));

children.push(pRich([
  ["Цель буклета", { bold: true, color: P.primary }],
  [" — дать дистрибьютору готовое практическое руководство по запуску и эксплуатации зоны колеровки Эмпилс в торговой точке. Буклет содержит конкретные SKU, технические характеристики оборудования, регламенты установки и эксплуатации, коммерческие условия и контактные данные специалистов. После прочтения у дистрибьютора должно сложиться полное понимание: что закупать, где ставить, как оформлять зону, как эксплуатировать и к кому обращаться при сбоях.", {}],
]));

children.push(infobox("Ключевые цифры колеровочной системы Эмпилс", [
  { text: "5 линеек базовых продуктов под колеровку (AvantGarde, NEWTONE, ОРЕОЛ, РАСЦВЕТ, ПРОСТОКРАШЕНО)", bold: true },
  { text: "16 колорантов MONICOLOR PRO (с долей пигмента от 7% до 65%)" },
  { text: "8 поддерживаемых цветовых каталогов (RAL, NCS, Symphony, Caparol, Dulux, Monicolor, Colorgrain)" },
  { text: "Дозатор COROB D180 + автоматический миксер MultiSpeed CABO" },
  { text: "ПО Innovatint 3/4 — управление колеровкой и печать этикеток" },
  { text: "Условная плата 700 руб/мес за оборудование при выполнении плана колеровки" },
], { fill: "F2F7FC" }));

// === РАЗДЕЛ 4. ЧТО ТАКОЕ КОЛЕРОВКА ===
children.push(h1("2. Что такое колеровка и зачем она нужна"));

children.push(h2("Определение"));

children.push(p(
  "Колеровка — это процесс получения краски заданного цвета путём добавления пигментных паст (колорантов) в базовую белую или прозрачную краску непосредственно в точке продажи. В базе данных колеровочного программного обеспечения выбирается нужный цвет, машина автоматически дозирует колоранты по рецептуре в банку с базовой краской, после чего краска перемешивается до однородной массы. Готовый колерованный продукт можно сразу выдавать покупателю."
));

children.push(h2("Преимущества колеровки для дистрибьютора"));

children.push(p(
  "По сравнению с продажей готовых цветных красок (когда каждый цвет — отдельная SKU на складе), колеровочная система даёт ряд неоспоримых преимуществ, которые напрямую влияют на экономику дистрибьюторского бизнеса:"
));

children.push(nli([
  new TextRun({ text: "Неограниченный выбор цветов ", bold: true, color: c(P.primary), font: FONT }),
  new TextRun({ text: "— можно получить краску любого цвета за считанные минуты прямо в магазине. В базе данных Эмпилс — более 8 поддерживаемых каталогов (RAL, NCS, Symphony и др.), суммарно десятки тысяч оттенков.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(nli([
  new TextRun({ text: "Гарантия попадания в цвет ", bold: true, color: c(P.primary), font: FONT }),
  new TextRun({ text: "— пигментные пасты MONICOLOR PRO производятся по жёстким стандартам, рецептуры калиброваны под конкретные базовые краски Эмпилс. Это исключает сюрпризы «получилось не то».", size: 22, color: c(P.body), font: FONT }),
]));

children.push(nli([
  new TextRun({ text: "Повторное воспроизведение цвета ", bold: true, color: c(P.primary), font: FONT }),
  new TextRun({ text: "— если покупатель вернётся через месяц за ещё одной банкой того же цвета, колеровочная машина выдаст идентичный результат с точностью до миллилитра.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(nli([
  new TextRun({ text: "Сохранение свойств краски ", bold: true, color: c(P.primary), font: FONT }),
  new TextRun({ text: "— колоранты MONICOLOR PRO совместимы со всеми базовыми красками Эмпилс и не ухудшают их физико-механические свойства (стойкость к мытью, истиранию, атмосферным воздействиям).", size: 22, color: c(P.body), font: FONT }),
]));

children.push(nli([
  new TextRun({ text: "Экономичность и оптимизация склада ", bold: true, color: c(P.primary), font: FONT }),
  new TextRun({ text: "— вместо 50+ SKU готовых цветных красок на складе достаточно держать 5–6 базовых белых красок и 16 колорантов. Это снижает замороженные в складе средства и упрощает управление остатками.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(h2("Бизнес-смысл для дистрибьютора"));

children.push(p([
  "Колеровка в магазине — это ",
  ["маржинальный инструмент", { bold: true, color: P.accent }],
  ". Стоимость колерованной краски для покупателя выше, чем белой базовой, потому что в неё заложены колоранты, работа оборудования и сервис. Дистрибьютор, инвестирующий в зону колеровки, получает:"
]));

children.push(liBold("Рост среднего чека", "покупатель, выбравший цвет, берёт не 1 банку, а комплект под весь объект (стены, потолок, двери), и готов платить за готовый цвет."));
children.push(liBold("Дифференциация от конкурентов", "магазин с колеровочной зоной выигрывает у тех, кто продаёт только готовые цвета — у него в 20 раз больше выбора оттенков."));
children.push(liBold("Привязку покупателя", "если покупатель заколеровал краску у вас, он вернётся к вам же за дополнительной банкой — другие машины дадут немного другой оттенок."));
children.push(liBold("Услуги профи", "дизайнеры, маляры, бригадиры предпочитают работать с магазинами, где есть колеровка — это стабильный канал B2B-продаж."));

// === РАЗДЕЛ 5. КОМПОНЕНТЫ СИСТЕМЫ ===
children.push(h1("3. Компоненты колеровочной системы Эмпилс"));

children.push(p(
  "Современная высококачественная колеровочная система Эмпилс состоит из шести взаимоувязанных компонентов. Все они контролируются на заводе и калибруются друг под друга — именно поэтому у дистрибьютора всегда получается точный результат."
));

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TB,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tCellHeader("№", 5),
      tCellHeader("Компонент", 25),
      tCellHeader("Что это", 45),
      tCellHeader("Где описан в буклете", 25),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("1", 5, { bold: true, align: AlignmentType.CENTER, fill: "F2F7FC" }),
      tCell("Базовые краски", 25, { bold: true, fill: "F2F7FC" }),
      tCell("Краски с контролируемым цветовым тоном, насыщенностью и заполнением тары, в которые вводятся колоранты. Практически все краски Эмпилс являются базисными и колеруются по системе.", 45, { fill: "F2F7FC" }),
      tCell("Раздел 4", 25, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("2", 5, { bold: true, align: AlignmentType.CENTER }),
      tCell("Пигментные пасты (колоранты)", 25, { bold: true }),
      tCell("Пигментные дисперсии MONICOLOR PRO — 16 цветов с долей пигмента от 7% до 65%. Цвет, цветосила и реологические свойства контролируются.", 45),
      tCell("Раздел 5", 25),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("3", 5, { bold: true, align: AlignmentType.CENTER, fill: "F2F7FC" }),
      tCell("Рецептуры и ПО", 25, { bold: true, fill: "F2F7FC" }),
      tCell("Рецептуры определяют комбинацию и количество пигментных паст для каждого цвета. ПО Innovatint 3/4 управляет дозатором, ведёт статистику, печатает этикетки.", 45, { fill: "F2F7FC" }),
      tCell("Раздел 7", 25, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("4", 5, { bold: true, align: AlignmentType.CENTER }),
      tCell("Дозирующее оборудование", 25, { bold: true }),
      tCell("Дозатор COROB D180 — автоматически дозирует колоранты в банку с базовой краской по рецептуре. Точность — доли миллилитра.", 45),
      tCell("Раздел 6", 25),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("5", 5, { bold: true, align: AlignmentType.CENTER, fill: "F2F7FC" }),
      tCell("Смеситель (миксер)", 25, { bold: true, fill: "F2F7FC" }),
      tCell("MultiSpeed CABO — автоматический гироскопический миксер, перемешивает краску с колорантами до однородной массы.", 45, { fill: "F2F7FC" }),
      tCell("Раздел 6", 25, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("6", 5, { bold: true, align: AlignmentType.CENTER }),
      tCell("Маркетинговые материалы", 25, { bold: true }),
      tCell("Рекламная стойка NEWTONE, цветовые веера RAL K7 и NCS index 1950, каталоги цветов, демонстрационные образцы выкрасок.", 45),
      tCell("Разделы 8–9", 25),
    ]}),
  ],
}));

children.push(caption("Все 6 компонентов — от одного поставщика, что гарантирует их совместимость и точность результата."));

children.push(h2("Дополнительные компоненты (опционально)"));

children.push(p([
  "Помимо базовых шести компонентов, колеровочную систему можно расширить за счёт ",
  ["спектрофотометра", { bold: true }],
  " — прибора, измеряющего цвет образца материала покупателя (обои, ткань, плитка). Это позволяет заколеровать краску точно в цвет образца, даже если неизвестен его код. Эмпилс не включает спектрофотометр в базовый комплект, но при необходимости дистрибьютор может приобрести его отдельно — совместимы модели X-RITE RM 61 (настольный, 500 000 измерений, 1,1 кг) или MINOLTA CM-2500D (портативный, 4 батарейки АА, память 700 измерений)."
]));

// === РАЗДЕЛ 6. БАЗОВЫЕ ПРОДУКТЫ ===
children.push(h1("4. Базовые продукты под колеровку"));

children.push(p(
  "Базовые краски — это краски с контролируемым цветовым тоном, насыщенностью и заполнением тары, в которые вводятся колоранты MONICOLOR PRO для получения требуемого оттенка. Практически все краски Эмпилс поставляются как базисные и колеруются по единой системе. Для автоматической колеровки предназначены 11 SKU из пяти линеек брендов."
));

children.push(h2("4.1. Продукция под автоматическую колеровку"));

children.push(p([
  "Для работы колеровочной машины в торговой точке требуется ",
  ["обязательное наличие", { bold: true, color: P.accent }],
  " в ассортименте следующих базовых продуктов. Это условие предоставления оборудования от Эмпилс (см. раздел 14)."
]));

children.push(h3("Линейка AvantGarde (премиум-сегмент)"));

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TB,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tCellHeader("Наименование продукта", 50),
      tCellHeader("База", 15),
      tCellHeader("Фасовки", 35),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Краска интерьерная моющаяся AvantGarde", 50, { bold: true, fill: "F2F7FC" }),
      tCell("А и С", 15, { align: AlignmentType.CENTER, fill: "F2F7FC" }),
      tCell("0,7 л / 4,5 л / 8,5 л", 35, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Краска фасадная AvantGarde", 50, { bold: true }),
      tCell("А и С", 15, { align: AlignmentType.CENTER }),
      tCell("4,5 л / 8,5 л", 35),
    ]}),
  ],
}));

children.push(p([
  "База А ", { bold: true },
  "(база «альфа») — белая базовая краска, в которую добавляются колоранты для получения светлых и средних оттенков. ",
  "База С ", { bold: true },
  "(база «чарли») — полупрозрачная базовая краска, в которую добавляются колоранты для получения насыщенных тёмных оттенков. Для одного цвета может требоваться либо база А, либо база С — в зависимости от желаемой интенсивности.",
]));

children.push(h3("Линейка NEWTONE (средний сегмент)"));

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TB,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tCellHeader("Наименование продукта", 45),
      tCellHeader("Тип", 25),
      tCellHeader("База", 10),
      tCellHeader("Фасовки", 20),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("NEWTONE INTERIO суперстойкая интерьерная латексная", 45, { bold: true, fill: "F2F7FC" }),
      tCell("Интерьерная", 25, { fill: "F2F7FC" }),
      tCell("А и С", 10, { align: AlignmentType.CENTER, fill: "F2F7FC" }),
      tCell("0,8 / 4,5 / 8,5 л", 20, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("NEWTONE MODERN PRO фасадная силиконовая", 45, { bold: true }),
      tCell("Фасадная", 25),
      tCell("А и С", 10, { align: AlignmentType.CENTER }),
      tCell("0,8 / 4,5 / 8,5 л", 20),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("NEWTONE LOTUS 3 моющаяся интерьерная", 45, { bold: true, fill: "F2F7FC" }),
      tCell("Интерьерная моющаяся", 25, { fill: "F2F7FC" }),
      tCell("А и С", 10, { align: AlignmentType.CENTER, fill: "F2F7FC" }),
      tCell("0,8 / 4,5 / 8,5 л", 20, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("NEWTONE LOTUS 7 моющаяся интерьерная", 45, { bold: true }),
      tCell("Интерьерная моющаяся", 25),
      tCell("А и С", 10, { align: AlignmentType.CENTER }),
      tCell("0,8 / 4,5 / 8,5 л", 20),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("NEWTONE LOTUS 20 моющаяся интерьерная", 45, { bold: true, fill: "F2F7FC" }),
      tCell("Интерьерная моющаяся", 25, { fill: "F2F7FC" }),
      tCell("А и С", 10, { align: AlignmentType.CENTER, fill: "F2F7FC" }),
      tCell("0,8 / 4,5 / 8,5 л", 20, { fill: "F2F7FC" }),
    ]}),
  ],
}));

children.push(h3("Линейка ОРЕОЛ (массовый сегмент)"));

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TB,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tCellHeader("Наименование продукта", 45),
      tCellHeader("Тип", 25),
      tCellHeader("База", 10),
      tCellHeader("Фасовки", 20),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Краска «ОРЕОЛ» для интерьеров моющаяся", 45, { bold: true, fill: "F2F7FC" }),
      tCell("Интерьерная моющаяся", 25, { fill: "F2F7FC" }),
      tCell("А и С", 10, { align: AlignmentType.CENTER, fill: "F2F7FC" }),
      tCell("2,2 л / 9 л", 20, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Краска «ОРЕОЛ» для фасадов", 45, { bold: true }),
      tCell("Фасадная", 25),
      tCell("А и С", 10, { align: AlignmentType.CENTER }),
      tCell("4,5 л / 9 л", 20),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Эмаль акриловая «ОРЕОЛ»", 45, { bold: true, fill: "F2F7FC" }),
      tCell("Эмаль", 25, { fill: "F2F7FC" }),
      tCell("А и С", 10, { align: AlignmentType.CENTER, fill: "F2F7FC" }),
      tCell("0,9 (0,8) кг / 1,9 кг", 20, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Эмаль по пластику «ОРЕОЛ» (белая)", 45, { bold: true }),
      tCell("Эмаль по пластику", 25),
      tCell("—", 10, { align: AlignmentType.CENTER }),
      tCell("0,9 кг", 20),
    ]}),
  ],
}));

children.push(h3("Линейки РАСЦВЕТ и ПРОСТОКРАШЕНО (алкидные эмали)"));

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TB,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tCellHeader("Наименование продукта", 50),
      tCellHeader("Тип", 20),
      tCellHeader("База", 10),
      tCellHeader("Фасовки", 20),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Эмаль «РАСЦВЕТ» ПФ-115 Универсальная (белая)", 50, { bold: true, fill: "F2F7FC" }),
      tCell("Алкидная эмаль", 20, { fill: "F2F7FC" }),
      tCell("—", 10, { align: AlignmentType.CENTER, fill: "F2F7FC" }),
      tCell("0,9 / 1,9 / 2,7 кг", 20, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Эмаль «ПРОСТОКРАШЕНО» ПФ-115 для наружных и внутренних работ (белая)", 50, { bold: true }),
      tCell("Алкидная эмаль", 20),
      tCell("—", 10, { align: AlignmentType.CENTER }),
      tCell("0,9 / 1,9 / 2,7 кг", 20),
    ]}),
  ],
}));

children.push(h2("4.2. Рекомендуемые объёмы колеровки"));

children.push(p([
  "Эмпилс устанавливает ",
  ["минимальные объёмы колеровки", { bold: true, color: P.accent }],
  " для каждого бренда в месяц. Это связано с необходимостью обеспечивать окупаемость оборудования и регулярное обновление колорантов. Целевые показатели — для одной колеровочной машины:"
]));

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TB,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tCellHeader("Бренд", 30),
      tCellHeader("Рекомендуемый объём колеровки в месяц", 35),
      tCellHeader("Годовой объём с одной машины", 35),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("NEWTONE", 30, { bold: true, fill: "F2F7FC" }),
      tCell("1 500 кг", 35, { bold: true, align: AlignmentType.CENTER, color: P.accent, fill: "F2F7FC" }),
      tCell("18 000 кг/год", 35, { align: AlignmentType.CENTER, fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("AvantGarde", 30, { bold: true }),
      tCell("2 000 кг", 35, { bold: true, align: AlignmentType.CENTER, color: P.accent }),
      tCell("24 000 кг/год", 35, { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("ОРЕОЛ", 30, { bold: true, fill: "F2F7FC" }),
      tCell("по плану (не менее 500 кг)", 35, { align: AlignmentType.CENTER, fill: "F2F7FC" }),
      tCell("6 000 кг/год", 35, { align: AlignmentType.CENTER, fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("ИТОГО с одной машины", 30, { bold: true, color: "FFFFFF", fill: P.primary }),
      tCell("≈ 4 000 кг/мес", 35, { bold: true, align: AlignmentType.CENTER, color: "FFFFFF", fill: P.primary }),
      tCell("≈ 48 000 кг/год", 35, { bold: true, align: AlignmentType.CENTER, color: "FFFFFF", fill: P.primary }),
    ]}),
  ],
}));

children.push(infobox("Минимальное условие предоставления оборудования", [
  { text: "Во втором сезоне работы дистрибьютор обязан выйти на объём колеровки не менее 24 тонн/год (2 тонны/мес) с одной машины.", bold: true, color: P.accent },
  { text: "При невыполнении плана Эмпилс оставляет за собой право пересмотреть условия аренды оборудования (700 руб/мес)." },
], { fill: "FFF2CC", titleColor: P.warn }));

// === РАЗДЕЛ 7. КОЛОРАНТЫ MONICOLOR PRO ===
children.push(h1("5. Колоранты MONICOLOR PRO"));

children.push(p(
  "Пигментные пасты MONICOLOR PRO — это пигментные дисперсии, цвет, цветосила и реологические свойства которых тщательно контролируются на производстве. Они изготавливаются по жёстким стандартам, гарантирующим точную повторяемость цвета при каждой колеровке. Колеранты MONICOLOR PRO совместимы со всеми базовыми красками Эмпилс и не снижают их физико-механические свойства."
));

children.push(h2("5.1. Состав колорантного набора"));

children.push(p(
  "Колерантный набор MONICOLOR PRO включает 16 цветов с различной долей пигмента (от 7% до 65%). Доля пигмента определяет красящую способность колоранта — чем выше процент, тем меньше колоранта нужно для получения насыщенного цвета. Это важно для расчёта расхода колорантов и планирования закупок."
));

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TB,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tCellHeader("№", 7),
      tCellHeader("Код колоранта", 23),
      tCellHeader("Доля пигмента, %", 20),
      tCellHeader("Группа по насыщенности", 25),
      tCellHeader("Уровень расхода", 25),
    ]}),
    ...[
      [1, "MONICOLOR PRO 1", 65, "Высококонцентрированный", "Очень низкий"],
      [2, "MONICOLOR PRO 2", 20, "Среднеконцентрированный", "Средний"],
      [3, "MONICOLOR PRO 3", 37, "Среднеконцентрированный", "Средний"],
      [4, "MONICOLOR PRO 4", 20, "Среднеконцентрированный", "Средний"],
      [5, "MONICOLOR PRO 5", 55, "Высококонцентрированный", "Низкий"],
      [6, "MONICOLOR PRO 6", 29, "Среднеконцентрированный", "Средний"],
      [7, "MONICOLOR PRO 7", 10, "Низкоконцентрированный", "Высокий"],
      [8, "MONICOLOR PRO 8", 37, "Среднеконцентрированный", "Средний"],
      [9, "MONICOLOR PRO 9", 65, "Высококонцентрированный", "Очень низкий"],
      [10, "MONICOLOR PRO 10", 25, "Среднеконцентрированный", "Средний"],
      [11, "MONICOLOR PRO 11", 10, "Низкоконцентрированный", "Высокий"],
      [12, "MONICOLOR PRO 12", 7, "Низкоконцентрированный", "Очень высокий"],
      [13, "MONICOLOR PRO 13", 45, "Высококонцентрированный", "Низкий"],
      [14, "MONICOLOR PRO 14", 10, "Низкоконцентрированный", "Высокий"],
      [15, "MONICOLOR PRO 15", 40, "Высококонцентрированный", "Низкий"],
      [16, "MONICOLOR PRO 16", 35, "Среднеконцентрированный", "Средний"],
    ].map((row, idx) => {
      const fill = idx % 2 === 1 ? "F2F7FC" : null;
      return new TableRow({ cantSplit: true, children: [
        tCell(row[0], 7, { bold: true, align: AlignmentType.CENTER, fill }),
        tCell(row[1], 23, { bold: true, fill }),
        tCell(row[2] + "%", 20, { align: AlignmentType.CENTER, color: row[2] >= 45 ? P.accent : P.body, bold: row[2] >= 45, fill }),
        tCell(row[3], 25, { fill }),
        tCell(row[4], 25, { fill }),
      ]});
    }),
  ],
}));

children.push(h2("5.2. Хранение и подготовка колорантов"));

children.push(p([
  "Колоранты MONICOLOR PRO поставляются в ",
  ["фасовке по 950 мл", { bold: true }],
  " — это объём одной заправки колоранта в канистру дозатора. Стандартные канистры COROB D180 имеют объём 1,5 л или 2,5 л, но колоранты поставляются в более мелкой фасовке 950 мл (для удобства логистики и хранения). Хранить колоранты следует при температуре ",
  ["не ниже +15°C", { bold: true, color: P.accent }],
  " — при более низкой температуре они густеют, что приводит к некорректной дозировке и сбоям в работе дозатора."
]));

children.push(h3("Правила работы с колорантами"));

children.push(liBold("Перемешивание перед заливкой", "перед каждой заливкой колоранты перемешиваются в миксере в течение 1 минуты. Колоранты с высокой вязкостью перемешиваются 3 минуты."));
children.push(liBold("Удаление воздуха при первичной заливке", "при первой заливке колоранта в канистру необходимо прокачать насос bellows pump 5–6 раз для удаления воздуха из системы."));
children.push(liBold("Контроль уровня", "уровень красителя в колбе должен быть не ниже 900 мл. При падении ниже — долить колорант."));
children.push(liBold("Запрет на перемешивание миксером", "при добавлении колоранта в дозатор НЕ допускается его перемешивание миксером — только вручную, в течение 1 минуты."));
children.push(liBold("Хранение", "хранить канистры с колорантом в вертикальном положении, вдали от прямых солнечных лучей и отопительных приборов."));

// === РАЗДЕЛ 8. КОЛЕРОВОЧНОЕ ОБОРУДОВАНИЕ ===
children.push(h1("6. Колеровочное оборудование"));

children.push(p(
  "Эмпилс предоставляет дистрибьюторам комплект колеровочного оборудования, включающий дозатор и автоматический миксер. Оборудование передаётся по договору аренды с условной платой 700 руб/мес (см. раздел 14). Пуско-наладка и обучение персонала проводятся специалистами Эмпилс бесплатно."
));

children.push(h2("6.1. Дозатор COROB D180"));

children.push(p([
  "Дозатор COROB D180 — ",
  ["автоматический дозатор колорантов", { bold: true, color: P.primary }],
  " производства Corob Group (Италия) — одного из мировых лидеров в производстве колеровочного оборудования. COROB D180 предназначен для розничной колеровки красок в торговых точках средних и крупных размеров. Оборудование сочетает в себе надёжность насосов bellows pump (гофрированный мембранный шланговый насос), точность дозирования и удобство ежедневного обслуживания."
]));

children.push(h3("Технические характеристики COROB D180"));

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TB,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tCellHeader("Параметр", 40),
      tCellHeader("Значение", 60),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Производитель", 40, { bold: true, fill: "F2F7FC" }),
      tCell("Corob Group S.p.A., Италия", 60, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Тип", 40, { bold: true }),
      tCell("Автоматический дозатор колорантов", 60),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Количество канистр для колорантов", 40, { bold: true, fill: "F2F7FC" }),
      tCellMulti([
        "16 канистр (соответствует колерантному набору MONICOLOR PRO)",
        "Объём канистры: 1,5 л или 2,5 л (стандартные типоразмеры COROB)",
      ], 60, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Тип дозирования", 40, { bold: true }),
      tCell("Bellows pump (гофрированный/мембранный шланговый насос)", 60, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Пробойник", 40, { bold: true, fill: "F2F7FC" }),
      tCell("Ручной пробойник Ø25 мм", 60, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Стол для банок", 40, { bold: true }),
      tCell("Ручная выдвижная полка для банки (регулируемая по высоте)", 60, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Система увлажнения форсунок", 40, { bold: true, fill: "F2F7FC" }),
      tCell("Патентованная система увлажнения форсунок (минимизирует очистку)", 60, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Электропитание", 40, { bold: true }),
      tCell("100–240 В ±10%, 50/60 Гц, макс. 150 Вт", 60, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Подключение", 40, { bold: true, fill: "F2F7FC" }),
      tCellMulti([
        "Сечение провода: не менее 2,5 мм²",
        "Обязательное заземление",
        "Евро-розетка",
      ], 60, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Требования к полу", 40, { bold: true }),
      tCell("Ровный бетонный пол площадью не менее 1 м² в месте установки", 60),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Температура эксплуатации", 40, { bold: true, fill: "F2F7FC" }),
      tCell("Не ниже +15°C (при более низкой температуре колоранты густеют)", 60, { fill: "F2F7FC", color: P.accent, bold: true }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Минимальная доза: 1/576 fl.oz. (≈0,051 мл)", 60, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCellMulti([
        "Габариты: 1250 × 710 × 880 мм",
        "Вес: 90 кг (с упаковкой 115 кг)",
        "Площадь основания: 0,62 м²",
      ], 60, { fill: "F2F7FC" }),
    ]}),
  ],
}));

children.push(caption("Точные габариты и вес COROB D180 уточняйте у специалистов Эмпилс на момент установки. Характеристики могут незначительно отличаться в зависимости от модификации."));

children.push(h3("Особенности работы COROB D180"));

children.push(p([
  "Дозатор COROB D180 построен на базе ",
  ["bellows pump (гофрированных мембранных шланговых насосов)", { bold: true }],
  " — проверенной технологии, обеспечивающей точную дозировку пигментных паст. Машина способна дозировать даже самые маленькие дозы колоранта (до 0,051 мл), что критично для светлых оттенков. Bellows pump устойчив к загустевшим колорантам и обеспечивает длительный срок службы без замены."
]));

children.push(p([
  "Система ",
  ["увлажнения форсунок", { bold: true }],
  " минимизирует засорение и снижает частоту очистки. После каждой колеровки форсунки автоматически промываются, что предотвращает высыхание колоранта на соплах. Тем не менее, ежедневная промывка и продувка дозатора в начале и в конце рабочего дня — обязательная процедура (см. раздел 15)."
]));

children.push(p([
  "Ручная выдвижная полка для банки ", { bold: true },
  "позволяет устанавливать банки разной высоты (от 70 до 375 мм, диаметр 46–320 мм, макс. вес 35 кг) без перестройки машины. Оператор просто устанавливает банку на полку, выдвигает её к дозирующей головке, и машина начинает дозирование.",
]));

children.push(h2("6.2. Автоматический миксер MultiSpeed CABO"));

children.push(p([
  "MultiSpeed CABO — ",
  ["автоматический гироскопический миксер", { bold: true, color: P.primary }],
  " производства CABO S.r.l. (Италия, Paderno Dugnano). В отличие от вибрационных шейкеров, MultiSpeed CABO использует гироскопическое вращение (1:1,8) с 3 программируемыми скоростями 80–180 об/мин. После того как дозатор COROB D180 ввёл колоранты в банку с базовой краской, банка устанавливается в миксер MultiSpeed CABO, который перемешивает содержимое до однородной массы. Весь процесс колеровки — от выбора цвета до готовой банки — занимает считанные минуты."
]));

children.push(h3("Технические характеристики MultiSpeed CABO"));

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TB,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tCellHeader("Параметр", 40),
      tCellHeader("Значение", 60),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Производитель", 40, { bold: true, fill: "F2F7FC" }),
      tCell("CABO S.r.l., Италия (Paderno Dugnano)", 60, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Тип", 40, { bold: true }),
      tCell("Автоматический гироскопический миксер (НЕ вибрационный шейкер)", 60),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Скорости", 40, { bold: true, fill: "F2F7FC" }),
      tCell("3 программируемые скорости: 80–180 об/мин, гироскопическая скорость 1:1,8", 60, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Время смешивания", 40, { bold: true }),
      tCell("20 секунд — 30 минут (программируется)", 60),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Макс. вес тары", 40, { bold: true, fill: "F2F7FC" }),
      tCell("40 кг", 60, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Высота тары", 40, { bold: true }),
      tCell("90–490 мм", 60),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Макс. диаметр тары", 40, { bold: true, fill: "F2F7FC" }),
      tCell("380 мм", 60, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Форма тары", 40, { bold: true }),
      tCell("Круглая / квадратная / овальная", 60),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Электропитание", 40, { bold: true, fill: "F2F7FC" }),
      tCell("220–240 В, 50/60 Гц", 60, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Макс. потребляемая мощность", 40, { bold: true }),
      tCell("1,2 кВт", 60),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Габариты (Ш×В×Г)", 40, { bold: true, fill: "F2F7FC" }),
      tCell("800 × 860 × 1040 мм", 60, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Вес", 40, { bold: true }),
      tCell("170 кг (брутто 190 кг)", 60),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Тип открытия", 40, { bold: true, fill: "F2F7FC" }),
      tCell("Рольставни", 60, { fill: "F2F7FC" }),
    ]}),
  ],
}));

children.push(caption("Источник: официальная спецификация CABO S.r.l. через российского импортёра ООО «АРТ индустрия» (art-industria.ru)."));

children.push(h3("Особенности работы MultiSpeed CABO"));

children.push(p([
  "Миксер MultiSpeed CABO оснащён ",
  ["автоматическим зажимом банки", { bold: true }],
  " — оператор просто устанавливает банку и закрывает рольставни. Машина сама определяет размер тары и подбирает оптимальный режим смешивания. Разная скорость для больших и маленьких банок предотвращает разбрызгивание и обеспечивает равномерное перемешивание."
]));

children.push(p([
  "Время смешивания составляет ",
  ["20 секунд — 30 минут (программируется)", { bold: true }],
  " в зависимости от объёма тары и типа краски. После завершения смешивания оператор открывает дверцу и забирает готовую колерованную краску. Каждая партия проверяется визуально — краска должна быть однородной по цвету, без разводов и неперемешанных участков."
]));

children.push(h2("6.3. Принтер этикеток"));

children.push(p(
  "Помимо дозатора и миксера, в комплект зоны колеровки рекомендуется включить принтер этикеток. Хотя он не входит в базовый комплект оборудования Эмпилс, без него невозможно организовать полноценный сервис для покупателя. Принтер позволяет печатать наклейки со следующей информацией: название продукции, код цвета, размер тары, цену базисной краски и пигментов, штрих-код, дату колеровки."
));

children.push(infobox("Рекомендуемая модель принтера: Seiko SLP-430", [
  { text: "Размер бумаги: 54 мм" },
  { text: "Разрешение: 203 DPI" },
  { text: "Технология печати: термопечать" },
  { text: "Интерфейс: USB и Serial (RS232)" },
  { text: "Скорость печати: 2–3 секунды на наклейку" },
  { text: "Размеры: 157,9 × 117 × 160,9 мм" },
  { text: "Совместимость с ПО Innovatint: поддерживается" },
], { fill: "F2F7FC" }));

// === РАЗДЕЛ 9. ПО INNOVATINT ===
children.push(h1("7. Программное обеспечение Innovatint"));

children.push(p([
  "Программное обеспечение ",
  ["Innovatint", { bold: true, color: P.primary }],
  " (версии 3 и 4) — это управляющая программа для колеровочного оборудования. Она соединяет в себе все компоненты колеровочной системы: базу данных рецептур, статистику колеровок, расчёт стоимости и печать этикеток. Innovatint работает с дозатором COROB D180 и обеспечивает точную дозировку колорантов MONICOLOR PRO по выбранному цвету."
]));

children.push(h2("7.1. Основные функции Innovatint"));

children.push(liBold("Поиск формул по цвету", "оператор выбирает цвет из каталога (RAL, NCS, Symphony и др.), и программа автоматически подбирает рецептуру — какие колоранты и в каком количестве добавить в базовую краску."));
children.push(liBold("Управление дозатором", "Innovatint передаёт команды на COROB D180 — какие колоранты дозировать, в каком объёме, в какой последовательности. Дозатор выполняет команды автоматически."));
children.push(liBold("Расчёт себестоимости", "программа рассчитывает стоимость колерованной краски: цена базовой краски + стоимость колорантов + наценка. Это позволяет сразу назвать покупателю итоговую цену."));
children.push(liBold("Статистика колеровок", "Innovatint ведёт историю всех колеровок: дата, время, цвет, объём, оператор. Это помогает анализировать спрос и планировать закупки колорантов."));
children.push(liBold("Печать этикеток", "программа печатает наклейки с названием продукции, кодом цвета, размером тары, ценой, штрих-кодом — всё, что нужно для маркировки готовой банки."));
children.push(liBold("Связь с базой рецептур", "Innovatint содержит рецептуры под все базовые краски Эмпилс (AvantGarde, NEWTONE, ОРЕОЛ, РАСЦВЕТ, ПРОСТОКРАШЕНО) и обновляется онлайн при появлении новых продуктов."));

children.push(h2("7.2. Поддерживаемые цветовые каталоги"));

children.push(p(
  "Одно из ключевых преимуществ ПО Innovatint — поддержка 8 крупнейших мировых цветовых каталогов. Это означает, что покупатель может прийти с веером RAL, NCS или Symphony — и вы заколеруете краску точно в выбранный цвет. Не нужно отправлять покупателя к другому магазину, потому что «у нас такого каталога нет»."
));

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TB,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tCellHeader("№", 5),
      tCellHeader("Каталог", 35),
      tCellHeader("Описание", 40),
      tCellHeader("Тип", 20),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("1", 5, { bold: true, align: AlignmentType.CENTER, fill: "F2F7FC" }),
      tCell("RAL 840 HR", 35, { bold: true, fill: "F2F7FC" }),
      tCell("Классический немецкий стандарт цветов RAL Classic (213 цветов)", 40, { fill: "F2F7FC" }),
      tCell("Международный", 20, { align: AlignmentType.CENTER, fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("2", 5, { bold: true, align: AlignmentType.CENTER }),
      tCell("RAL Design D2", 35, { bold: true }),
      tCell("Расширенный стандарт RAL Design (1 625 цветов)", 40),
      tCell("Международный", 20, { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("3", 5, { bold: true, align: AlignmentType.CENTER, fill: "F2F7FC" }),
      tCell("NCS second Edition 2009", 35, { bold: true, fill: "F2F7FC" }),
      tCell("Natural Color System — скандинавский стандарт (1 950 цветов)", 40, { fill: "F2F7FC" }),
      tCell("Международный", 20, { align: AlignmentType.CENTER, fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("4", 5, { bold: true, align: AlignmentType.CENTER }),
      tCell("Symphony I and II", 35, { bold: true }),
      tCell("Фирменная система Tikkurila (2 436 цветов)", 40),
      tCell("Фирменный", 20, { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("5", 5, { bold: true, align: AlignmentType.CENTER, fill: "F2F7FC" }),
      tCell("Monicolor Nova 2024", 35, { bold: true, fill: "F2F7FC" }),
      tCell("Система цветов Nobel / AkzoNobel (1 600 цветов)", 40, { fill: "F2F7FC" }),
      tCell("Фирменный", 20, { align: AlignmentType.CENTER, fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("6", 5, { bold: true, align: AlignmentType.CENTER }),
      tCell("Caparol 3D PLUS", 35, { bold: true }),
      tCell("Фирменная система Caparol (1 600 цветов)", 40),
      tCell("Фирменный", 20, { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("7", 5, { bold: true, align: AlignmentType.CENTER, fill: "F2F7FC" }),
      tCell("Dulux Trade Color Palette 5", 35, { bold: true, fill: "F2F7FC" }),
      tCell("Фирменная система Dulux / AkzoNobel (1 200 цветов)", 40, { fill: "F2F7FC" }),
      tCell("Фирменный", 20, { align: AlignmentType.CENTER, fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("8", 5, { bold: true, align: AlignmentType.CENTER }),
      tCell("Colorgrain 2016", 35, { bold: true }),
      tCell("Система для текстурных покрытий (защита дерева)", 40),
      tCell("Специализированный", 20, { align: AlignmentType.CENTER }),
    ]}),
  ],
}));

children.push(infobox("Итого доступно цветов", [
  { text: "Суммарно через ПО Innovatint дистрибьютор Эмпилс может предложить покупателю свыше 12 000 цветов из 8 каталогов — от классических RAL и NCS до фирменных систем Tikkurila, Dulux и Caparol.", bold: true, color: P.accent },
], { fill: "FFF2CC", titleColor: P.warn }));

// === РАЗДЕЛ 10. ЦВЕТОВЫЕ ВЕЕРА И КАТАЛОГИ ===
children.push(h1("8. Цветовые веера и каталоги"));

children.push(p(
  "Для выбора цвета покупателем в торговой точке Эмпилс предоставляет два основных цветовых веера, которые входят в базовый комплект оборудования:"
));

children.push(h2("8.1. Веер RAL K7"));

children.push(p([
  "RAL K7 — ",
  ["классический веер цветов RAL Classic", { bold: true, color: P.primary }],
  ". Содержит 213 цветов — самые популярные оттенки, используемые в промышленности, строительстве и дизайне. Формат: 5 цветов на лист, размер 50×200 мм. Веер компактный, легко помещается в карман продавца-консультанта. Каждый цвет имеет четырёхзначный код (например, RAL 9010 — белый, RAL 9005 — чёрный, RAL 3009 — красно-коричневый)."
]));

children.push(p([
  "RAL K7 — это ",
  ["международный стандарт", { bold: true }],
  ", узнаваемый всеми — от маляров до архитекторов. Если покупатель говорит «мне нужен RAL 7016», продавец сразу понимает, о каком серо-графитовом цвете речь. Это упрощает коммуникацию и ускоряет обслуживание."
]));

children.push(h2("8.2. Веер NCS index 1950"));

children.push(p([
  "NCS index 1950 — ",
  ["веер Natural Color System", { bold: true, color: P.primary }],
  " — скандинавского стандарта цветов. Содержит 1 950 цветов — значительно больше, чем RAL. NCS описывает цвета через их восприятие человеком: чёрный/белый, жёлтый/красный/синий/зелёный, насыщенность. Это позволяет точно передать цвет словами."
]));

children.push(p([
  "NCS особенно популярен среди ",
  ["дизайнеров интерьеров и архитекторов", { bold: true }],
  " — они часто работают именно в этой системе. Если ваш магазин работает с профессиональными покупателями (дизайн-студии, строительные компании), веер NCS обязателен. Каждый цвет в NCS имеет код вида «NCS S 1500-N» (нейтральный серый) или «NCS S 3010-Y90R» (тёмно-красный с жёлтым оттенком)."
]));

children.push(h2("8.3. Дополнительные каталоги"));

children.push(p(
  "Помимо базовых вееров RAL K7 и NCS index 1950, в ПО Innovatint поддерживаются ещё 6 каталогов (см. раздел 7.2). Если покупатель приходит с веером Symphony, Monicolor Nova, Caparol или Dulux — вы можете заколеровать краску точно в его цвет. Дополнительно можно заказать у Эмпилс фирменные колерные карты под каждый бренд (AvantGarde, NEWTONE, ОРЕОЛ) для демонстрации на стенде."
));

// === РАЗДЕЛ 11. МАРКЕТИНГОВЫЕ МАТЕРИАЛЫ ===
children.push(h1("9. Маркетинговые материалы"));

children.push(p(
  "Помимо оборудования и колорантов, Эмпилс предоставляет дистрибьюторам маркетинговые материалы для оформления зоны колеровки. Эти материалы привлекают внимание покупателей, помогают им выбрать цвет и создают единый фирменный стиль точки продаж."
));

children.push(h2("9.1. Рекламная стойка NEWTONE"));

children.push(p([
  "В базовый комплект оборудования Эмпилс входит ",
  ["рекламная стойка NEWTONE", { bold: true, color: P.primary }],
  " — фирменный POS-стенд с брендингом NEWTONE. Стойка предназначена для размещения в зоне колеровки и привлекает внимание покупателей к возможности колеровки красок в colours NEWTONE."
]));

children.push(p([
  "Стойка NEWTONE выполняет несколько функций: ",
  ["информационную", { bold: true }],
  " (сообщает покупателю о возможности получить любой цвет), ",
  ["демонстрационную", { bold: true }],
  " (показывает образцы выкрасок цветов NEWTONE) и ",
  ["навигационную", { bold: true }],
  " (направляет покупателя к зоне колеровки)."
]));

children.push(h2("9.2. Фирменные веера цветов"));

children.push(p([
  "Эмпилс предоставляет дистрибьюторам ",
  ["два базовых цветовых веера", { bold: true }],
  " — RAL K7 и NCS index 1950 (см. раздел 8). Дополнительно можно заказать фирменные веера под каждый бренд:"
]));

children.push(liBold("Веер AvantGarde", "фильтрованная палитра цветов, оптимизированных под базовые краски AvantGarde. Включает популярные интерьерные и фасадные оттенки."));
children.push(liBold("Веер NEWTONE", "палитра цветов для бренда NEWTONE — от пастельных интерьерных до насыщенных фасадных."));
children.push(liBold("Веер ОРЕОЛ", "палитра цветов для массового сегмента — базовые цвета, которые выбирают чаще всего."));
children.push(liBold("Колерные карты РАСЦВЕТ и ПРОСТОКРАШЕНО", "буклеты с цветами для алкидных эмалей ПФ-115."));

children.push(h2("9.3. Демонстрационные образцы выкрасок"));

children.push(p([
  "Для оформления зоны колеровки Эмпилс рекомендует использовать ",
  ["демонстрационные образцы выкрасок", { bold: true }],
  " — это листы с нанесённой краской разных цветов, которые показывают, как будет выглядеть оттенок на стене. Выкраски размещаются на стенде в зоне колеровки и помогают покупателю визуально оценить цвет. В отличие от бумажных вееров, выкраски дают реальное представление о фактуре и насыщенности."
]));

children.push(p(
  "По аналогии с практикой Tikkurila (стенды Grande Bello, Piccolo Bello), Эмпилс рекомендует следующие элементы оформления:"
));

children.push(liBold("Стенд с выкрасками", "перелистываемые модули с образцами цветов по каталогам RAL и NCS. Размещается на стене за колеровочной машиной."));
children.push(liBold("Карман для проспектов", "стеллаж для брошюр и колерных карт, которые покупатель может взять с собой."));
children.push(liBold("Образцы накрасок А4", "крупные образцы цветов для наглядной демонстрации (по аналогии с башней «Тайка» Tikkurila)."));

// === РАЗДЕЛ 12. ГДЕ СТАВИТЬ ОБОРУДОВАНИЕ ===
children.push(h1("10. Где ставить оборудование в магазине"));

children.push(p(
  "Правильное размещение зоны колеровки в торговом зале напрямую влияет на объём колеровки и, соответственно, на окупаемость оборудования. По аналогии с классификацией Tikkurila, Эмпилс различает три типа торговых точек — под каждый тип рекомендуются свои решения по размещению."
));

children.push(h2("10.1. Типы торговых точек"));

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TB,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tCellHeader("Тип магазина", 25),
      tCellHeader("Площадь, м²", 15),
      tCellHeader("Рекомендуемое место зоны колеровки", 35),
      tCellHeader("Площадь зоны, м²", 25),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Маленький DIY", 25, { bold: true, fill: "F2F7FC" }),
      tCell("до 100", 15, { align: AlignmentType.CENTER, fill: "F2F7FC" }),
      tCell("У дальней стены или в углу. Зона видна с входа, но не загромождает проход. Компактная стойка 1×1 м.", 35, { fill: "F2F7FC" }),
      tCell("1–1,5 м²", 25, { align: AlignmentType.CENTER, fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Средний", 25, { bold: true }),
      tCell("100–500", 15, { align: AlignmentType.CENTER }),
      tCell("Отдельный «колеровочный остров» в центре зала или вдоль стены. Зона выделяется светом и фирменными стендами.", 35),
      tCell("2–4 м²", 25, { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Крупный DIY-гипермаркет", 25, { bold: true, fill: "F2F7FC" }),
      tCell("500+", 15, { align: AlignmentType.CENTER, fill: "F2F7FC" }),
      tCell("Отдельная зона 4×4 м с полноценным стендом, компьютером, принтером этикеток. Возможно несколько машин при больших объёмах.", 35, { fill: "F2F7FC" }),
      tCell("4–6 м²", 25, { align: AlignmentType.CENTER, fill: "F2F7FC" }),
    ]}),
  ],
}));

children.push(h2("10.2. Принципы размещения"));

children.push(p("Независимо от типа магазина, при выборе места для зоны колеровки соблюдайте следующие принципы:"));

children.push(nli([
  new TextRun({ text: "Видимость с входа. ", bold: true, color: c(P.primary), font: FONT }),
  new TextRun({ text: "Зона колеровки должна быть видна покупателю сразу при входе в магазин. Это сразу сообщает: «у нас можно получить любой цвет».", size: 22, color: c(P.body), font: FONT }),
]));

children.push(nli([
  new TextRun({ text: "Рядом с красками. ", bold: true, color: c(P.primary), font: FONT }),
  new TextRun({ text: "Зона колеровки должна располагаться рядом с полками базовых красок (AvantGarde, NEWTONE, ОРЕОЛ), чтобы покупатель мог взять банку и сразу пройти к колеровке.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(nli([
  new TextRun({ text: "Доступность для обслуживания. ", bold: true, color: c(P.primary), font: FONT }),
  new TextRun({ text: "Оставьте пространство для подхода к машине с технической стороны — для замены колорантов, очистки форсунок, ремонта. Минимум 0,5 м со стороны технического доступа.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(nli([
  new TextRun({ text: "Стабильный пол. ", bold: true, color: c(P.primary), font: FONT }),
  new TextRun({ text: "Миксер MultiSpeed CABO создаёт вибрационную нагрузку — ровный бетонный пол обязателен. ", size: 22, color: c(P.body), font: FONT, bold: true }),
  new TextRun({ text: "Важно: гироскопический миксер MultiSpeed CABO НЕ создаёт вибрационной нагрузки (в отличие от вибрационных шейкеров). Пол в месте установки должен быть ровным для устойчивости оборудования, но не обязательно вибростойким.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(nli([
  new TextRun({ text: "Электропитание. ", bold: true, color: c(P.primary), font: FONT }),
  new TextRun({ text: "Минимум две евро-розетки рядом с местом установки дозатора и миксера. Питание по проводам сечением не менее 2,5 мм² с заземлением.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(nli([
  new TextRun({ text: "Температура. ", bold: true, color: c(P.primary), font: FONT }),
  new TextRun({ text: "Температура в помещении не должна опускаться ниже +15°C — иначе колоранты густеют и дозатор работает некорректно.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(nli([
  new TextRun({ text: "Освещение. ", bold: true, color: c(P.primary), font: FONT }),
  new TextRun({ text: "Желательно три типа освещения (дневной свет, люминесцентная лампа, лампа накаливания) — для оценки цвета при разном свете (как у стенда Grande Bello Tikkurila). Минимум — хорошая общая подсветка.", size: 22, color: c(P.body), font: FONT }),
]));

// === РАЗДЕЛ 13. ОФОРМЛЕНИЕ ЗОНЫ ===
children.push(h1("11. Оформление зоны колеровки"));

children.push(p([
  "По аналогии с лучшими практиками Tikkurila (буклет «Симфония», раздел «Маркетинговые материалы»), ",
  ["оформление зоны колеровки", { bold: true, color: P.primary }],
  " должно создавать единый фирменный стиль и помогать покупателю выбрать цвет. Ниже — принципы и чек-лист оформления."
]));

children.push(h2("11.1. Принципы оформления"));

children.push(h3("Серый цвет алюминия как нейтральный фон"));

children.push(p(
  "Серый цвет алюминия — нейтральный фон, который подчёркивает яркость выкрасок и скрывает пыль/грязь, неизбежно появляющиеся в зоне колеровки со временем. Используйте серый ламинат или окрашенные панели для стенда и backdrop."
));

children.push(h3("Белый парус с логотипом"));

children.push(p([
  "Тыльная сторона стенда — ",
  ["белый парус с логотипом Эмпилс", { bold: true }],
  " (или брендом AvantGarde / NEWTONE / ОРЕОЛ). Белый цвет создаёт контраст с серым фоном и фокусирует внимание покупателя на зоне. Парус подсвечивается — это даёт эффект «витрины»."
]));

children.push(h3("Подсветка с 3 типами света"));

children.push(p([
  "Критично для выбора цвета: покупатель должен видеть, как цвет будет выглядеть ",
  ["при разном освещении дома", { bold: true, color: P.accent }],
  ". Поэтому стенд зоны колеровки должен иметь 3 типа подсветки:"
]));

children.push(liBold("Лампа дневного света", "имитирует естественное освещение — основной режим, в котором покупатель выбирает цвет."));
children.push(liBold("Люминесцентная лампа", "имитирует офисное / торговое освещение — показывает, как цвет будет выглядеть в общественных помещениях."));
children.push(liBold("Лампа накаливания", "имитирует тёплое домашнее освещение — показывает, как цвет будет выглядеть в жилых комнатах с лампами накаливания."));

children.push(h3("Шкафчики для хранения"));

children.push(p(
  "В нижней части стенда — двухдверные шкафчики для хранения проспектов, колерных карт, расходных материалов (промывочная жидкость, салфетки, запасные этикетки). Это поддерживает порядок в зоне колеровки и ускоряет работу оператора."
));

children.push(h3("Передвижные каретки"));

children.push(p(
  "Мобильные стойки с проспектами и колерными картами, которые можно поднести к покупателю, чтобы он выбрал цвет. Это особенно удобно в крупных магазинах, где покупатель может стоять в 5–10 метрах от стенда."
));

children.push(h2("11.2. Чек-лист оформления зоны колеровки"));

children.push(p("Используйте этот чек-лист при запуске зоны колеровки:"));

const checklist = [
  ["Дозатор COROB D180", "установлен на ровном бетонном полу, подключён к электросети с заземлением"],
  ["Миксер MultiSpeed CABO", "установлен рядом с дозатором, подключён к электросети"],
  ["Компьютер с ПО Innovatint", "установлен, активирован, соединён с дозатором"],
  ["Принтер этикеток", "установлен, подключён, заправлен термобумагой 54 мм"],
  ["16 колорантов MONICOLOR PRO", "залиты в канистры дозатора, уровень не ниже 900 мл"],
  ["Веер RAL K7", "размещён на стенде, доступен покупателю"],
  ["Веер NCS index 1950", "размещён на стенде, доступен покупателю"],
  ["Рекламная стойка NEWTONE", "установлена в зоне видимости с входа в магазин"],
  ["Стенд с выкрасками", "размещён на стене за колеровочной машиной"],
  ["Белый парус с логотипом", "установлен с тыльной стороны, подсвечен"],
  ["3 типа освещения", "дневной свет, люминесцентная, лампа накаливания — все работают"],
  ["Шкафчики для хранения", "заполнены проспектами, колерными картами, расходниками"],
  ["Знак «Колеровка»", "размещён над зоной, виден с входа"],
  ["Требования к температуре", "в помещении не ниже +15°C (контроль по термометру)"],
];

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TB,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tCellHeader("✓", 7),
      tCellHeader("Элемент", 38),
      tCellHeader("Что проверить", 55),
    ]}),
    ...checklist.map((row, idx) => {
      const fill = idx % 2 === 1 ? "F2F7FC" : null;
      return new TableRow({ cantSplit: true, children: [
        tCell("☐", 7, { align: AlignmentType.CENTER, size: 28, fill }),
        tCell(row[0], 38, { bold: true, fill }),
        tCell(row[1], 55, { fill }),
      ]});
    }),
  ],
}));

// === РАЗДЕЛ 14. ТРЕБОВАНИЯ К ПОМЕЩЕНИЮ ===
children.push(h1("12. Требования к помещению"));

children.push(p(
  "Перед доставкой колеровочного оборудования дистрибьютор обязан подготовить помещение. Требования ниже — обязательные; их невыполнение может привести к некорректной работе оборудования, поломке или отказу в гарантийном обслуживании."
));

children.push(h2("12.1. Пол"));

children.push(p([
  "Пол в месте установки дозатора COROB D180 и миксера MultiSpeed CABO должен быть ",
  ["ровным", { bold: true, color: P.accent }],
  ". Площадка — не менее 1 м². Гироскопический миксер MultiSpeed CABO не создаёт вибрационной нагрузки (в отличие от вибрационных шейкеров), но требует ровной поверхности для устойчивости оборудования. Бетонная основа рекомендуется для долговременной установки, но не является строгим требованием."
]));

children.push(p(
  "Если пол в магазине деревянный или кафельный, перед установкой необходимо подготовить бетонную площадку 1×1 м или установить металлическую раму-основание."
));

children.push(h2("12.2. Электропитание"));

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TB,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tCellHeader("Параметр", 40),
      tCellHeader("Требование", 35),
      tCellHeader("Зачем", 25),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Сечение провода", 40, { bold: true, fill: "F2F7FC" }),
      tCell("Не менее 2,5 мм²", 35, { bold: true, color: P.accent, align: AlignmentType.CENTER, fill: "F2F7FC" }),
      tCell("Для стабильной работы насосов и мотора без просадок напряжения", 25, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Заземление", 40, { bold: true }),
      tCell("Обязательно", 35, { bold: true, color: P.accent, align: AlignmentType.CENTER }),
      tCell("Безопасность оператора и защита электроники дозатора", 25),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Розетки", 40, { bold: true, fill: "F2F7FC" }),
      tCell("Минимум 2 евро-розетки", 35, { bold: true, align: AlignmentType.CENTER, fill: "F2F7FC" }),
      tCell("Одна — для дозатора, вторая — для миксера (можно ещё одну для компьютера)", 25, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Напряжение", 40, { bold: true }),
      tCell("230 В, 50 Гц", 35, { align: AlignmentType.CENTER }),
      tCell("Стандартное бытовое напряжение в РФ", 25),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Бесперебойность", 40, { bold: true, fill: "F2F7FC" }),
      tCell("Стабильное, без скачков", 35, { fill: "F2F7FC" }),
      tCell("Скачки напряжения могут привести к сбою ПО и некорректной дозировке", 25, { fill: "F2F7FC" }),
    ]}),
  ],
}));

children.push(h2("12.3. Температура"));

children.push(infobox("⚠ Температура в помещении: не ниже +15°C", [
  { text: "Если температура опустится ниже +15°C, колоранты MONICOLOR PRO начнут густеть. Это приведёт к:", bold: true, color: P.accent },
  { text: "• некорректной дозировке (машина не сможет точно отмерить нужный объём колоранта)" },
  { text: "• засорению форсунок и сопел" },
  { text: "• повышенному износу насосов bellows pump" },
  { text: "• сбою в работе программного обеспечения" },
  { text: "В зимний период обязательно отапливать помещение. Контролировать температуру по термометру, размещённому рядом с машиной.", bold: true, color: P.accent },
], { fill: "FFF2CC", titleColor: P.warn }));

children.push(h2("12.4. Площадь"));

children.push(p([
  "Минимальная площадь зоны колеровки — ",
  ["1,5 м²", { bold: true, color: P.accent }],
  " (для маленького DIY-магазина). Включает: место под дозатор COROB D180 (≈0,6×0,8 м), место под миксер MultiSpeed CABO (≈0,5×0,7 м), место для оператора (0,5 м перед машинами). Для средних и крупных магазинов рекомендуется 2–6 м² — с учётом стенда, компьютера и принтера этикеток."
]));

// === РАЗДЕЛ 15. ЭТАПЫ УСТАНОВКИ ===
children.push(h1("13. Этапы установки и пуско-наладки"));

children.push(p(
  "Установка колеровочного оборудования проводится в 8 этапов. Первые 4 этапа — подготовка помещения и распаковка, последние 4 — установка, активация, подготовка колорантов и первая колеровка. Этапы 5–8 выполняются специалистами Эмпилс при пуско-наладке; этапы 1–4 — ответственность дистрибьютора."
));

const stages = [
  ["1", "Подготовка помещения", "Дистрибьютор", "Бетонный пол 1×1 м, 2 евро-розетки с заземлением, температура +15°C и выше. Проверка соответствия требованиям раздела 12."],
  ["2", "Распаковка инструмента и составляющих", "Дистрибьютор", "Распаковать оборудование и коробки, идущие с ним. Достать шестигранник и комплектующие. Проверить комплектность по упаковочному листу."],
  ["3", "Удаление транспортировочных крепежей", "Дистрибьютор / Специалист", "На миксере и дозаторе удалить транспортировочные крепежи: штыри на миксере снизу (перекусив стяжки бокорезами), крепежные элементы на поворотном столе дозатора с двух сторон."],
  ["4", "Установка монитора, мыши и кабелей", "Специалист Эмпилс", "Установить держатель монитора на внешней стороне дозатора. Подключить монитор (HDMI к системному блоку и к бесперебойнику). Соединить дозатор с системным блоком синим кабелем. Все провода зафиксировать зажимами для безопасного движения подвижного стола."],
  ["5", "Активация программы Innovatint", "Специалист Эмпилс", "Загрузить Innovatint, ввести серийный номер дозатора (расположен справа на корпусе). Связаться с техником для активации программы (онлайн через AnyDesk). Получить код авторизации, активировать."],
  ["6", "Подготовка колорантов к работе", "Специалист Эмпилс", "Открыть в Innovatint список колорантов с номерами канистр. Строго по списку расклеить самоклеящиеся наклейки на канистры и крышки (наклейка не должна закрывать отверстие на крышке). Перемешать колоранты в миксере (1 минута для обычных, 3 минуты для вязких). Залить колоранты в соответствующие канистры. При первичной заливке — прокачать насос bellows pump 5–6 раз для удаления воздуха. Залить воду в очистительный лоток. Указать в программе количество залитого колоранта — 950 мл."],
  ["7", "Подготовка сопел к корректному дозированию", "Специалист Эмпилс", "Провести первичную продувку всех 16 канистр без исключения. Дозировать вручную по 3 мл каждого колоранта, проверить качество: ровная струя без пузырей. Повторить для 2 мл и 1 мл. Если дозирование неровное — процедуру повторить. Внести цену за 1 л каждого колоранта в скрытой строке параметров."],
  ["8", "Первая колеровка", "Специалист Эмпилс + Оператор", "Провести первую колеровку по выбранному цвету из каталога RAL. Проверить результат: цвет должен соответствовать эталону. Распечатать этикетку. Передать оборудование в эксплуатацию, провести обучение оператора."],
];

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TB,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tCellHeader("№", 5),
      tCellHeader("Этап", 25),
      tCellHeader("Ответственный", 20),
      tCellHeader("Описание", 50),
    ]}),
    ...stages.map((row, idx) => {
      const fill = idx % 2 === 1 ? "F2F7FC" : null;
      return new TableRow({ cantSplit: true, children: [
        tCell(row[0], 5, { bold: true, align: AlignmentType.CENTER, size: 26, color: P.primary, fill }),
        tCell(row[1], 25, { bold: true, fill }),
        tCell(row[2], 20, { align: AlignmentType.CENTER, fill }),
        tCell(row[3], 50, { fill }),
      ]});
    }),
  ],
}));

children.push(caption("Пуско-наладка проводится специалистами Эмпилс бесплатно. Обучение оператора — в день пуско-наладки."));

// === РАЗДЕЛ 16. УСЛОВИЯ ПРЕДОСТАВЛЕНИЯ ===
children.push(h1("14. Условия предоставления оборудования"));

children.push(p([
  "Эмпилс предоставляет колеровочное оборудование (дозатор COROB D180 и миксер MultiSpeed CABO) ",
  ["дистрибьюторам по договору аренды", { bold: true, color: P.primary }],
  ". Условия аренды разработаны так, чтобы оборудование окупалось через объёмы колеровки и было доступно даже небольшим торговым точкам."
]));

children.push(h2("14.1. Базовые условия"));

children.push(infobox("Условная плата: 700 руб/мес", [
  { text: "Эмпилс предоставляет колеровочное оборудование в аренду за условную плату 700 рублей в месяц.", bold: true, color: P.accent },
  { text: "В стоимость аренды входят: дозатор COROB D180, миксер MultiSpeed CABO, рекламная стойка NEWTONE, цветовые веера RAL K7 и NCS index 1950, установка, пуско-наладка, обучение оператора." },
  { text: "Колоранты MONICOLOR PRO (16 шт.) — закупаются дистрибьютором по прайсу Эмпилс. ПО Innovatint — включено в аренду (лицензия предоставляется Эмпилс)." },
], { fill: "F2F7FC" }));

children.push(h2("14.2. Обязательные условия для дистрибьютора"));

children.push(p([
  "Для получения оборудования дистрибьютор обязан выполнить следующие условия:"
]));

children.push(nli([
  new TextRun({ text: "Наличие в ассортименте ", bold: true, color: c(P.primary), font: FONT }),
  new TextRun({ text: "всего ассортимента базовых продуктов (колеруемой продукции) — см. раздел 4. Минимум: AvantGarde, NEWTONE, ОРЕОЛ.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(nli([
  new TextRun({ text: "Выход во втором сезоне ", bold: true, color: c(P.primary), font: FONT }),
  new TextRun({ text: "на объёмы колеровки не менее 24 тонн/год (2 тонны/месяц) с одной машины. Рекомендуемые объёмы по брендам: NEWTONE — 1 500 кг/мес, AvantGarde — 2 000 кг/мес, ОРЕОЛ — не менее 500 кг/мес.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(nli([
  new TextRun({ text: "Обязательное наличие ", bold: true, color: c(P.primary), font: FONT }),
  new TextRun({ text: "торговых марок «AvantGarde», «NewTone» и «ОРЕОЛ» в ассортименте.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(nli([
  new TextRun({ text: "Подготовка помещения ", bold: true, color: c(P.primary), font: FONT }),
  new TextRun({ text: "по требованиям раздела 12: ровный бетонный пол 1×1 м, 2 евро-розетки с заземлением, температура не ниже +15°C.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(nli([
  new TextRun({ text: "Назначение обученного оператора ", bold: true, color: c(P.primary), font: FONT }),
  new TextRun({ text: "— сотрудника, который пройдёт обучение работе с оборудованием (проводится специалистами Эмпилс в день пуско-наладки). Не допускается работа на оборудовании лиц, не прошедших обучение.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(h2("14.3. Что входит в комплект"));

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TB,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tCellHeader("№", 7),
      tCellHeader("Элемент", 35),
      tCellHeader("Источник", 25),
      tCellHeader("Стоимость", 33),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("1", 7, { bold: true, align: AlignmentType.CENTER, fill: "F2F7FC" }),
      tCell("Дозатор COROB D180", 35, { bold: true, fill: "F2F7FC" }),
      tCell("Эмпилс (в аренду)", 25, { align: AlignmentType.CENTER, fill: "F2F7FC" }),
      tCell("700 руб/мес (условная плата)", 33, { align: AlignmentType.CENTER, color: P.accent, bold: true, fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("2", 7, { bold: true, align: AlignmentType.CENTER }),
      tCell("Миксер MultiSpeed CABO", 35, { bold: true }),
      tCell("Эмпилс (в аренду)", 25, { align: AlignmentType.CENTER }),
      tCell("Включено в аренду", 33, { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("3", 7, { bold: true, align: AlignmentType.CENTER, fill: "F2F7FC" }),
      tCell("Рекламная стойка NEWTONE", 35, { bold: true, fill: "F2F7FC" }),
      tCell("Эмпилс (в аренду)", 25, { align: AlignmentType.CENTER, fill: "F2F7FC" }),
      tCell("Включено в аренду", 33, { align: AlignmentType.CENTER, fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("4", 7, { bold: true, align: AlignmentType.CENTER }),
      tCell("Веер RAL K7", 35, { bold: true }),
      tCell("Эмпилс (в аренду)", 25, { align: AlignmentType.CENTER }),
      tCell("Включено в аренду", 33, { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("5", 7, { bold: true, align: AlignmentType.CENTER, fill: "F2F7FC" }),
      tCell("Веер NCS index 1950", 35, { bold: true, fill: "F2F7FC" }),
      tCell("Эмпилс (в аренду)", 25, { align: AlignmentType.CENTER, fill: "F2F7FC" }),
      tCell("Включено в аренду", 33, { align: AlignmentType.CENTER, fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("6", 7, { bold: true, align: AlignmentType.CENTER }),
      tCell("Колоранты MONICOLOR PRO (16 шт.)", 35, { bold: true }),
      tCell("Закупка дистрибьютора", 25, { align: AlignmentType.CENTER }),
      tCell("По прайсу Эмпилс", 33, { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("7", 7, { bold: true, align: AlignmentType.CENTER, fill: "F2F7FC" }),
      tCell("ПО Innovatint 3/4", 35, { bold: true, fill: "F2F7FC" }),
      tCell("Эмпилс (лицензия)", 25, { align: AlignmentType.CENTER, fill: "F2F7FC" }),
      tCell("Включено в аренду", 33, { align: AlignmentType.CENTER, fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("8", 7, { bold: true, align: AlignmentType.CENTER }),
      tCell("Пуско-наладка и обучение", 35, { bold: true }),
      tCell("Эмпилс (бесплатно)", 25, { align: AlignmentType.CENTER }),
      tCell("Бесплатно", 33, { align: AlignmentType.CENTER, color: P.positive, bold: true }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("9", 7, { bold: true, align: AlignmentType.CENTER, fill: "F2F7FC" }),
      tCell("Сервисное обслуживание", 35, { bold: true, fill: "F2F7FC" }),
      tCell("Эмпилс", 25, { align: AlignmentType.CENTER, fill: "F2F7FC" }),
      tCell("За счёт Эмпилс (при соблюдении регламента)", 33, { align: AlignmentType.CENTER, color: P.positive, bold: true, fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("10", 7, { bold: true, align: AlignmentType.CENTER }),
      tCell("Компьютер + принтер этикеток", 35, { bold: true }),
      tCell("Закупка дистрибьютора", 25, { align: AlignmentType.CENTER }),
      tCell("За счёт дистрибьютора", 33, { align: AlignmentType.CENTER }),
    ]}),
  ],
}));

// === РАЗДЕЛ 17. ЭКСПЛУАТАЦИЯ ===
children.push(h1("15. Эксплуатация и техобслуживание"));

children.push(p(
  "После пуско-наладки оборудование передаётся дистрибьютору в эксплуатацию. Дистрибьютор обязан соблюдать регламент ежедневного обслуживания и не допускать к работе на оборудовании лиц, не прошедших обучение. Несоблюдение регламента может привести к поломке оборудования и потере гарантии."
));

children.push(h2("15.1. Ежедневные процедуры"));

children.push(p([
  "Эти процедуры оператор выполняет ",
  ["в начале и в конце каждого рабочего дня", { bold: true, color: P.accent }],
  ":"
]));

children.push(h3("В начале дня:"));

children.push(nli("Включить дозатор COROB D180 и миксер MultiSpeed CABO."));
children.push(nli("Запустить ПО Innovatint, проверить соединение с дозатором."));
children.push(nli([
  new TextRun({ text: "Провести промывку и продувку дозатора ", size: 22, color: c(P.body), font: FONT }),
  new TextRun({ text: "(в соответствии с руководством по эксплуатации).", size: 22, color: c(P.body), font: FONT, bold: true }),
]));
children.push(nli("Проверить наличие воды в чашке очистительного лотка. При необходимости — долить."));
children.push(nli([
  new TextRun({ text: "Проверить уровень колорантов ", size: 22, color: c(P.body), font: FONT }),
  new TextRun({ text: "в каждой из 16 канистр. Уровень должен быть не ниже 900 мл.", size: 22, color: c(P.body), font: FONT, bold: true }),
]));
children.push(nli("Проверить температуру в помещении — не ниже +15°C."));

children.push(h3("В конце дня:"));

children.push(nli("Провести промывку и продувку дозатора (обязательно)."));
children.push(nli("Протереть наружные поверхности оборудования чистой салфеткой."));
children.push(nli("Закрыть крышки канистр с колорантами (предотвращает высыхание)."));
children.push(nli("Выключить дозатор, миксер и компьютер."));
children.push(nli("Закрыть зону колеровки (если есть возможность)."));

children.push(h2("15.2. Что КАТЕГОРИЧЕСКИ НЕЛЬЗЯ делать"));

children.push(infobox("⚠ Запреты при эксплуатации", [
  { text: "Не допускается перемешивание колоранта миксером — только вручную, в течение 1 минуты.", bold: true, color: P.accent },
  { text: "Не допускать к работе на оборудовании людей, не прошедших обучение." },
  { text: "Не использовать оборудование при температуре в помещении ниже +15°C." },
  { text: "Не оставлять оборудование без присмотра во время работы." },
  { text: "Не пытаться самостоятельно ремонтировать оборудование — только специалисты Эмпилс." },
  { text: "Не использовать колоранты других производителей — только MONICOLOR PRO." },
  { text: "Не допускать падения уровня колоранта ниже 900 мл." },
], { fill: "FFF2CC", titleColor: P.warn }));

children.push(h2("15.3. Что делать при сбое"));

children.push(p([
  "Если оборудование вышло из строя или в его работе произошёл сбой, оператор обязан:"
]));

children.push(nli([
  new TextRun({ text: "Немедленно прекратить эксплуатацию ", bold: true, color: c(P.accent), font: FONT }),
  new TextRun({ text: "оборудования.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(nli([
  new TextRun({ text: "В течение одного рабочего дня ", bold: true, color: c(P.accent), font: FONT }),
  new TextRun({ text: "уведомить Арендодателя (Эмпилс) о нарушении в работе оборудования.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(nli("Зафиксировать характер сбоя (что произошло, при какой операции, код ошибки если есть)."));

children.push(nli("Не пытаться устранить неисправность самостоятельно — дождаться специалиста Эмпилс."));

children.push(p([
  "Ремонт оборудования проводится ",
  ["силами и за счёт Эмпилс", { bold: true, color: P.positive }],
  " — при условии соблюдения дистрибьютором регламента эксплуатации. Если поломка произошла вследствие несоблюдения требований руководства по эксплуатации или действий сотрудников дистрибьютора, ремонт проводится за счёт дистрибьютора."
]));

children.push(h2("15.4. Регламентное техническое обслуживание"));

children.push(p(
  "Помимо ежедневных процедур, проводится регламентное техническое обслуживание оборудования специалистами Эмпилс. Периодичность и состав работ определяются руководством по эксплуатации COROB D180 и MultiSpeed CABO. Базовая периодичность — раз в 6 месяцев."
));

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TB,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tCellHeader("Периодичность", 25),
      tCellHeader("Работы", 50),
      tCellHeader("Ответственный", 25),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Ежедневно", 25, { bold: true, align: AlignmentType.CENTER, fill: "F2F7FC" }),
      tCell("Промывка/продувка дозатора, контроль уровня колорантов, проверка воды в чашке", 50, { fill: "F2F7FC" }),
      tCell("Оператор", 25, { align: AlignmentType.CENTER, fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Раз в месяц", 25, { bold: true, align: AlignmentType.CENTER }),
      tCell("Глубокая очистка форсунок, проверка калибровки дозатора", 50),
      tCell("Оператор + специалист", 25, { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Раз в 6 месяцев", 25, { bold: true, align: AlignmentType.CENTER, fill: "F2F7FC" }),
      tCell("Регламентное ТО: замена изношенных деталей, проверка электроники, обновление ПО", 50, { fill: "F2F7FC" }),
      tCell("Специалист Эмпилс", 25, { align: AlignmentType.CENTER, fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("По необходимости", 25, { bold: true, align: AlignmentType.CENTER }),
      tCell("Замена колорантов при истощении запаса, обновление базы рецептур Innovatint", 50),
      tCell("Оператор + специалист", 25, { align: AlignmentType.CENTER }),
    ]}),
  ],
}));

// === РАЗДЕЛ 18. КОНТАКТЫ ===
children.push(h1("16. Сервисная поддержка и контакты"));

children.push(p(
  "По всем вопросам, связанным с колеровочным оборудованием, колорантами, программным обеспечением и сервисным обслуживанием, обращайтесь к специалистам Эмпилс. Мы обеспечиваем полный цикл поддержки — от подбора оборудования до ремонта."
));

children.push(h2("16.1. Специалисты по колеровочным системам"));

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TB,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tCellHeader("Специалист", 30),
      tCellHeader("Должность", 30),
      tCellHeader("Контакты", 40),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Козловский Станислав", 30, { bold: true, color: P.primary, fill: "F2F7FC" }),
      tCellMulti([
        "Главный специалист",
        "по колеровочным системам",
        "ЗАО «ЭМПИЛС»",
      ], 30, { fill: "F2F7FC" }),
      tCellMulti([
        "Тел./факс: (863) 278-68-06, 278-65-81, доб. 4222",
        "Моб.: +7 (988) 556 07 86",
        "E-mail: kozlovskiy@empils.ru",
      ], 40, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Каретин Алексей", 30, { bold: true, color: P.primary }),
      tCellMulti([
        "Специалист",
        "по колеровочным системам",
        "ЗАО «ЭМПИЛС»",
      ], 30),
      tCellMulti([
        "Тел./факс: (863) 278-68-06, 278-65-81, доб. 4142",
        "Моб.: +7 (988) 589 82 35",
        "E-mail: karetin@empils.ru",
      ], 40),
    ]}),
  ],
}));

children.push(h2("16.2. Горячая линия"));

children.push(infobox("Телефон горячей линии: 8 (800) 200-49-11", [
  { text: "Звонок бесплатный по всей России.", bold: true },
  { text: "Режим работы: понедельник–пятница, 9:00–18:00 (по московскому времени)." },
  { text: "По телефону горячей линии можно:" },
  { text: "• оставить заявку на установку оборудования" },
  { text: "• сообщить о сбое в работе оборудования" },
  { text: "• заказать колоранты MONICOLOR PRO" },
  { text: "• получить консультацию по колеровке" },
  { text: "• запросить выезд специалиста" },
], { fill: "F2F7FC" }));

children.push(h2("16.3. Реквизиты"));

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TB,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tCellHeader("Параметр", 30),
      tCellHeader("Значение", 70),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Полное наименование", 30, { bold: true, fill: "F2F7FC" }),
      tCell("ЗАО «ЭМПИЛС»", 70, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Адрес производства", 30, { bold: true }),
      tCellMulti([
        "344090, г. Ростов-на-Дону,",
        "1-й Машиностроительный пер., 21",
      ], 70),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Телефон", 30, { bold: true, fill: "F2F7FC" }),
      tCell("(863) 278-68-06, 278-65-81", 70, { fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Горячая линия", 30, { bold: true }),
      tCell("8 (800) 200-49-11 (бесплатно по РФ)", 70),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("Сайт", 30, { bold: true, fill: "F2F7FC" }),
      tCell("www.empils.ru", 70, { color: P.primary, fill: "F2F7FC" }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tCell("E-mail (общий)", 30, { bold: true }),
      tCell("info@empils.ru", 70, { color: P.primary }),
    ]}),
  ],
}));

children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 600, after: 200 },
  border: {
    top: { style: BorderStyle.SINGLE, size: 12, color: c(P.primary) },
    bottom: { style: BorderStyle.SINGLE, size: 12, color: c(P.primary) },
  },
  children: [new TextRun({
    text: "ЗАО «ЭМПИЛС» — КОЛЕРОВОЧНОЕ ОБОРУДОВАНИЕ ДЛЯ ДИСТРИБЬЮТОРОВ",
    bold: true, size: 22, color: c(P.primary), font: FONT,
  })],
}));

children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 200 },
  children: [new TextRun({
    text: "Издание 2026 · Версия 1.0",
    italics: true, size: 18, color: c(P.muted), font: FONT,
  })],
}));

// === ДОКУМЕНТ ===
const doc = new Document({
  creator: "Z.ai",
  title: "Буклет Эмпилс — Колеровочное оборудование для дистрибьюторов",
  description: "Подробное практическое руководство по запуску и эксплуатации зоны колеровки Эмпилс",
  styles: {
    default: {
      document: {
        run: { font: { ascii: "Calibri", eastAsia: "Calibri" }, size: 22, color: c(P.body) },
        paragraph: { spacing: { line: 312 } },
      },
    },
  },
  numbering: {
    config: [{
      reference: "main-num",
      levels: [{
        level: 0,
        format: LevelFormat.DECIMAL,
        text: "%1.",
        alignment: AlignmentType.START,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } },
      }],
    }],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, bottom: 1440, left: 1701, right: 1417 },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { line: 240 },
          children: [new TextRun({
            text: "Эмпилс — Колеровочное оборудование для дистрибьюторов",
            size: 18, color: c(P.muted), italics: true, font: FONT,
          })],
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { line: 240 },
          children: [
            new TextRun({ text: "Стр. ", size: 18, color: c(P.muted), font: FONT }),
            new TextRun({ children: [PageNumber.CURRENT], size: 18, color: c(P.muted), font: FONT }),
            new TextRun({ text: " из ", size: 18, color: c(P.muted), font: FONT }),
            new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, color: c(P.muted), font: FONT }),
            new TextRun({ text: " · 8 (800) 200-49-11 · www.empils.ru", size: 18, color: c(P.muted), font: FONT }),
          ],
        })],
      }),
    },
    children: children,
  }],
});

Packer.toBuffer(doc).then(buffer => {
  const outPath = "/home/z/my-project/download/Буклет_Эмпилс_Колеровочное_оборудование.docx";
  fs.writeFileSync(outPath, buffer);
  console.log(`✓ Буклет сохранён: ${outPath}`);
  console.log(`  Размер: ${buffer.length.toLocaleString()} байт`);
});
