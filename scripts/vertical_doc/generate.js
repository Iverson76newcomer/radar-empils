// Справка по дистрибьютору «Вертикаль»
// Бизнес-аналитический стиль, 8-12 страниц

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, PageNumber, AlignmentType, HeadingLevel,
  WidthType, BorderStyle, ShadingType, PageBreak, LevelFormat,
} = require("docx");
const fs = require("fs");

// =================== ПАЛИТРА ===================
const P = {
  primary: "#1F4E79",       // тёмно-синий — для заголовков
  accent: "#C00000",         // тёмно-красный — акценты
  body: "#1A1A1A",           // основной текст
  secondary: "#595959",      // подписи
  muted: "#808080",          // колонтитулы
  tableHeader: "#1F4E79",   // шапка таблицы
  tableRowAlt: "#F2F2F2",   // альтернативная строка
  positive: "#548235",       // зелёный — позитивные маркеры
  warning: "#BF8F00",       // жёлто-оранжевый — предупреждения
  negative: "#C00000",       // красный — риски
};
const c = (hex) => hex.replace("#", "");

// =================== ШРИФТЫ ===================
const FONT = { ascii: "Calibri", eastAsia: "Calibri" };
const FONT_HEAD = { ascii: "Calibri", eastAsia: "Calibri" };

// =================== УТИЛИТЫ ===================
function safeText(value, placeholder) {
  if (value === undefined || value === null || value === "" ||
      String(value) === "NaN" || String(value) === "undefined") {
    return placeholder || "—";
  }
  return String(value);
}

// Заголовок H1
function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 180, line: 312 },
    children: [new TextRun({
      text, bold: true, size: 32, color: c(P.primary), font: FONT_HEAD,
    })],
  });
}

// Заголовок H2
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 140, line: 312 },
    children: [new TextRun({
      text, bold: true, size: 28, color: c(P.primary), font: FONT_HEAD,
    })],
  });
}

// Заголовок H3
function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 220, after: 120, line: 312 },
    children: [new TextRun({
      text, bold: true, size: 24, color: c(P.primary), font: FONT_HEAD,
    })],
  });
}

// Параграф основного текста
function p(text, opts = {}) {
  let runs;
  if (Array.isArray(text)) {
    // Массив элементов: строка ИЛИ [строка, {bold, color, ...}] ИЛИ TextRun
    runs = text.map(part => {
      if (typeof part === "string") {
        return new TextRun({ text: part, size: 22, color: c(P.body), font: FONT });
      }
      if (Array.isArray(part)) {
        const [t, props = {}] = part;
        return new TextRun({
          text: t, size: props.size || 22,
          color: c(props.color || P.body),
          bold: !!props.bold,
          italics: !!props.italic,
          font: FONT,
        });
      }
      // Уже a TextRun
      return part;
    });
  } else {
    runs = [new TextRun({
      text, size: 22, color: c(P.body), font: FONT,
    })];
  }
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 312, after: 120 },
    indent: { firstLine: 420 },
    ...opts,
    children: runs,
  });
}

// Параграф с inline-форматированием: массив пар [text, {bold, color, ...}]
function pRich(parts, opts = {}) {
  const runs = parts.map(part => {
    if (typeof part === "string") {
      return new TextRun({ text: part, size: 22, color: c(P.body), font: FONT });
    }
    const [text, props = {}] = Array.isArray(part) ? part : [part.text, part];
    return new TextRun({
      text, size: props.size || 22,
      color: c(props.color || P.body),
      bold: !!props.bold,
      italics: !!props.italic,
      font: FONT,
    });
  });
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 312, after: 120 },
    indent: { firstLine: 420 },
    ...opts,
    children: runs,
  });
}

// Пункт списка (без отступа первой строки)
function li(text, opts = {}) {
  let runs;
  if (Array.isArray(text)) {
    runs = text.map(part => {
      if (typeof part === "string") {
        return new TextRun({ text: part, size: 22, color: c(P.body), font: FONT });
      }
      if (Array.isArray(part)) {
        const [t, props = {}] = part;
        return new TextRun({
          text: t, size: props.size || 22,
          color: c(props.color || P.body),
          bold: !!props.bold,
          italics: !!props.italic,
          font: FONT,
        });
      }
      return part;
    });
  } else {
    runs = [new TextRun({ text, size: 22, color: c(P.body), font: FONT })];
  }
  return new Paragraph({
    spacing: { line: 312, after: 80 },
    bullet: { level: 0 },
    ...opts,
    children: runs,
  });
}

// Пункт нумерованного списка
function nli(text, opts = {}) {
  let runs;
  if (Array.isArray(text)) {
    runs = text.map(part => {
      if (typeof part === "string") {
        return new TextRun({ text: part, size: 22, color: c(P.body), font: FONT });
      }
      return part;
    });
  } else {
    runs = [new TextRun({ text, size: 22, color: c(P.body), font: FONT })];
  }
  return new Paragraph({
    spacing: { line: 312, after: 80 },
    numbering: { reference: "main-numbering", level: 0 },
    ...opts,
    children: runs,
  });
}

// Буллет с bold-меткой: "Метка — значение"
function liKV(label, value, opts = {}) {
  return li([
    new TextRun({ text: label, size: 22, bold: true, color: c(P.primary), font: FONT }),
    new TextRun({ text: " — ", size: 22, color: c(P.body), font: FONT }),
    new TextRun({ text: value, size: 22, color: c(P.body), font: FONT }),
  ], opts);
}

// Поле-сепаратор (горизонтальная линия)
function hr() {
  return new Paragraph({
    spacing: { before: 120, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: c(P.primary) } },
    children: [new TextRun({ text: "", size: 2 })],
  });
}

// Заголовок таблицы
function tableCellHeader(text, widthPct) {
  return new TableCell({
    width: { size: widthPct, type: WidthType.PERCENTAGE },
    shading: { type: ShadingType.CLEAR, fill: c(P.tableHeader) },
    margins: { top: 100, bottom: 100, left: 140, right: 140 },
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { line: 280 },
      children: [new TextRun({
        text, bold: true, size: 22, color: "FFFFFF", font: FONT_HEAD,
      })],
    })],
  });
}

// Ячейка таблицы
function tableCell(text, widthPct, opts = {}) {
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
  if (fill) {
    cellProps.shading = { type: ShadingType.CLEAR, fill: c(fill) };
  }
  return new TableCell(cellProps);
}

// Ячейка с многострочным содержимым
function tableCellMulti(lines, widthPct, opts = {}) {
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
  if (fill) {
    cellProps.shading = { type: ShadingType.CLEAR, fill: c(fill) };
  }
  return new TableCell(cellProps);
}

// Стандартные границы таблицы
const TABLE_BORDERS = {
  top: { style: BorderStyle.SINGLE, size: 4, color: c(P.primary) },
  bottom: { style: BorderStyle.SINGLE, size: 4, color: c(P.primary) },
  left: { style: BorderStyle.SINGLE, size: 2, color: "BFBFBF" },
  right: { style: BorderStyle.SINGLE, size: 2, color: "BFBFBF" },
  insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "D0D0D0" },
  insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "D0D0D0" },
};

// Заголовок-примечание (мелкая подпись)
function caption(text) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 80, after: 200, line: 280 },
    children: [new TextRun({
      text, italics: true, size: 18, color: c(P.secondary), font: FONT,
    })],
  });
}

// =================== СОДЕРЖАНИЕ ДОКУМЕНТА ===================
const children = [];

// === ВВЕДЕНИЕ (Summary) ===
children.push(new Paragraph({
  alignment: AlignmentType.LEFT,
  spacing: { before: 0, after: 180, line: 280 },
  children: [new TextRun({
    text: "Справка подготовлена для внутреннего использования коммерческой службой Эмпилс.",
    italics: true, size: 20, color: c(P.secondary), font: FONT,
  })],
}));

children.push(new Paragraph({
  alignment: AlignmentType.LEFT,
  spacing: { before: 0, after: 240, line: 280 },
  children: [new TextRun({
    text: "Цель: сбор фактических данных по контрагенту ООО «Вертикаль» для подготовки к повторным переговорам, намеченным на сентябрь 2026 г.",
    italics: true, size: 20, color: c(P.secondary), font: FONT,
  })],
}));

children.push(hr());

// === КРАТКАЯ СПРАВКА (Executive Summary) ===
children.push(h1("1. Краткая справка"));

children.push(p([
  "ООО «Вертикаль» (ИНН 5027161766) — ",
  ["крупнейший российский дистрибьютор товаров для строительства и обустройства дома", { bold: true, color: P.primary }],
  ", работающий с 1996 года (30 лет на рынке). По данным сайта компании, в ассортименте более 35 000 товарных позиций, 7 региональных складов, свыше 900 сотрудников и 20 000 активных покупателей. География присутствия — от Москвы до Урала. По информации из внутренней CRM-карточки, выручка компании за 2025 год составила порядка 12 млрд руб.",
]));

children.push(p([
  "Юридическое лицо зарегистрировано в г. Обнинск Калужской области (",
  ["адрес: ул. Курсатова, 31А, помещ. 7, офис 3", { bold: true }],
  "). Единственный учредитель и генеральный директор — ",
  ["Евдоков Владимир Иванович", { bold: true }],
  " (гражданин РФ, ИНН 772902204680, доля 100%). Уставный капитал — 25 000 руб. Основной ОКВЭД с 10.04.2026: 46.73.4 «Торговля оптовая лакокрасочными материалами».",
]));

children.push(p([
  "Ключевой ЛПР со стороны «Вертикали» для переговоров с Эмпилс — ",
  ["Юрий Григорьевский, руководитель отдела продаж", { bold: true, color: P.accent }],
  " (тел. 8 903 030-29-65, e-mail: grigoryevsky@vertical.ru). Последняя встреча — 01.03.2026; клиент рассмотрел коммерческое предложение, возражение: «на сезон 2026 г. новых контрактов не заключают», договорились о повторных переговорах в сентябре 2026 г.",
]));

children.push(p([
  "Компания работает с брендами ",
  ["Лакра, Текс, Рогнеда (основной ЛКМ-ассортимент), Профилюкс, Новоколор (дополнительный)", { bold: true }],
  ". На сайте vertical.ru в каталоге ЛКМ представлены 6 944 SKU, среди крупнейших по объёму брендов — ABC Farben (708 SKU), Акватекс (261 SKU), Альфа (79 SKU). У компании есть ",
  ["собственная торговая марка V", { bold: true, color: P.accent }],
  " (например, грунт-эмаль 3 в 1 V, клеи для обоев V), а также собственная розничная сеть магазинов «ДомДом».",
]));

children.push(p([
  "Заинтересованность клиента: возможность изготовления ",
  ["дешёвой собственной торговой марки", { bold: true, color: P.accent }],
  " и разведение ассортимента на непересекающиеся наименования с уже работающим на территории дистрибьютором. Это открывает для Эмпилс два чётких направления захода: СТМ-партнёрство и ассортиментная дополняемость.",
]));

// === РАЗДЕЛ 2: ОБЩИЕ СВЕДЕНИЯ О КОМПАНИИ ===
children.push(h1("2. Общие сведения о компании"));

children.push(h2("2.1. Реквизиты юридического лица"));

children.push(p(
  "Сведения приведены по официальной выписке из ЕГРЮЛ, полученной на сервисе ФНС egrul.nalog.ru по состоянию на 21.10.2026."
));

// Таблица реквизитов
children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TABLE_BORDERS,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tableCellHeader("Показатель", 35),
      tableCellHeader("Значение", 65),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Полное наименование", 35, { bold: true }),
      tableCell("Общество с ограниченной ответственностью «Вертикаль»", 65),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Сокращённое наименование", 35, { bold: true, fill: P.tableRowAlt }),
      tableCell("ООО «Вертикаль»", 65, { fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("ОГРН", 35, { bold: true }),
      tableCell("1105027005874", 65),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("ИНН", 35, { bold: true, fill: P.tableRowAlt }),
      tableCell("5027161766", 65, { fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("КПП", 35, { bold: true }),
      tableCell("402501001", 65),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Дата регистрации", 35, { bold: true, fill: P.tableRowAlt }),
      tableCell("04.05.2010 (создание юридического лица)", 65, { fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Регистрирующий орган", 35, { bold: true }),
      tableCell("Управление Федеральной налоговой службы по Калужской области", 65),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Юридический адрес", 35, { bold: true, fill: P.tableRowAlt }),
      tableCellMulti([
        "249032, Калужская область,",
        "г. Обнинск, ул. Курсатова, д. 31А,",
        "помещ. 7, офис 3",
      ], 65, { fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Уставный капитал", 35, { bold: true }),
      tableCell("25 000 руб.", 65),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Основной ОКВЭД", 35, { bold: true, fill: P.tableRowAlt }),
      tableCellMulti([
        "46.73.4 — Торговля оптовая лакокрасочными материалами",
        "(изменён 10.04.2026 — это указывает на недавнюю стратегическую переориентацию на ЛКМ)",
      ], 65, { fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Дополнительные ОКВЭД", 35, { bold: true }),
      tableCell("26 кодов — см. раздел 2.3", 65),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Регистрация в ПФР/ФСС", 35, { bold: true, fill: P.tableRowAlt }),
      tableCellMulti([
        "Рег. № 1018718740 (СФР по Калужской области)",
        "Дата постановки на учёт: 19.06.2018",
      ], 65, { fill: P.tableRowAlt }),
    ]}),
  ],
}));

children.push(caption("Источник: выписка из ЕГРЮЛ, egrul.nalog.ru, 21.10.2026."));

children.push(h2("2.2. Что за компания и история"));

children.push(p(
  "По данным собственного сайта vertical.ru, «Вертикаль» позиционирует себя как крупнейший в России дистрибьютор товаров для строительства и обустройства дома. В 2026 году компания отметила 30-летие: фактическая деятельность началась в 1996 году, тогда как текущее юридическое лицо ООО «Вертикаль» зарегистрировано в 2010 году. Это означает, что бизнес существует как минимум 30 лет, причём в разные периоды он мог быть оформлен через другие юридические лица (например, ИП или ранее действовавшее ООО)."
));

children.push(p([
  "Миссия компании: ",
  ["«создавать уют в вашем доме»", { italic: true }],
  ". Специализация — лакокрасочные и строительные материалы, ручной инструмент, садовый инвентарь, товары для дома, бытовая химия. На старте в 1996 году ассортимент на 70% состоял из ЛКМ, сухих строительных смесей и ручного инструмента; 30% приходилось на сопутствующие товары. По состоянию на 2026 год доля ЛКМ остаётся ключевой — в каталоге 6 944 позиции ЛКМ из 35 000 общего ассортимента (около 20% всех SKU).",
]));

children.push(p("Ключевые этапы развития компании:"));

// Этапы истории — таблица
children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TABLE_BORDERS,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tableCellHeader("Год", 10),
      tableCellHeader("Этап", 30),
      tableCellHeader("Ключевое содержание", 60),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("1996", 10, { bold: true, align: AlignmentType.CENTER }),
      tableCell("Становление", 30, { bold: true }),
      tableCell("Запуск бизнеса. 70% ассортимента — ЛКМ, сухие смеси, ручной инструмент; 30% — сопутствующие товары. Развитие за счёт увеличения доли рынка.", 60),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("2003", 10, { bold: true, align: AlignmentType.CENTER, fill: P.tableRowAlt }),
      tableCell("Успешное развитие", 30, { bold: true, fill: P.tableRowAlt }),
      tableCell("Сформированная клиентская база. Активное расширение через торговых представителей. Бесплатная доставка. Внедрение кредитования и отсрочки платежа. Расширение географии по РФ.", 60, { fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("2008", 10, { bold: true, align: AlignmentType.CENTER }),
      tableCell("Расширение ассортимента", 30, { bold: true }),
      tableCell("Доля хозтоваров доведена до 20%, товаров для дома и сада — до 15%. Открытие новых складов (кросс-док). География — до 1 000 км. Запуск СТМ, ротация ассортимента.", 60),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("2019", 10, { bold: true, align: AlignmentType.CENTER, fill: P.tableRowAlt }),
      tableCell("Инновации и рост", 30, { bold: true, fill: P.tableRowAlt }),
      tableCell("Развитие DIY-сети и e-commerce. Запуск направления «Тендеры». Внедрение CRM. Развитие онлайн-кабинета. Запуск розничной сети магазинов «ДомДом».", 60, { fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("2021", 10, { bold: true, align: AlignmentType.CENTER }),
      tableCell("Расширение направлений", 30, { bold: true }),
      tableCell("Ассортимент расширен до 23 000 SKU. Развитие направлений бытовой техники, химии, посуды, текстиля, мебели. Укрепление портфеля СТМ в ЛКМ, инструменте, товарах для сада.", 60),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("2024", 10, { bold: true, align: AlignmentType.CENTER, fill: P.tableRowAlt }),
      tableCell("Цифровой сервис", 30, { bold: true, fill: P.tableRowAlt }),
      tableCell("Партнёрская конференция. Личным кабинетом пользуются более 5 000 клиентов. Объединение цифрового сервиса с поддержкой менеджеров и торговых представителей.", 60, { fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("2026", 10, { bold: true, align: AlignmentType.CENTER }),
      tableCell("30 лет компании", 30, { bold: true }),
      tableCell("Ассортимент — 35 000 SKU. Прирост выручки: e-commerce +26%, профессиональный канал +21%, направление торговых сетей +11%. 10.04.2026 изменён основной ОКВЭД на «Торговля оптовая лакокрасочными материалами».", 60),
    ]}),
  ],
}));

children.push(caption("Источник: страница «О компании» сайта vertical.ru/company."));

children.push(h2("2.3. Виды экономической деятельности (ОКВЭД)"));

children.push(p(
  "У компании 26 дополнительных кодов ОКВЭД, что отражает диверсифицированную модель бизнеса: оптовая и розничная торговля по широкому спектру товарных групп, логистика и складирование, реклама, исследования рынка, аренда нежилой недвижимости. Это даёт компании гибкость и объясняет, почему в каталоге 35 000 SKU по 22 товарным категориям."
));

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TABLE_BORDERS,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tableCellHeader("Код ОКВЭД", 15),
      tableCellHeader("Наименование вида деятельности", 85),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("46.73.4", 15, { bold: true, align: AlignmentType.CENTER, color: P.accent }),
      tableCell("Торговля оптовая лакокрасочными материалами — ОСНОВНОЙ (с 10.04.2026)", 85, { bold: true, color: P.accent }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("68.20.29", 15, { bold: true, align: AlignmentType.CENTER, fill: P.tableRowAlt }),
      tableCell("Аренда и управление собственным или арендованным прочим нежилым недвижимым имуществом", 85, { fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("46.13.2", 15, { bold: true, align: AlignmentType.CENTER }),
      tableCell("Деятельность агентов по оптовой торговле строительными материалами", 85),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("46.73.6", 15, { bold: true, align: AlignmentType.CENTER, fill: P.tableRowAlt }),
      tableCell("Торговля оптовая прочими строительными материалами и изделиями", 85, { fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("46.74.1", 15, { bold: true, align: AlignmentType.CENTER }),
      tableCell("Торговля оптовая скобяными изделиями", 85),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("46.74.3", 15, { bold: true, align: AlignmentType.CENTER, fill: P.tableRowAlt }),
      tableCell("Торговля оптовая ручными инструментами", 85, { fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("46.90", 15, { bold: true, align: AlignmentType.CENTER }),
      tableCell("Торговля оптовая неспециализированная", 85),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("47.52", 15, { bold: true, align: AlignmentType.CENTER, fill: P.tableRowAlt }),
      tableCell("Торговля розничная скобяными изделиями, лакокрасочными материалами и стеклом в специализированных магазинах", 85, { fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("47.91", 15, { bold: true, align: AlignmentType.CENTER }),
      tableCell("Торговля розничная по почте или по информационно-коммуникационной сети Интернет", 85),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("52.10.9", 15, { bold: true, align: AlignmentType.CENTER, fill: P.tableRowAlt }),
      tableCell("Хранение и складирование прочих грузов", 85, { fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("52.24.2", 15, { bold: true, align: AlignmentType.CENTER }),
      tableCell("Транспортная обработка прочих грузов", 85),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("52.29", 15, { bold: true, align: AlignmentType.CENTER, fill: P.tableRowAlt }),
      tableCell("Деятельность вспомогательная прочая, связанная с перевозками", 85, { fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("73.11", 15, { bold: true, align: AlignmentType.CENTER }),
      tableCell("Деятельность рекламных агентств", 85),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("73.20", 15, { bold: true, align: AlignmentType.CENTER, fill: P.tableRowAlt }),
      tableCell("Исследование конъюнктуры рынка и изучение общественного мнения", 85, { fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("82.92", 15, { bold: true, align: AlignmentType.CENTER }),
      tableCell("Деятельность по упаковыванию товаров", 85),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Прочие", 15, { bold: true, align: AlignmentType.CENTER, fill: P.tableRowAlt }),
      tableCell("Ещё 11 кодов: торговля текстилём, обувью, чистящими средствами, бытовыми приборами, садово-огородной техникой, мебелью и т.д.", 85, { fill: P.tableRowAlt, italic: true }),
    ]}),
  ],
}));

// === РАЗДЕЛ 3: ФИНАНСОВЫЕ ПОКАЗАТЕЛИ ===
children.push(h1("3. Финансовые показатели"));

children.push(h2("3.1. Выручка и динамика"));

children.push(p([
  "По данным внутренней CRM-карточки Эмпилс, выручка ООО «Вертикаль» за 2025 год составила ",
  ["порядка 12,0 млрд руб.", { bold: true, color: P.accent }],
  " Данные за 2024 год в карточке не заполнены, однако по косвенным признакам (рост электронной коммерции +26%, профессионального канала +21%, торговых сетей +11% в 2026 году) компания демонстрирует устойчивый рост выручки. Учитывая 30-летнюю историю, диверсифицированную модель бизнеса и 7 региональных складов, объём выручки 12 млрд руб. вполне согласуется с рыночным положением «крупнейшего российского дистрибьютора строительных товаров».",
]));

children.push(p([
  "Полные финансовые показатели (прибыль, рентабельность, ликвидность, динамика за ряд лет) требуют запроса расширенной выписки или платных источников типа СПАРК/Контур.Фокус. Базовая выписка из ЕГРЮЛ такие данные не содержит. Рекомендуется перед переговорами запросить через партнёров или менеджеров «Вертикали» справку 2-НДФЛ/Ф1 или воспользоваться платными сервисами для углублённого финансового анализа.",
]));

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TABLE_BORDERS,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tableCellHeader("Показатель", 50),
      tableCellHeader("Значение", 50),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Выручка за 2025 г. (по данным CRM)", 50, { bold: true }),
      tableCell("≈ 12,0 млрд руб.", 50, { bold: true, color: P.accent }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Выручка за 2024 г.", 50, { bold: true, fill: P.tableRowAlt }),
      tableCell("не подтверждена (требуется проверка через СПАРК/Контур.Фокус)", 50, { fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Чистая прибыль", 50, { bold: true }),
      tableCell("нет данных в открытых источниках", 50),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Рентабельность продаж", 50, { bold: true, fill: P.tableRowAlt }),
      tableCell("нет данных в открытых источниках", 50, { fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Динамика в 2026 г. (по данным компании)", 50, { bold: true }),
      tableCellMulti([
        "e-commerce: +26%",
        "профессиональный канал: +21%",
        "торговые сети: +11%",
      ], 50, { color: P.positive, bold: true }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Уставный капитал", 50, { bold: true, fill: P.tableRowAlt }),
      tableCell("25 000 руб.", 50, { fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Численность сотрудников", 50, { bold: true }),
      tableCell("≈ 900 человек (по данным сайта)", 50),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Покупатели (клиентская база)", 50, { bold: true, fill: P.tableRowAlt }),
      tableCell("≈ 20 000 покупателей (по данным сайта)", 50, { fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Ассортимент", 50, { bold: true }),
      tableCell("≈ 35 000 SKU (по данным сайта), в т.ч. ЛКМ — 6 944 SKU", 50),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Производительность труда (выручка/сотрудник)", 50, { bold: true, fill: P.tableRowAlt }),
      tableCell("≈ 13,3 млн руб./чел. — высокий показатель для оптовой торговли", 50, { fill: P.tableRowAlt, bold: true, color: P.positive }),
    ]}),
  ],
}));

children.push(caption("Источники: выписка ЕГРЮЛ (egrul.nalog.ru), сайт vertical.ru, внутренняя CRM-карточка Эмпилс."));

children.push(h2("3.2. Структура собственности и руководство"));

children.push(p([
  "Компания имеет ",
  ["упрощённую структуру собственности", { bold: true }],
  ": единственный учредитель и генеральный директор — одно и то же лицо. Это указывает на частный, не институциональный бизнес, где ключевые решения принимаются оперативно. Отсутствие соучредителей снижает риск корпоративных конфликтов, но также означает, что переговоры и решения по существенным вопросам проходят через одного человека."
]));

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TABLE_BORDERS,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tableCellHeader("Параметр", 30),
      tableCellHeader("Значение", 70),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Единственный участник (учредитель)", 30, { bold: true }),
      tableCellMulti([
        "Евдоков Владимир Иванович",
        "ИНН: 772902204680 (регистрация в г. Москве)",
        "Доля: 100% (номинальная стоимость — 25 000 руб.)",
        "Внесён в ЕГРЮЛ: 22.09.2021",
      ], 70),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Генеральный директор", 30, { bold: true, fill: P.tableRowAlt }),
      tableCellMulti([
        "Евдоков Владимир Иванович (тот же человек)",
        "Назначен на должность: 09.11.2021",
        "ИНН: 772902204680",
        "Гражданство: РФ",
      ], 70, { fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Совпадение учредитель = директор", 30, { bold: true }),
      tableCell("Да. Это снижает корпоративные риски, но означает концентрацию решений в одних руках.", 70, { color: P.warning }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Предыдущий представитель (по документам)", 30, { bold: true, fill: P.tableRowAlt }),
      tableCellMulti([
        "Коник А.Ю. (упоминается в доверенности от 03.05.2018)",
        "Возможно, прежний директор или представитель до 2021 года",
      ], 70, { fill: P.tableRowAlt }),
    ]}),
  ],
}));

// === РАЗДЕЛ 4: ИНФРАСТРУКТУРА ===
children.push(h1("4. Инфраструктура: офисы, склады, логистика"));

children.push(h2("4.1. Сеть пунктов самовывоза и складов"));

children.push(p([
  "По данным страницы «Контакты» сайта vertical.ru, у компании ",
  ["8 пунктов самовывоза для юридических лиц", { bold: true, color: P.primary }],
  " в крупных городах России. По данным сайта «О компании», складов — 7 (возможно, один из пунктов самовывоза не является полноценным складом, а выполняет функцию точки выдачи). В файле клиентской карточки также указано: «наличие склада — да, наличие логистики — да», что подтверждается сайтом.",
]));

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TABLE_BORDERS,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tableCellHeader("№", 5),
      tableCellHeader("Город", 15),
      tableCellHeader("Адрес", 60),
      tableCellHeader("Часы работы", 20),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("1", 5, { bold: true, align: AlignmentType.CENTER }),
      tableCell("Москва (Быково)", 15, { bold: true }),
      tableCell("Аэропортовская ул., 14, р.п. Быково, Раменский городской округ", 60),
      tableCell("Пн-Пт 08:00-17:00", 20, { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("2", 5, { bold: true, align: AlignmentType.CENTER, fill: P.tableRowAlt }),
      tableCell("Воронеж", 15, { bold: true, fill: P.tableRowAlt }),
      tableCell("Базовая ул., 9Б, г. Воронеж", 60, { fill: P.tableRowAlt }),
      tableCell("Пн-Пт 08:00-15:00", 20, { align: AlignmentType.CENTER, fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("3", 5, { bold: true, align: AlignmentType.CENTER }),
      tableCell("Вологда", 15, { bold: true }),
      tableCell("Александра Клубова ул., 44А, г. Вологда", 60),
      tableCell("Пн-Пт 08:00-15:00", 20, { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("4", 5, { bold: true, align: AlignmentType.CENTER, fill: P.tableRowAlt }),
      tableCell("Нижний Новгород", 15, { bold: true, fill: P.tableRowAlt }),
      tableCell("Федосеенко ул., 43, г. Нижний Новгород", 60, { fill: P.tableRowAlt }),
      tableCell("Пн-Пт 08:00-15:00", 20, { align: AlignmentType.CENTER, fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("5", 5, { bold: true, align: AlignmentType.CENTER }),
      tableCell("Орёл", 15, { bold: true }),
      tableCell("Силикатный пер., 1, г. Орёл", 60),
      tableCell("Пн-Пт 08:00-15:00", 20, { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("6", 5, { bold: true, align: AlignmentType.CENTER, fill: P.tableRowAlt }),
      tableCell("Рязань", 15, { bold: true, fill: P.tableRowAlt }),
      tableCell("Ряжское ш., д. 20, Рязанская обл.", 60, { fill: P.tableRowAlt }),
      tableCell("Пн-Пт 08:00-15:00", 20, { align: AlignmentType.CENTER, fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("7", 5, { bold: true, align: AlignmentType.CENTER }),
      tableCell("Санкт-Петербург", 15, { bold: true }),
      tableCell("Кубинская ул., 75к1, г. Санкт-Петербург", 60),
      tableCell("Пн-Пт 08:00-17:00", 20, { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("8", 5, { bold: true, align: AlignmentType.CENTER, fill: P.tableRowAlt }),
      tableCell("Иваново", 15, { bold: true, fill: P.tableRowAlt }),
      tableCell("Коляново д., Рябиновая ул., стр. 11д, корп. 2, Ивановская обл.", 60, { fill: P.tableRowAlt }),
      tableCell("Пн-Пт 08:00-15:00", 20, { align: AlignmentType.CENTER, fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("—", 5, { bold: true, align: AlignmentType.CENTER }),
      tableCell("Обнинск (юр. адрес)", 15, { bold: true, italic: true }),
      tableCell("ул. Курсатова, 31А, помещ. 7, офис 3, г. Обнинск, Калужская обл.", 60, { italic: true }),
      tableCell("офис", 20, { align: AlignmentType.CENTER, italic: true }),
    ]}),
  ],
}));

children.push(caption("Источник: страница «Контакты» сайта vertical.ru/contacts."));

children.push(h2("4.2. География и логистика"));

children.push(p([
  "На главной странице сайта компания заявляет: ",
  ["«Территориальный охват: от Москвы до Урала»", { bold: true, color: P.accent }],
  ". Это значительно шире, чем указано во внутренней CRM-карточке («Воронежская, Белгородская»), что говорит о том, что данные в карточке могут относиться к зоне ответственности конкретного регионального менеджера Эмпилс, а не к фактической географии работы дистрибьютора. По данным сайта, бесплатная доставка предоставляется при заказе от 20 000 руб. по Центральному, Северо-Западному и Приволжскому федеральным округам.",
]));

children.push(p([
  "Канал продаж: B2B (только юридические лица). Для физических лиц продукция доступна через маркетплейсы ",
  ["OZON, Яндекс.Маркет, Wildberries", { bold: true }],
  ". Это означает, что «Вертикаль» присутствует и на e-commerce-рынке через marketplace-модель.",
]));

children.push(p([
  "Собственный автопарк и команда торговых представителей — согласно данным из CRM-карточки — присутствуют. На сайте указано: «Заказы 24/7 через личный кабинет» — то есть работает собственный онлайн-кабинет для B2B-клиентов (им пользуется свыше 5 000 клиентов по состоянию на 2024 год).",
]));

// === РАЗДЕЛ 5: БРЕНДЫ И КОНКУРЕНТНЫЙ КОНТЕКСТ ===
children.push(h1("5. Бренды и конкурентный контекст"));

children.push(h2("5.1. Текущий портфель брендов «Вертикали»"));

children.push(p(
  "По данным внутренней CRM-карточки и сайта vertical.ru, портфель ЛКМ-брендов «Вертикали» можно разделить на три блока: основные (из CRM), дополнительные (из CRM) и представленные в каталоге сайта (расширенный список)."
));

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TABLE_BORDERS,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tableCellHeader("Категория", 25),
      tableCellHeader("Бренды", 50),
      tableCellHeader("Комментарий", 25),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Основной ЛКМ-ассортимент (из CRM)", 25, { bold: true }),
      tableCellMulti([
        "• Лакра",
        "• Текс",
        "• Рогнеда",
      ], 50),
      tableCell("Прямые конкуренты Эмпилс по основному ассортименту — это прежде всего Текс и Лакра (конкуренты в сегменте масс-маркет)", 25, { color: P.accent }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Дополнительный ЛКМ-ассортимент (из CRM)", 25, { bold: true, fill: P.tableRowAlt }),
      tableCellMulti([
        "• Профилюкс",
        "• Новоколор",
      ], 50, { fill: P.tableRowAlt }),
      tableCell("Промо-бренды на сайте: ABC Farben, Акватекс, Формула Q8, LINNIMAX", 25, { fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Топ-бренды в каталоге ЛКМ (по кол-ву SKU)", 25, { bold: true }),
      tableCellMulti([
        "• ABC Farben (ЗАО «АВС Фарбен») — 708 SKU",
        "• Акватекс — 261 SKU",
        "• Альфа — 79 SKU",
        "• Арикон — 23 SKU",
        "• Всего в каталоге ЛКМ — 6 944 SKU",
      ], 50),
      tableCell("ABC Farben позиционируется как «крупнейший производитель ЛКМ России и ближнего зарубежья»", 25),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Доп. бренды в каталоге", 25, { bold: true, fill: P.tableRowAlt }),
      tableCellMulti([
        "• Krafor, VGT (ВГТ), Dali, Veres,",
        "  Престиж, Церезит, KUDO,",
        "  Tikkurila, LINNIMAX,",
        "  Формула Q8, Tikkurila",
      ], 50, { fill: P.tableRowAlt }),
      tableCell("Широкий premium- и средний сегмент", 25, { fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Собственная ТМ", 25, { bold: true, color: P.accent }),
      tableCellMulti([
        "• V — грунт-эмаль 3 в 1 V (20 кг, серая),",
        "  клеи для обоев V",
        "  (новинка от 22.01.2026)",
      ], 50, { color: P.accent }),
      tableCell("СТМ компании, запущена в 2008 г., укрепляется с 2021 г.", 25, { bold: true, color: P.accent }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Розничная сеть собственной розницы", 25, { bold: true, fill: P.tableRowAlt }),
      tableCellMulti([
        "• «ДомДом» — сеть магазинов DIY",
        "  (запущена в 2019 г.)",
      ], 50, { fill: P.tableRowAlt }),
      tableCell("Это второй (розничный) канал продаж помимо оптового", 25, { fill: P.tableRowAlt }),
    ]}),
  ],
}));

children.push(h2("5.2. Анализ замещения брендов Эмпилс"));

children.push(p(
  "Сопоставление текущего портфеля брендов «Вертикали» с ассортиментом Эмпилс (ОРЕОЛ TITAN, Ореол, Ореол Дисконт, NEWTONE, AvantGarde, Расцвет, Простокрашено! БАУ, PROFIWOOD) показывает, что прямого конфликта по большинству марок нет — за исключением частичного совпадения по категории «бюджетный ЛКМ». Это даёт Эмпилс возможность предложить «Вертикали» ассортимент, который дополнит, а не заместит уже представленные бренды."
));

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TABLE_BORDERS,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tableCellHeader("Сегмент", 25),
      tableCellHeader("Бренды «Вертикали»", 30),
      tableCellHeader("Возможная замена/дополнение от Эмпилс", 30),
      tableCellHeader("Тип захода", 15),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Премиум-ЛКМ", 25, { bold: true }),
      tableCellMulti(["Tikkurila", "ABC Farben", "LINNIMAX"], 30),
      tableCellMulti(["AvantGarde", "NEWTONE"], 30),
      tableCell("дополнение", 15, { bold: true, color: P.positive, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Средний ЛКМ", 25, { bold: true, fill: P.tableRowAlt }),
      tableCellMulti(["Лакра", "Текс", "Акватекс", "Veres", "Dali", "VGT"], 30, { fill: P.tableRowAlt }),
      tableCellMulti(["Ореол", "ОРЕОЛ TITAN", "Расцвет"], 30, { fill: P.tableRowAlt }),
      tableCell("замещение", 15, { bold: true, color: P.warning, align: AlignmentType.CENTER, fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Эконом-ЛКМ", 25, { bold: true }),
      tableCellMulti(["Рогнеда", "Профилюкс", "Новоколор"], 30),
      tableCellMulti(["Ореол Дисконт", "Простокрашено! БАУ"], 30),
      tableCell("замещение", 15, { bold: true, color: P.warning, align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Защита и декор дерева", 25, { bold: true, fill: P.tableRowAlt }),
      tableCellMulti(["Акватекс", "Veres", "ABC Farben"], 30, { fill: P.tableRowAlt }),
      tableCellMulti(["PROFIWOOD"], 30, { fill: P.tableRowAlt }),
      tableCell("дополнение", 15, { bold: true, color: P.positive, align: AlignmentType.CENTER, fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("СТМ", 25, { bold: true, color: P.accent }),
      tableCellMulti(["V (собств.)"], 30, { color: P.accent }),
      tableCellMulti(["Производство СТМ под брендом клиента — Эмпилс как производитель"], 30, { color: P.accent }),
      tableCell("новое", 15, { bold: true, color: P.accent, align: AlignmentType.CENTER }),
    ]}),
  ],
}));

children.push(p([
  "Важная стратегическая возможность: «Вертикаль» уже имеет собственную ТМ V в категории грунт-эмалей и клеёв, и в истории компании отмечено «укрепление портфеля СТМ в ЛКМ, инструменте и товарах для сада». В CRM-карточке клиента прямо зафиксировано: ",
  ["«возможно изготовление интересно собственных ТМ»", { bold: true, color: P.accent }],
  ". Это означает, что компания открыта к партнёрству по производству СТМ — Эмпилс как завод-производитель может предложить контрактное производство дешёвой СТМ под брендом «Вертикали». Это ключевой аргумент для переговоров.",
]));

// === РАЗДЕЛ 6: КОНТАКТНЫЙ БЛОК ===
children.push(h1("6. Контактный блок"));

children.push(h2("6.1. Контакты ключевых лиц"));

children.push(p([
  "Для переговоров с Эмпилс со стороны «Вертикали» ответственен ",
  ["Юрий Григорьевский — руководитель отдела продаж", { bold: true, color: P.accent }],
  ". Это означает, что в его подчинении находится команда менеджеров по продажам. Прямой выход на генерального директора Евдокова В.И. возможен через Юрия Григорьевского либо через офис компании.",
]));

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TABLE_BORDERS,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tableCellHeader("Роль / Должность", 30),
      tableCellHeader("ФИО", 25),
      tableCellHeader("Контакты", 45),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Генеральный директор / единственный учредитель", 30, { bold: true, color: P.primary }),
      tableCell("Евдоков Владимир Иванович", 25, { bold: true }),
      tableCellMulti([
        "ИНН: 772902204680",
        "Назначен: 09.11.2021",
        "Контакт через офис компании",
      ], 45),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Руководитель отдела продаж (ЛПР)", 30, { bold: true, color: P.accent, fill: P.tableRowAlt }),
      tableCell("Юрий Григорьевский", 25, { bold: true, color: P.accent, fill: P.tableRowAlt }),
      tableCellMulti([
        "Телефон: 8 903 030-29-65 (моб.)",
        "E-mail: grigoryevsky@vertical.ru",
      ], 45, { fill: P.tableRowAlt, color: P.accent }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Онлайн-консультант сайта", 30, { bold: true }),
      tableCell("Устюжанинова Оксана", 25),
      tableCellMulti([
        "Чат на vertical.ru",
        "(только для общих вопросов)",
      ], 45),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Ответственный РД со стороны Эмпилс", 30, { bold: true, fill: P.tableRowAlt }),
      tableCell("Еремеев В.Н.", 25, { fill: P.tableRowAlt }),
      tableCellMulti([
        "Регион: Центральный федеральный округ",
        "Дата последней встречи с клиентом: 01.03.2026",
      ], 45, { fill: P.tableRowAlt }),
    ]}),
  ],
}));

children.push(h2("6.2. Общие контакты компании"));

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TABLE_BORDERS,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tableCellHeader("Канал", 25),
      tableCellHeader("Значение", 75),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Сайт", 25, { bold: true }),
      tableCell("https://vertical.ru", 75, { color: P.primary }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Многоканальный телефон", 25, { bold: true, fill: P.tableRowAlt }),
      tableCell("8 (800) 222-4-000 — бесплатный по РФ", 75, { fill: P.tableRowAlt, bold: true }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Общая почта", 25, { bold: true }),
      tableCell("vopros@vertical.ru", 75, { color: P.primary }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Время работы", 25, { bold: true, fill: P.tableRowAlt }),
      tableCell("Пн-Пт, 08:00-17:00 (по МСК)", 75, { fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Личный кабинет B2B", 25, { bold: true }),
      tableCell("Доступ через сайт vertical.ru — для юридических лиц", 75),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Форма обратной связи", 25, { bold: true, fill: P.tableRowAlt }),
      tableCell("Кнопка «Отправить запрос» в шапке сайта / «Отправьте нам сообщение» в подвале", 75, { fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Маркетплейсы для физлиц", 25, { bold: true }),
      tableCell("OZON, Яндекс.Маркет, Wildberries", 75),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Юридический адрес", 25, { bold: true, fill: P.tableRowAlt }),
      tableCell("249032, Калужская обл., г. Обнинск, ул. Курсатова, 31А, помещ. 7, офис 3", 75, { fill: P.tableRowAlt }),
    ]}),
  ],
}));

children.push(h2("6.3. Пункты самовывоза по городам (детально)"));

children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: TABLE_BORDERS,
  rows: [
    new TableRow({ tableHeader: true, cantSplit: true, children: [
      tableCellHeader("Город", 20),
      tableCellHeader("Адрес", 60),
      tableCellHeader("Часы работы", 20),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Москва", 20, { bold: true }),
      tableCell("Аэропортовская ул., 14, р.п. Быково, Раменский р-н", 60),
      tableCell("Пн-Пт 08:00-17:00", 20, { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Санкт-Петербург", 20, { bold: true, fill: P.tableRowAlt }),
      tableCell("Кубинская ул., 75к1", 60, { fill: P.tableRowAlt }),
      tableCell("Пн-Пт 08:00-17:00", 20, { align: AlignmentType.CENTER, fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Воронеж", 20, { bold: true }),
      tableCell("Базовая ул., 9Б", 60),
      tableCell("Пн-Пт 08:00-15:00", 20, { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Нижний Новгород", 20, { bold: true, fill: P.tableRowAlt }),
      tableCell("Федосеенко ул., 43", 60, { fill: P.tableRowAlt }),
      tableCell("Пн-Пт 08:00-15:00", 20, { align: AlignmentType.CENTER, fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Вологда", 20, { bold: true }),
      tableCell("Александра Клубова ул., 44А", 60),
      tableCell("Пн-Пт 08:00-15:00", 20, { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Орёл", 20, { bold: true, fill: P.tableRowAlt }),
      tableCell("Силикатный пер., 1", 60, { fill: P.tableRowAlt }),
      tableCell("Пн-Пт 08:00-15:00", 20, { align: AlignmentType.CENTER, fill: P.tableRowAlt }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Рязань", 20, { bold: true }),
      tableCell("Ряжское ш., д. 20", 60),
      tableCell("Пн-Пт 08:00-15:00", 20, { align: AlignmentType.CENTER }),
    ]}),
    new TableRow({ cantSplit: true, children: [
      tableCell("Иваново", 20, { bold: true, fill: P.tableRowAlt }),
      tableCell("Коляново д., Рябиновая ул., стр. 11д, корп. 2", 60, { fill: P.tableRowAlt }),
      tableCell("Пн-Пт 08:00-15:00", 20, { align: AlignmentType.CENTER, fill: P.tableRowAlt }),
    ]}),
  ],
}));

// === РАЗДЕЛ 7: ПЕРЕГОВОРНАЯ ПОЗИЦИЯ ===
children.push(h1("7. Переговорная позиция"));

children.push(h2("7.1. Текущий статус переговоров"));

children.push(p(
  "По данным внутренней CRM-карточки Эмпилс, переговоры с «Вертикалью» находятся в следующем состоянии:"
));

children.push(nli([
  new TextRun({ text: "Дата последней встречи: ", size: 22, bold: true, color: c(P.primary), font: FONT }),
  new TextRun({ text: "01.03.2026 г.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(nli([
  new TextRun({ text: "Ответственный РД со стороны Эмпилс: ", size: 22, bold: true, color: c(P.primary), font: FONT }),
  new TextRun({ text: "Еремеев В.Н. (Центральный федеральный округ).", size: 22, color: c(P.body), font: FONT }),
]));

children.push(nli([
  new TextRun({ text: "Статус рассмотрения КП: ", size: 22, bold: true, color: c(P.primary), font: FONT }),
  new TextRun({ text: "клиент рассмотрел коммерческое предложение, ожидает переговоров в сентябре 2026 г.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(nli([
  new TextRun({ text: "Зафиксированное возражение: ", size: 22, bold: true, color: c(P.negative), font: FONT }),
  new TextRun({ text: "«на сезон 2026 г. новых контрактов не заключают».", size: 22, color: c(P.body), font: FONT }),
]));

children.push(nli([
  new TextRun({ text: "Согласованные следующие шаги: ", size: 22, bold: true, color: c(P.primary), font: FONT }),
  new TextRun({ text: "повторные переговоры в сентябре 2026 г.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(p([
  "Зафиксированный инструмент работы с клиентом: ",
  ["разведение ассортимента Эмпилс на непересекающиеся наименования с уже работающим дистрибьютором на территории; акцент на заинтересованности клиента в дешёвой СТМ", { italic: true }],
  ". Эти два направления должны стать основой повестки следующей встречи.",
]));

children.push(h2("7.2. Сильные стороны «Вертикали» как контрагента"));

children.push(li([
  new TextRun({ text: "Масштаб и охват: ", size: 22, bold: true, color: c(P.positive), font: FONT }),
  new TextRun({ text: "30-летний опыт, 7 региональных складов, 900 сотрудников, 20 000 покупателей, география от Москвы до Урала.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(li([
  new TextRun({ text: "Устойчивый рост: ", size: 22, bold: true, color: c(P.positive), font: FONT }),
  new TextRun({ text: "прирост выручки e-commerce +26%, профканала +21%, торговых сетей +11% в юбилейном 2026 г. Компания не стагнирует, а инвестирует в развитие.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(li([
  new TextRun({ text: "B2B-фокус: ", size: 22, bold: true, color: c(P.positive), font: FONT }),
  new TextRun({ text: "сотрудничество только с юрлицами + собственный личный кабинет для заказов 24/7 (5000+ активных клиентов). Это идеальный канал для оптовых поставок Эмпилс.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(li([
  new TextRun({ text: "ЛКМ-специализация: ", size: 22, bold: true, color: c(P.positive), font: FONT }),
  new TextRun({ text: "10.04.2026 компания сменила основной ОКВЭД на «Торговля оптовая лакокрасочными материалами» — это стратегическое подтверждение фокуса на ЛКМ. Эмпилс попадает в ядро их бизнес-модели.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(li([
  new TextRun({ text: "Розничная сеть «ДомДом» и маркетплейсы: ", size: 22, bold: true, color: c(P.positive), font: FONT }),
  new TextRun({ text: "дополнительный канал сбыта для товаров Эмпилс, помимо оптового.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(li([
  new TextRun({ text: "Развитая цифровая инфраструктура: ", size: 22, bold: true, color: c(P.positive), font: FONT }),
  new TextRun({ text: "CRM-система, личный кабинет, онлайн-заказы 24/7 — снижает операционные издержки на интеграцию нового поставщика.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(li([
  new TextRun({ text: "Упрощённая структура собственности: ", size: 22, bold: true, color: c(P.positive), font: FONT }),
  new TextRun({ text: "единственный учредитель = генеральный директор. Решения принимаются быстро, без корпоративных согласований.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(h2("7.3. Риски и зоны внимания"));

children.push(li([
  new TextRun({ text: "Возражение по сезону 2026: ", size: 22, bold: true, color: c(P.negative), font: FONT }),
  new TextRun({ text: "клиент прямо заявил, что на сезон 2026 г. новых контрактов не заключает. Сентябрь 2026 г. — крайний срок для переговоров о контракте на сезон 2027 г. Просрочка = потеря сезона.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(li([
  new TextRun({ text: "Конкуренция с уже работающим дистрибьютором: ", size: 22, bold: true, color: c(P.negative), font: FONT }),
  new TextRun({ text: "на территории есть ещё один дистрибьютор Эмпилс — требуется развести ассортимент по непересекающимся SKU, иначе каннибализируем собственные продажи.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(li([
  new TextRun({ text: "Уже сильные позиции Текс/Лакра: ", size: 22, bold: true, color: c(P.warning), font: FONT }),
  new TextRun({ text: "в сегменте масс-маркет «Вертикаль» уже работает с двумя ведущими брендами. Замещение их на Ореол/Расцвет потребует весомых ценовых или сервисных аргументов.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(li([
  new TextRun({ text: "Концентрация решений в одних руках: ", size: 22, bold: true, color: c(P.warning), font: FONT }),
  new TextRun({ text: "Евдоков В.И. — единственный учредитель и директор. Если по какой-то причине он недоступен или меняет стратегию, переговоры могут застопориться.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(li([
  new TextRun({ text: "Отсутствие публичных финансовых данных: ", size: 22, bold: true, color: c(P.warning), font: FONT }),
  new TextRun({ text: "прибыль и рентабельность не раскрыты в открытых источниках. Перед заключением крупного контракта желательна проверка через СПАРК/Контур.Фокус для оценки платёжеспособности.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(li([
  new TextRun({ text: "Широкая диверсификация: ", size: 22, bold: true, color: c(P.warning), font: FONT }),
  new TextRun({ text: "26 ОКВЭД, ассортимент 35 000 SKU по 22 категориям. ЛКМ — только часть бизнеса (~20% SKU). Эмпилс не получит эксклюзивного внимания, нужно бороться за долю в фокусе команды.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(h2("7.4. Возможность СТМ-партнёрства"));

children.push(p([
  "Это ключевая точка захода на переговоры. «Вертикаль» уже имеет собственную ТМ V (например, грунт-эмаль 3 в 1 V), запущенную в 2008 году и активно укрепляемую с 2021 г. В истории компании указано: ",
  ["«укрепление портфеля собственных торговых марок в ЛКМ, инструменте и товарах для сада»", { italic: true }],
  ". В клиентской карточке зафиксирован прямой интерес клиента к ",
  ["«изготовлению собственных ТМ»", { bold: true, color: P.accent }],
  " и его готовность рассмотреть дешёвую СТМ."
]));

children.push(p([
  "Эмпилс как завод-производитель с собственными ЛКМ-линейками (Ореол, Ореол Дисконт, Простокрашено! БАУ, Расцвет) может предложить «Вертикали» ",
  ["контрактное производство дешёвой СТМ под брендом клиента", { bold: true, color: P.accent }],
  ". Это даёт «Вертикали» маржинальность и независимость от внешних брендов, а Эмпилс получает стабильный объём отгрузок и нового канального партнёра. Конкретные категории для СТМ-партнёрства:"
]));

children.push(nli("Грунт-эмали 3 в 1 (по аналогии с существующей СТМ V, но расширение по цветам и фасовкам)."));
children.push(nli("Краски интерьерные и фасадные в эконом-сегменте (Ореол Дисконт как референс)."));
children.push(nli("Эмали ПФ-115 (самый ходовой SKU в ABC-анализе по дистрибьюторам Эмпилс)."));
children.push(nli("Грунтовки (ГФ-021, акриловые, глубокого проникновения)."));
children.push(nli("Защитно-декоративные покрытия для дерева (конкурент Акватексу)."));

// === РАЗДЕЛ 8: КРАТКИЕ ВЫВОДЫ И РЕКОМЕНДАЦИИ ===
children.push(h1("8. Краткие выводы и рекомендации для переговоров"));

children.push(p(
  "Основные выводы по компании «Вертикаль» и рекомендации для подготовки к повторным переговорам в сентябре 2026 г.:"
));

children.push(nli([
  new TextRun({ text: "Контрагент крупный и устойчивый: ", size: 22, bold: true, color: c(P.positive), font: FONT }),
  new TextRun({ text: "30 лет на рынке, 12 млрд руб. выручки (2025), 7 складов, 900 сотрудников, 20 000 покупателей. Банкротство или неплатёжеспособность маловероятны — это надёжный долгосрочный партнёр.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(nli([
  new TextRun({ text: "Переговоры идти обязательно: ", size: 22, bold: true, color: c(P.primary), font: FONT }),
  new TextRun({ text: "10.04.2026 компания сменила основной ОКВЭД на «оптовая торговля ЛКМ» — это стратегический фокус на нашей категории. Эмпилс попадает в зону их главного интереса.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(nli([
  new TextRun({ text: "Главный козырь — СТМ: ", size: 22, bold: true, color: c(P.accent), font: FONT }),
  new TextRun({ text: "клиент уже имеет СТМ V и прямо заинтересован в дешёвой собственной ТМ. Эмпилс как завод-производитель должен прийти с конкретным предложением по контрактному производству 3–5 SKU под брендом «Вертикали».", size: 22, color: c(P.body), font: FONT }),
]));

children.push(nli([
  new TextRun({ text: "Безопасное замещение: ", size: 22, bold: true, color: c(P.warning), font: FONT }),
  new TextRun({ text: "предлагать бренды Эмпилс (Ореол, ОРЕОЛ TITAN, Расцвет, Простокрашено! БАУ, Ореол Дисконт) на замену Тексу, Лакре и Рогнеде — только в непересекающихся SKU, чтобы не каннибализировать продажи уже работающего на территории дистрибьютора Эмпилс.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(nli([
  new TextRun({ text: "Заход через премиум-сегмент: ", size: 22, bold: true, color: c(P.positive), font: FONT }),
  new TextRun({ text: "AvantGarde и NEWTONE не конкурируют напрямую с Лакрой/Тексом — это дополнение к Tikkurila и ABC Farben. Заход через премиум менее конфликтен.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(nli([
  new TextRun({ text: "Подготовить расширенную финпроверку: ", size: 22, bold: true, color: c(P.warning), font: FONT }),
  new TextRun({ text: "через СПАРК или Контур.Фокус до переговоров получить данные по прибыли, рентабельности, арбитражной практике. Это позволит точнее оценить кредитный лимит и условия отсрочки платежа.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(nli([
  new TextRun({ text: "Канал захода: ", size: 22, bold: true, color: c(P.primary), font: FONT }),
  new TextRun({ text: "ЛПР — Юрий Григорьевский (8 903 030-29-65, grigoryevsky@vertical.ru). При недоступности — многоканальный 8 800 222-4-000 или форма обратной связи на сайте. До сентябрьской встречи — связаться в августе для подтверждения даты и повестки.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(nli([
  new TextRun({ text: "Учитывать сроки: ", size: 22, bold: true, color: c(P.negative), font: FONT }),
  new TextRun({ text: "возражение клиента — «на сезон 2026 г. новых контрактов не заключают». Цель — закрыть переговоры до конца сентября 2026 г., чтобы попасть в бюджет сезона 2027 г.", size: 22, color: c(P.body), font: FONT }),
]));

// === ИСТОЧНИКИ ===
children.push(h1("Источники информации"));

children.push(li("Выписка из ЕГРЮЛ по ИНН 5027161766, полученная через сервис ФНС egrul.nalog.ru 21.10.2026 (10 страниц)."));
children.push(li("Сайт ООО «Вертикаль» — https://vertical.ru: главная страница, страница «О компании» (/company), «Контакты» (/contacts), «Доставка» (/about-delivery), «Новости» (/news), каталог ЛКМ (/catalog/lakokrasocnye-materialy)."));
children.push(li("Внутренняя CRM-карточка Эмпилс: лист «Белые пятна», строка «Вертикаль» (файл «Вертикаль.xlsx»)."));
children.push(li("Лист «Программы» из файла «Вертикаль.xlsx» — информация о программах работы с ключевыми дистрибьюторами в Центральном федеральном округе."));

// =================== ДОКУМЕНТ ===================
const doc = new Document({
  creator: "Z.ai",
  title: "Справка по дистрибьютору «Вертикаль»",
  description: "Сводная бизнес-аналитическая справка по ООО «Вертикаль» (ИНН 5027161766)",
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
      reference: "main-numbering",
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
            text: "Справка: ООО «Вертикаль» (ИНН 5027161766) — внутренний документ",
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
            new TextRun({
              text: "Стр. ",
              size: 18, color: c(P.muted), font: FONT,
            }),
            new TextRun({
              children: [PageNumber.CURRENT],
              size: 18, color: c(P.muted), font: FONT,
            }),
            new TextRun({
              text: " из ",
              size: 18, color: c(P.muted), font: FONT,
            }),
            new TextRun({
              children: [PageNumber.TOTAL_PAGES],
              size: 18, color: c(P.muted), font: FONT,
            }),
          ],
        })],
      }),
    },
    children: children,
  }],
});

Packer.toBuffer(doc).then(buffer => {
  const outPath = "/home/z/my-project/download/Справка_Вертикаль_дистрибьютор.docx";
  fs.writeFileSync(outPath, buffer);
  console.log(`✓ Документ сохранён: ${outPath}`);
  console.log(`  Размер: ${buffer.length.toLocaleString()} байт`);
  console.log(`  Количество абзацев: ${children.length}`);
});
