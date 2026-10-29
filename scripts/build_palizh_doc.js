/**
 * TITAN x PALIZH — Идеи кооперации
 * Документ: DOCX, R1 cover, DS-1 палитра
 */
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, PageNumber, PageBreak,
  BorderStyle, ShadingType, WidthType, TableLayoutType, NumberFormat,
  SectionType,
} = require("docx");
const fs = require("fs");

// Палитра DS-1
const P = {
  bg: "0B1C2C",
  accent: "529286",
  cover: { titleColor: "FFFFFF", subtitleColor: "B0B8C0", metaColor: "90989F", footerColor: "687078" },
  table: { headerBg: "529286", headerText: "FFFFFF", innerLine: "BECFCC", surface: "E8ECEB" },
};

const BODY_COLOR = "1A2B3D";
const HEADING_COLOR = "0B1C2C";
const ACCENT_DARK = "3D6B62";
const c = (hex) => hex.replace("#", "");

const NB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: NB, bottom: NB, left: NB, right: NB };
const allNoBorders = { top: NB, bottom: NB, left: NB, right: NB, insideHorizontal: NB, insideVertical: NB };
const thinBorder = (color) => ({ style: BorderStyle.SINGLE, size: 4, color });
const thinBorders = (color) => ({
  top: thinBorder(color), bottom: thinBorder(color),
  left: thinBorder(color), right: thinBorder(color),
  insideHorizontal: thinBorder(color), insideVertical: thinBorder(color),
});

// === Cover helpers ===
function splitTitleLines(title, charsPerLine) {
  if (title.length <= charsPerLine) return [title];
  const breakAfter = new Set([...'—–·/', ...' \t']);
  const lines = [];
  let remaining = title;
  while (remaining.length > charsPerLine) {
    let breakAt = -1;
    for (let i = charsPerLine; i >= Math.floor(charsPerLine * 0.6); i--) {
      if (i < remaining.length && breakAfter.has(remaining[i - 1])) { breakAt = i; break; }
    }
    if (breakAt === -1) {
      const limit = Math.min(remaining.length, Math.ceil(charsPerLine * 1.3));
      for (let i = charsPerLine + 1; i < limit; i++) {
        if (breakAfter.has(remaining[i - 1])) { breakAt = i; break; }
      }
    }
    if (breakAt === -1) breakAt = charsPerLine;
    lines.push(remaining.slice(0, breakAt).trim());
    remaining = remaining.slice(breakAt).trim();
  }
  if (remaining) lines.push(remaining);
  if (lines.length > 1 && lines[lines.length - 1].length <= 2) {
    const last = lines.pop();
    lines[lines.length - 1] += " " + last;
  }
  return lines;
}

function calcTitleLayout(title, maxWidthTwips, preferredPt = 36, minPt = 22) {
  const charWidth = (pt) => pt * 12;
  const charsPerLine = (pt) => Math.floor(maxWidthTwips / charWidth(pt));
  let titlePt = preferredPt, lines;
  while (titlePt >= minPt) {
    const cpl = charsPerLine(titlePt);
    if (cpl < 4) { titlePt -= 2; continue; }
    lines = splitTitleLines(title, cpl);
    if (lines.length <= 3) break;
    titlePt -= 2;
  }
  if (!lines || lines.length > 3) {
    lines = splitTitleLines(title, charsPerLine(minPt));
    titlePt = minPt;
  }
  return { titlePt, titleLines: lines };
}

function calcCoverSpacing({ titleLineCount = 1, titlePt = 36, hasSubtitle = false, hasEnglishLabel = false, metaLineCount = 0, fixedHeight = 800, pageHeight = 16838 }) {
  const SAFETY = 1200;
  const usableHeight = pageHeight - SAFETY;
  const titleHeight = titleLineCount * (titlePt * 23 + 200);
  const subtitleHeight = hasSubtitle ? (12 * 23 + 600) : 0;
  const englishLabelHeight = hasEnglishLabel ? (9 * 23 + 600) : 0;
  const metaHeight = metaLineCount * (10 * 23 + 100);
  const implicitParaHeight = 3 * 300;
  const contentHeight = titleHeight + subtitleHeight + englishLabelHeight + metaHeight + fixedHeight + implicitParaHeight;
  const safeRemaining = Math.max(usableHeight - contentHeight, 400);
  const FOOTER_MIN = 800;
  const rawTop = Math.floor(safeRemaining * 0.45);
  const rawBottom = Math.floor(safeRemaining * 0.45);
  const bottomSpacing = Math.max(rawBottom, FOOTER_MIN);
  const topSpacing = Math.max(rawTop - Math.max(0, FOOTER_MIN - rawBottom), 400);
  return { topSpacing, bottomSpacing };
}

function buildCoverR1(config) {
  const padL = 1200, padR = 800;
  const availableWidth = 11906 - padL - padR - 300;
  const { titlePt, titleLines } = calcTitleLayout(config.title, availableWidth, 32, 22);
  const titleSize = titlePt * 2;
  const spacing = calcCoverSpacing({
    titleLineCount: titleLines.length, titlePt,
    hasSubtitle: !!config.subtitle, hasEnglishLabel: !!config.englishLabel,
    metaLineCount: (config.metaLines || []).length, fixedHeight: 400,
  });
  const accentLeft = { style: BorderStyle.SINGLE, size: 8, color: P.accent, space: 12 };
  const children = [];
  children.push(new Paragraph({ spacing: { before: spacing.topSpacing } }));
  if (config.englishLabel) {
    children.push(new Paragraph({
      indent: { left: padL, right: padR }, spacing: { after: 500 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: P.accent, space: 8 } },
      children: [new TextRun({
        text: config.englishLabel.split("").join("  "),
        size: 18, color: P.accent, font: { ascii: "Calibri" }, characterSpacing: 40,
      })],
    }));
  }
  for (let i = 0; i < titleLines.length; i++) {
    children.push(new Paragraph({
      indent: { left: padL },
      spacing: { after: i < titleLines.length - 1 ? 100 : 300, line: Math.ceil(titlePt * 23), lineRule: "atLeast" },
      children: [new TextRun({
        text: titleLines[i], size: titleSize, bold: true,
        color: P.cover.titleColor, font: { ascii: "Arial" },
      })],
    }));
  }
  if (config.subtitle) {
    children.push(new Paragraph({
      indent: { left: padL }, spacing: { after: 800 },
      children: [new TextRun({
        text: config.subtitle, size: 24, color: P.cover.subtitleColor, font: { ascii: "Arial" },
      })],
    }));
  }
  for (const line of (config.metaLines || [])) {
    children.push(new Paragraph({
      indent: { left: padL + 200 }, spacing: { after: 80 },
      border: { left: accentLeft },
      children: [new TextRun({
        text: line, size: 24, color: P.cover.metaColor, font: { ascii: "Arial" },
      })],
    }));
  }
  children.push(new Paragraph({ spacing: { before: spacing.bottomSpacing } }));
  children.push(new Paragraph({
    indent: { left: padL, right: padR },
    border: { top: { style: BorderStyle.SINGLE, size: 2, color: P.accent, space: 8 } },
    spacing: { before: 200 },
    children: [
      new TextRun({ text: config.footerLeft || "", size: 16, color: P.cover.footerColor, font: { ascii: "Arial" } }),
      new TextRun({ text: "                                        " }),
      new TextRun({ text: config.footerRight || "", size: 16, color: P.cover.footerColor, font: { ascii: "Arial" } }),
    ],
  }));
  return [new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: allNoBorders,
    rows: [new TableRow({
      height: { value: 16838, rule: "exact" },
      children: [new TableCell({
        shading: { type: ShadingType.CLEAR, fill: P.bg },
        borders: noBorders,
        children,
      })],
    })],
  })];
}

// === Body helpers ===
function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 480, after: 240, line: 312 },
    keepNext: true,
    children: [new TextRun({
      text, bold: true, size: 32, color: c(HEADING_COLOR), font: { ascii: "Arial" },
    })],
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 320, after: 180, line: 312 },
    keepNext: true,
    children: [new TextRun({
      text, bold: true, size: 28, color: c(HEADING_COLOR), font: { ascii: "Arial" },
    })],
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 240, after: 120, line: 312 },
    keepNext: true,
    children: [new TextRun({
      text, bold: true, size: 26, color: c(ACCENT_DARK), font: { ascii: "Arial" },
    })],
  });
}

function body(text) {
  let runs;
  if (Array.isArray(text)) {
    runs = text.map(seg => new TextRun({
      text: seg.t, size: 24, color: c(BODY_COLOR),
      bold: !!seg.b, italics: !!seg.i, font: { ascii: "Arial" },
    }));
  } else {
    runs = [new TextRun({ text, size: 24, color: c(BODY_COLOR), font: { ascii: "Arial" } })];
  }
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 312, after: 120 },
    indent: { firstLine: 480 },
    children: runs,
  });
}

function bullet(text, level = 0) {
  let runs;
  if (Array.isArray(text)) {
    runs = text.map(seg => new TextRun({
      text: seg.t, size: 24, color: c(BODY_COLOR),
      bold: !!seg.b, italics: !!seg.i, font: { ascii: "Arial" },
    }));
  } else {
    runs = [new TextRun({ text, size: 24, color: c(BODY_COLOR), font: { ascii: "Arial" } })];
  }
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { line: 312, after: 80 },
    bullet: { level },
    children: runs,
  });
}

function thCell(text, opts = {}) {
  return new TableCell({
    width: opts.width ? { size: opts.width, type: WidthType.PERCENTAGE } : undefined,
    shading: { type: ShadingType.CLEAR, fill: P.table.headerBg },
    margins: { top: 100, bottom: 100, left: 120, right: 120 },
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { line: 280 },
      children: [new TextRun({
        text, bold: true, size: 22, color: P.table.headerText, font: { ascii: "Arial" },
      })],
    })],
  });
}

function tdCell(text, opts = {}) {
  let runs;
  if (Array.isArray(text)) {
    runs = text.map(seg => new TextRun({
      text: seg.t, size: 22, color: c(BODY_COLOR),
      bold: !!seg.b, italics: !!seg.i, font: { ascii: "Arial" },
    }));
  } else {
    runs = [new TextRun({
      text: String(text), size: 22, color: c(BODY_COLOR),
      bold: !!opts.bold, font: { ascii: "Arial" },
    })];
  }
  return new TableCell({
    width: opts.width ? { size: opts.width, type: WidthType.PERCENTAGE } : undefined,
    shading: opts.surface ? { type: ShadingType.CLEAR, fill: P.table.surface } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.LEFT,
      spacing: { line: 280 },
      children: runs,
    })],
  });
}

function makeTable(headers, rows, widths) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: thinBorders(P.table.innerLine),
    rows: [
      new TableRow({
        tableHeader: true, cantSplit: true,
        children: headers.map((h, i) => thCell(h, { width: widths[i] })),
      }),
      ...rows.map((row, ri) => new TableRow({
        cantSplit: true,
        children: row.map((cell, ci) => {
          const opts = { width: widths[ci] };
          if (ri % 2 === 1) opts.surface = true;
          return tdCell(cell, opts);
        }),
      })),
    ],
  });
}

// ============================================================
// СОДЕРЖАНИЕ
// ============================================================
const blocks = [];

// === 1. Резюме ===
blocks.push(h1("1. Резюме"));
blocks.push(body([
  {t: "Документ фиксирует ", b: false},
  {t: "5 конкретных идей кооперации с Palizh (ООО «Новый дом», Ижевск)", b: true},
  {t: " — российским производителем ЛКМ с 1996 года, владеющим брендами Palizh, GOL, Lesnoy, KolerPark. Идеи основаны на переговорах с Ириной Килиной (директор по маркетингу Palizh) и анализе сайта palizh.com. Ключевой вывод: Palizh — это не конкурент Эмпилс в товарных категориях, а "},
  {t: "комплементарный партнёр", b: true},
  {t: ", чья сильная сторона (цвет, колеровка, сервис для дизайнеров) закрывает слабые стороны Эмпилс."}
]));
blocks.push(body([
  {t: "Окно возможности ", b: true},
  {t: "по основной идее (расширение поставок колеровочных паст) ограничено: бюджет Palizh на следующий сезон закрывается в сентябре–октябре 2026. Если не успеть с обновлённым ценовым предложением до 15.10.2026, контракт переносится на 2027 год."}
]));

// === 2. Профиль Palizh ===
blocks.push(h1("2. Профиль партнёра"));
blocks.push(body([
  {t: "ООО «Новый дом» — российский производитель ЛКМ с 30-летней историей, headquartered в Ижевске (Удмуртия). Компания работает в 4 направлениях: бытовые колеровочные пасты, промышленные колеровочные пасты, колеровочное оборудование и готовые ЛКМ (лаки, эмали, пропитки, огнебиозащита, акриловые краски). Клиентская база — 2 000+ клиентов в 34 странах. Имеет знак «Сделано в России», собственную лабораторию и академию Palizh Academy."}
]));

blocks.push(h3("Таблица 1. Профиль Palizh"));
blocks.push(makeTable(
  ["Параметр", "Значение"],
  [
    ["Юридическое лицо", "ООО «Новый дом» (ИНН 1831057110)"],
    ["Бренды", "Palizh, GOL, Lesnoy, KolerPark"],
    ["Год основания", "1996 (30 лет на рынке)"],
    ["Локация", "г. Ижевск, ул. Салютовская 31"],
    ["Контакты", "8 (800) 234-67-00, market@palitra.udm.ru"],
    ["Клиентская база", "2 000+ клиентов в 34 странах"],
    ["Ключевые компетенции", "Колеровочные пасты, цветовые коллекции, обучение, работа с дизайнерами"],
    ["Каналы сбыта", "B2B (промышленность), DIY (для дома), художественный сегмент (для творчества)"],
    ["Сервисы", "Контрактное производство, обслуживание оборудования, обучение, услуги лаборатории"],
    ["Маркетинговые активы", "Palizh Academy (3-дневный форум), книга «Продсвет», сеть шоурумов с дизайнерами"],
    ["Сильная сторона", "Цвет и сервис — экспертиза в колеровке и работе с дизайнерским каналом"],
    ["Слабая сторона", "Уязвимы в готовых ЛКМ (химия, технологии) — это сильная сторона Эмпилс"],
  ],
  [30, 70]
));

blocks.push(h2("2.1. Сильные и слабые стороны Palizh"));
blocks.push(body([
  {t: "Сильные стороны: ", b: true},
  {t: "(1) 30-летний опыт в колеровке и цвете — это редкая экспертиза на российском рынке; (2) собственная академия и реферальная программа — устойчивая лояльность дистрибьюторов; (3) сеть шоурумов с потоком дизайнеров (20–30 студий по РФ); (4) знак «Сделано в России» — преимущество в тендерах и экспорте; (5) контрактное производство и собственная лаборатория — гибкость под B2B-задачи."}
]));
blocks.push(body([
  {t: "Слабые стороны: ", b: true},
  {t: "(1) ограниченная продуктовая линейка готовых ЛКМ — основной фокус на колеровке и специализированных продуктах (акрил, огнебиозащита, пропитки); (2) нет грунт-эмалей по металлу в категории «3 в 1» — прямая категория TITAN; (3) нет молотковой эмали — прямая категория TITAN; (4) ограниченный ассортимент промышленных грунтовок — где Эмпилс имеет избыточную экспертизу."}
]));
blocks.push(body([
  {t: "Вывод: ", b: true},
  {t: "Palizh и Эмпилс не пересекаются по основным продуктовым категориям. Palizh — это цвет и сервис; Эмпилс — это химия и готовые ЛКМ. Это идеальная база для комплементарной кооперации."}
]));

// === 3. Идеи кооперации ===
blocks.push(h1("3. Идеи кооперации"));
blocks.push(body([
  {t: "Ниже — 5 проработанных идей кооперации. Для каждой указаны: суть, что делает Эмпилс, что делает Palizh, сроки, ответственные, метрики успеха. Идеи упорядочены по приоритету: от самой срочной (окно закрывается в октябре 2026) к среднесрочным."}
]));

// --- Идея 1 ---
blocks.push(h2("3.1. Расширение поставок колеровочных паст"));
blocks.push(body([
  {t: "Суть: ", b: true},
  {t: "Palizh уже закупает у Эмпилс промышленные пигменты. По итогам переговоров на выставке (сентябрь 2026) Ирина Килина подтвердила готовность расширить закупку колеровочных паст до контейнерных объёмов — при условии снижения цены ниже текущего уровня. Эмпилс уже снизил цену; осталось отправить обновлённое предложение и образцы."}
]));
blocks.push(body([
  {t: "Объём контракта: ", b: true},
  {t: "от контейнера одного цвета или от вагона — для обоих форматов готовы. Ориентировочный годовой потенциал: 80–150 тонн колеровочных паст при цене ниже аналогов из Китая и «Коллористики»."}
]));
blocks.push(body([{t: "Действия Эмпилс:", b: true}]));
blocks.push(bullet([{t: "До 15.10.2026: ", b: true}, {t: "отправить обновлённое ценовое предложение + свежие образцы колеровочных паст (снабжение Эмпилс + Ирина Килина как точка контроля)"}]));
blocks.push(bullet([{t: "До 30.10.2026: ", b: true}, {t: "получить обратную связь от лаборатории Palizh по качеству (попадание в цвет, отклонения, совместимость с базовыми материалами)"}]));
blocks.push(bullet([{t: "До 30.11.2026: ", b: true}, {t: "подписать годовой контракт с фиксацией цены и объёмов по кварталам"}]));
blocks.push(body([
  {t: "Метрика успеха: ", b: true},
  {t: "контракт на ≥ 80 тонн/год, цена ниже аналогов «Коллористики» на 8–12%, повторная закупка в Q1 2027."}
]));
blocks.push(body([
  {t: "Риски: ", b: true},
  {t: "конкурент «Коллористика» (от которого Palizh сейчас уходит по качеству) может снизить цену. Митигация: фиксированная цена в годовом контракте + сервис «резервный склад» на 30 дней."}
]));

// --- Идея 2 ---
blocks.push(h2("3.2. СТМ-аэрозоли под брендом Palizh"));
blocks.push(body([
  {t: "Суть: ", b: true},
  {t: "Palizh 2 года назад запустил собственное производство аэрозольных баллончиков, спрос есть и хочет расти. Однако собственное производство Palizh хочет масштабировать без капитальных вложений — ищет контрактного производителя. Аэрозольная компетенция Эмпилс требует проверки."}
]));
blocks.push(body([
  {t: "Объём контракта: ", b: true},
  {t: "потенциал 100–300 тыс. баллонов/год под СТМ Palizh (аэрозольные эмали, грунты, лаки). Средний ценовой сегмент — конкуренты Kudo, New Ton, Asteria."}
]));
blocks.push(body([{t: "Действия Эмпилс:", b: true}]));
blocks.push(bullet([{t: "До 01.12.2026: ", b: true}, {t: "проверить компетенцию по аэрозольному производству (R&D Эмпилс + контракты с контрактными производителями в РФ)"}]));
blocks.push(bullet([{t: "До 15.01.2027: ", b: true}, {t: "подготовить КП по контрактному производству СТМ-аэрозолей: цена за баллон, MOQ, сроки поставки"}]));
blocks.push(bullet([{t: "До 28.02.2027: ", b: true}, {t: "провести пилотную партию 5 000–10 000 баллонов под брендом Palizh для теста рынка"}]));
blocks.push(body([
  {t: "Метрика успеха: ", b: true},
  {t: "пилотная партия продана через каналы Palizh, повторный заказ ≥ 50 000 баллонов в течение 6 месяцев."}
]));
blocks.push(body([
  {t: "Риски: ", b: true},
  {t: "(1) у Эмпилс нет собственной аэрозольной линии — потребуется контракт с третьей стороной. (2) Качество может не совпасть с ожиданиями Palizh (их бывший собственный опыт был неудачным). Митигация: референс-партия 5 000 шт. с возвратом, если не пройдёт тест."}
]));

// --- Идея 3 ---
blocks.push(h2("3.3. Совместный шоурум с дизайнерами (TITAN + AVANTGARDE)"));
blocks.push(body([
  {t: "Суть: ", b: true},
  {t: "У Palizh есть сеть шоурумов с потоком дизайнеров (20–30 студий по РФ). Эмпилс хочет войти в дизайнерский канал, но не имеет собственной сети. Модель: Palizh открывает двери своих шоурумов для выкладки TITAN и AVANTGARDE — как дополнение к ассортименту Palizh, не конкурируя с их продуктом."}
]));
blocks.push(body([
  {t: "Обмен: ", b: true},
  {t: "Palizh ставит TITAN и AVANTGARDE на витрину в 3 пилотных шоурумах (Москва + СПб + Поволжье). Эмпилс включает Palizh в свою программу обучения (мастер-классы в Ростове, Казани, Екатеринбурге) — Palizh получает доступ к аудитории Эмпилс."}
]));
blocks.push(body([{t: "Действия Эмпилс:", b: true}]));
blocks.push(bullet([{t: "До 30.11.2026: ", b: true}, {t: "согласовать список 3 пилотных шоурумов с Palizh; подписать пилотное соглашение на 6 месяцев"}]));
blocks.push(bullet([{t: "До 31.01.2027: ", b: true}, {t: "отгрузить POSm-материалы TITAN + AVANTGARDE (выкрасы, каталоги цветов, видео-инструкции)"}]));
blocks.push(bullet([{t: "До 30.06.2027: ", b: true}, {t: "оценить продажи TITAN через шоурумы Palizh, принять решение о масштабировании на 10–15 точек"}]));
blocks.push(body([
  {t: "Метрика успеха: ", b: true},
  {t: "продажи TITAN через шоурумы Palizh ≥ 400 000 ₽/квартал на точку к 6-му месяцу; 30% дизайнеров в шоурумах Palizh рекомендуют TITAN."}
]));

// --- Идея 4 ---
blocks.push(h2("3.4. Прикассовые витрины POSm: обмен макетами"));
blocks.push(body([
  {t: "Суть: ", b: true},
  {t: "У Palizh есть рабочие прикассовые витрины для колеровочных паст — формат с доказанным спросом. Palizh готов делиться макетами бесплатно — это был прямой запрос от Ирины Килиной. Эмпилс использует их для выкладки колеровочных паст в собственных точках продаж."}
]));
blocks.push(body([
  {t: "Ценность для Эмпилс: ", b: true},
  {t: "экономия 2–4 млн ₽ на разработке POSm с нуля. Palizh уже отработал формат: тип витрины, высота, ширина, цветовое решение, рекомендации по выкладке."}
]));
blocks.push(body([{t: "Действия Эмпилс:", b: true}]));
blocks.push(bullet([{t: "До 30.11.2026: ", b: true}, {t: "получить от Palizh технические чертежи прикассовых витрин + рекомендации по выкладке"}]));
blocks.push(bullet([{t: "До 31.12.2026: ", b: true}, {t: "адаптировать макеты под колеровочные пасты Эмпилс (размеры, логотип)"}]));
blocks.push(bullet([{t: "До 28.02.2027: ", b: true}, {t: "произвести пилотную партию 50 витрин; разместить в топ-50 точках продаж"}]));
blocks.push(body([
  {t: "Метрика успеха: ", b: true},
  {t: "рост продаж колеровочных паст на 25–40% в точках с новыми витринами (замер vs контрольная группа без витрин) за 3 месяца."}
]));

// --- Идея 5 ---
blocks.push(h2("3.5. Форум Palizh Academy (10–12 сентября 2027): командное участие"));
blocks.push(body([
  {t: "Суть: ", b: true},
  {t: "Palizh проводит 3-дневный форум Palizh Academy в Ижевске — экскурсия по заводу, 19 спикеров (включая федеральных и зарубежных), нетворкинг с 200+ B2B-партнёрами. Стоимость участия: 40 000 ₽/чел. при команде от 10 человек; реферальная программа позволяет получать места бесплатно."}
]));
blocks.push(body([
  {t: "Ценность для Эмпилс: ", b: true},
  {t: "(1) изучение модели Palizh Academy (можно воспроизвести для Эмпилс с фокусом на TITAN); (2) доступ к 200+ B2B-партнёрам Palizh как потенциальным клиентам Эмпилс; (3) обмен опытом по работе с дизайнерами — где Palizh эксперт."}
]));
blocks.push(body([{t: "Действия Эмпилс:", b: true}]));
blocks.push(bullet([{t: "До 31.03.2027: ", b: true}, {t: "войти в реферальную программу Palizh для накопления баллов и бесплатных мест"}]));
blocks.push(bullet([{t: "До 01.08.2027: ", b: true}, {t: "согласовать командировку 6–10 человек: коммерческий директор + B2B-отдел + R&D + маркетинг"}]));
blocks.push(bullet([{t: "10–12.09.2027: ", b: true}, {t: "командное участие в форуме, нетворкинг, сбор контактов партнёров Palizh"}]));
blocks.push(body([
  {t: "Метрика успеха: ", b: true},
  {t: "30+ новых B2B-контактов после форума; 3 пилотных контракта с партнёрами Palizh в течение 6 месяцев после мероприятия."}
]));

// === 4. Дополнительные возможности ===
blocks.push(h1("4. Дополнительные возможности (за рамками 5 основных идей)"));

blocks.push(h2("4.1. Бартер ЛКМ ↔ колеровочные пасты"));
blocks.push(body([
  {t: "Суть: ", b: true},
  {t: "По аналогии с моделью Bostik (для Технониколь). Эмпилс поставляет Palizh готовые ЛКМ (TITAN, AVANTGARDE) для собственного потребления или перепродажи. Palizh поставляет Эмпилс колеровочные пасты. Снижение кассового разрыва для обеих сторон — особенно ценно для Эмпилс в условиях дефицита оборотных средств."}
]));
blocks.push(body([
  {t: "Потенциал: ", b: true},
  {t: "бартерный объём 5–10 млн ₽ в год — эквивалентный обмен, без движения денежных средств."}
]));

blocks.push(h2("4.2. Совместная разработка цвета (премиальная молотковая линейка)"));
blocks.push(body([
  {t: "Суть: ", b: true},
  {t: "У Palizh сильная лаборатория цвета (колеровка, попадание в RAL, Nuance). У Эмпилс — сильное химическое R&D (формулы, технологии). Возможен совместный продукт в категории «декоративные ЛКМ с необычными эффектами» — например, премиальная молотковая линейка TITAN с расширенной палитрой (15–20 цветов вместо 3 текущих)."}
]));
blocks.push(body([
  {t: "Срок: ", b: true},
  {t: "разработка 6–9 месяцев; запуск к сезону 2028. Потенциал: 2–3 млн ₽/год дополнительной выручки в категории молотковой эмали."}
]));

blocks.push(h2("4.3. Маркетинговая коллаборация в художественном сегменте"));
blocks.push(body([
  {t: "Суть: ", b: true},
  {t: "Palizh развивает направление FluidArt (жидкий акрил для творчества) и краски для творчества. Это сегмент, где Эмпилс отсутствует. Возможно продвижение Palizh через каналы Эмпилс в DIY-сегменте (хозмаги), где Palizh слаб — обмен аудиториями."}
]));
blocks.push(body([
  {t: "Формат: ", b: true},
  {t: "включение FluidArt Palizh в каталог Эмпилс для хозмагов; встречное включение TITAN в каталог Palizh для художественных магазинов."}
]));

// === 5. Что просит Palizh взамен ===
blocks.push(h1("5. Что просит Palizh взамен"));
blocks.push(body([
  {t: "По итогам переговоров Palizh сформулировал конкретные запросы к Эмпилс. Эти запросы — условие продолжения сотрудничества по основным идеям (особенно по п. 3.1 — расширение поставок паст)."}
]));
blocks.push(makeTable(
  ["№", "Запрос Palizh", "Срок", "Ответственный Эмпилс", "Приоритет"],
  [
    ["1", "Обновлённое ценовое предложение по колеровочным пастам (ниже текущей цены)", "15.10.2026", "ТП / Снабжение", "🔴 Срочно"],
    ["2", "Свежие образцы колеровочных паст для лабораторного сравнения с «Коллористикой»", "15.10.2026", "ТП / Лаборатория", "🔴 Срочно"],
    ["3", "Прикассовые витрины POSm с рекомендациями по выкладке (макеты от Palizh)", "30.11.2026", "Маркетинг", "🟡 Высокий"],
    ["4", "Сервис клиентский: рекомендации по продажам продукта Palizh в точках Эмпилс", "30.11.2026", "Маркетинг / ТП", "🟡 Высокий"],
    ["5", "Подтверждение участия в форуме Palizh Academy (10–12.09.2027)", "01.08.2027", "Коммерческий директор", "🟢 Средний"],
  ],
  [5, 50, 14, 21, 10]
));

// === 6. Сводная таблица идей ===
blocks.push(h1("6. Сводная таблица идей кооперации"));
blocks.push(makeTable(
  ["№", "Идея", "Срок запуска", "Потенциал (млн ₽/год)", "Приоритет"],
  [
    ["1", "Расширение поставок колеровочных паст (контейнер/вагон)", "Q4 2026", "15–25", "🔴 Срочно"],
    ["2", "СТМ-аэрозоли под брендом Palizh (контрактное производство)", "Q1 2027", "30–80 (3 года)", "🟡 Высокий"],
    ["3", "Совместный шоурум с дизайнерами (TITAN + AVANTGARDE)", "Q1 2027", "8–15", "🟡 Высокий"],
    ["4", "Прикассовые витрины POSm (обмен макетами)", "Q4 2026", "Экономия 2–4 + рост 30%", "🟡 Высокий"],
    ["5", "Форум Palizh Academy (командное участие)", "Q3 2027", "Нетворкинг + опыт", "🟢 Средний"],
    ["+", "Бартер ЛКМ ↔ пасты (доп. возможность)", "Q1 2027", "5–10 (бартер)", "🟢 Средний"],
    ["+", "Совместная разработка цвета (молотковая премиум)", "Q4 2027", "2–3 (3 года)", "🟢 Долгосрочный"],
    ["+", "Маркетинговая коллаборация (FluidArt ↔ TITAN)", "Q2 2027", "Обмен аудиторией", "🟢 Долгосрочный"],
  ],
  [5, 38, 14, 25, 18]
));

// === 7. Дорожная карта ===
blocks.push(h1("7. Дорожная карта реализации"));
blocks.push(body([
  {t: "Дорожная карта сгруппирована по 3 фазам: подготовка (октябрь–декабрь 2026), пилот (Q1 2027), масштабирование (Q2–Q3 2027). Каждая фаза завершается контрольной точкой с пересмотром приоритетов."}
]));

blocks.push(h3("Таблица 2. Дорожная карта по фазам"));
blocks.push(makeTable(
  ["Фаза", "Месяцы", "Ключевые активности", "Контрольная точка"],
  [
    ["Подготовка", "Окт–Дек 2026", "Отправка КП по пастам; образцы; проверка компетенции по аэрозолям; получение макетов POSm", "Решение по контракту на пасты (15.12.2026)"],
    ["Пилот", "Янв–Мар 2027", "Пилотная партия СТМ-аэрозолей (5 000 шт.); пилот 3 шоурумов Palizh; пилот 50 POSm-витрин", "Решение о масштабировании (31.03.2027)"],
    ["Масштабирование", "Апр–Сен 2027", "Расширение на 10–15 шоурумов; контракт на 100–300 тыс. баллонов; запуск бартера; участие в форуме Palizh Academy", "Отчёт за 12 мес. (30.09.2027)"],
  ],
  [16, 16, 50, 18]
));

// === 8. Риски и митигация ===
blocks.push(h1("8. Риски и митигация"));
blocks.push(makeTable(
  ["Риск", "Вероятность", "Влияние", "Митигация"],
  [
    ["Окно возможности по пастам закрывается 15.10.2026 (бюджет Palizh)", "Высокая", "Высокое", "Параллельная отправка КП + образцов в течение 7 дней; личный визит ТП"],
    ["Конкурент «Коллористика» снижает цену", "Средняя", "Высокое", "Годовой контракт с фиксацией цены; сервис «резервный склад 30 дней»"],
    ["У Эмпилс нет аэрозольной компетенции", "Высокая", "Высокое", "Партнёрство с контрактным производителем; пилот 5 000 шт. с правом возврата"],
    ["Palizh отказывается от СТМ-аэрозолей после пилота", "Средняя", "Среднее", "Минимальный MOQ 5 000 шт.; возврат партии в случае провала"],
    ["Дизайнеры в шоурумах Palizh не рекомендуют TITAN", "Средняя", "Среднее", "Обучение продавцов Palizh по TITAN; бонусы за продажу"],
    ["Palizh передумает делиться POSm-макетами", "Низкая", "Низкое", "Фиксация договорённости в пилотном соглашении от 30.11.2026"],
    ["Форум Palizh Academy не приносит новых контактов", "Низкая", "Низкое", "Подготовка списка целевых партнёров Palizh заранее"],
  ],
  [42, 14, 14, 30]
));

// === 9. Заключение ===
blocks.push(h1("9. Заключение"));
blocks.push(body([
  {t: "Palizh — это комплементарный партнёр, а не конкурент Эмпилс. Их сильная сторона (цвет, колеровка, дизайнерский канал, сервис) закрывает слабые стороны Эмпилс (дизайнерский канал, POSm, обучение). Сильная сторона Эмпилс (химия, готовые ЛКМ, грунт-эмали TITAN) закрывает слабые стороны Palizh (готовые ЛКМ, молотковая эмаль, «3 в 1»)."}
]));
blocks.push(body([
  {t: "Ключевая идея №1 (расширение поставок колеровочных паст) — самая срочная: ", b: true},
  {t: "окно закрывается 15.10.2026. Если успеть с обновлённым ценовым предложением и образцами, Эмпилс получает контракт на 80–150 тонн/год и доступ к дизайн-каналу Palizh через идеи №3 и №4."}
]));
blocks.push(body([
  {t: "Стратегический потенциал кооперации: ", b: true},
  {t: "при успешной реализации всех 5 идей — суммарный годовой потенциал 60–130 млн ₽ к 2028 году, плюс доступ к дизайн-каналу и обмен опытом с одной из самых сервисно-развитых компаний ЛКМ-рынка России."}
]));
blocks.push(body([
  {t: "Следующие шаги: ", b: true},
  {t: "(1) согласовать ответственных по 5 идеям до 10.10.2026; (2) отправить обновлённое КП по пастам в течение 7 дней; (3) согласовать пилотное соглашение по шоурумам и POSm к 30.11.2026; (4) встреча с Ириной Килиной в ноябре 2026 для фиксации дорожной карты."}
]));

// ============================================================
// СБОРКА ДОКУМЕНТА
// ============================================================
const coverConfig = {
  title: "Кооперация с Palizh — 5 идей",
  subtitle: "Проработка конкретных проектов сотрудничества на основе переговоров на выставке (сентябрь 2026)",
  englishLabel: "TITAN x PALIZH COOPERATION",
  metaLines: [
    "Партнёр: ООО «Новый дом» (бренды Palizh, GOL, Lesnoy, KolerPark)",
    "Источник: palizh.com + переговоры с Ириной Килиной",
    "Подготовлено для: ЗАО «Эмпилс»",
    "Версия документа: 1.0",
    "Дата: сентябрь 2026",
  ],
  footerLeft: "TITAN x Palizh · v1.0",
  footerRight: "Конфиденциально · Внутренний документ",
};

const doc = new Document({
  creator: "ЗАО Эмпилс",
  title: "TITAN x Palizh — 5 идей кооперации",
  description: "Проработка конкретных проектов сотрудничества с Palizh",
  styles: {
    default: {
      document: {
        run: { font: { ascii: "Arial", eastAsia: "Arial" }, size: 24, color: c(BODY_COLOR) },
        paragraph: { spacing: { line: 312 } },
      },
    },
  },
  sections: [
    {
      properties: {
        page: { size: { width: 11906, height: 16838 }, margin: { top: 0, bottom: 0, left: 0, right: 0 } },
      },
      children: buildCoverR1(coverConfig),
    },
    {
      properties: {
        type: SectionType.NEXT_PAGE,
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, bottom: 1440, left: 1701, right: 1417 },
          pageNumbers: { start: 1, formatType: NumberFormat.DECIMAL },
        },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({
              text: "TITAN x Palizh · 5 идей кооперации",
              size: 18, color: "808080", font: { ascii: "Arial" },
            })],
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({
              children: [PageNumber.CURRENT],
              size: 18, color: "808080", font: { ascii: "Arial" },
            })],
          })],
        }),
      },
      children: blocks,
    },
  ],
});

Packer.toBuffer(doc).then(buf => {
  const outPath = "/home/z/my-project/download/TITAN_Palizh_кооперация_5_идей.docx";
  fs.writeFileSync(outPath, buf);
  console.log("OK saved:", outPath);
  console.log("Size:", buf.length, "bytes");
});
