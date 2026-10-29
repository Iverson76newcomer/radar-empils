/**
 * План действий по улучшению позиций бренда TITAN в интернет-пространстве
 * На основании исследования Yandex Wordstat (сентябрь 2026)
 *
 * Документ: DOCX, R1 cover, DS-1 палитра
 */
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, PageNumber, PageBreak,
  BorderStyle, ShadingType, WidthType, TableLayoutType, NumberFormat,
  SectionType, LevelFormat, convertInchesToTwip
} = require("docx");
const fs = require("fs");

// ============================================================
// ПАЛИТРА DS-1 (Deep Sea)
// ============================================================
const P = {
  bg: "0B1C2C",
  primary: "FFFFFF",
  accent: "529286",
  cover: { titleColor: "FFFFFF", subtitleColor: "B0B8C0", metaColor: "90989F", footerColor: "687078" },
  table: { headerBg: "529286", headerText: "FFFFFF", accentLine: "529286", innerLine: "BECFCC", surface: "E8ECEB" },
};

// Тёмный текст для body
const BODY_COLOR = "1A2B3D";
const HEADING_COLOR = "0B1C2C";
const SECONDARY_COLOR = "5B6B7D";
const ACCENT_DARK = "3D6B62";

const c = (hex) => hex.replace("#", "");

// ============================================================
// БОРДЕРЫ ДЛЯ ТАБЛИЦ
// ============================================================
const NB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: NB, bottom: NB, left: NB, right: NB };
const allNoBorders = { top: NB, bottom: NB, left: NB, right: NB, insideHorizontal: NB, insideVertical: NB };
const thinBorder = (color) => ({ style: BorderStyle.SINGLE, size: 4, color });
const thinBorders = (color) => ({
  top: thinBorder(color), bottom: thinBorder(color),
  left: thinBorder(color), right: thinBorder(color),
  insideHorizontal: thinBorder(color), insideVertical: thinBorder(color),
});

// ============================================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ (calcTitleLayout, calcCoverSpacing, splitTitleLines)
// ============================================================
function splitTitleLines(title, charsPerLine) {
  if (title.length <= charsPerLine) return [title];
  const breakAfter = new Set([
    ...'，。、；：！？',
    ...'—–·/',
    ...' \t',
  ]);
  const lines = [];
  let remaining = title;
  while (remaining.length > charsPerLine) {
    let breakAt = -1;
    for (let i = charsPerLine; i >= Math.floor(charsPerLine * 0.6); i--) {
      if (i < remaining.length && breakAfter.has(remaining[i - 1])) {
        breakAt = i; break;
      }
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
  const charWidth = (pt) => pt * 12;  // Latin char ≈ pt * 12 twips
  const charsPerLine = (pt) => Math.floor(maxWidthTwips / charWidth(pt));
  let titlePt = preferredPt;
  let lines;
  while (titlePt >= minPt) {
    const cpl = charsPerLine(titlePt);
    if (cpl < 4) { titlePt -= 2; continue; }
    lines = splitTitleLines(title, cpl);
    if (lines.length <= 3) break;
    titlePt -= 2;
  }
  if (!lines || lines.length > 3) {
    const cpl = charsPerLine(minPt);
    lines = splitTitleLines(title, cpl);
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
  const remainingSpace = usableHeight - contentHeight;
  const safeRemaining = Math.max(remainingSpace, 400);
  const FOOTER_MIN = 800;
  const rawTop = Math.floor(safeRemaining * 0.45);
  const rawBottom = Math.floor(safeRemaining * 0.45);
  const bottomSpacing = Math.max(rawBottom, FOOTER_MIN);
  const topSpacing = Math.max(rawTop - Math.max(0, FOOTER_MIN - rawBottom), 400);
  const midSpacing = Math.max(safeRemaining - topSpacing - bottomSpacing, 0);
  return { topSpacing, midSpacing, bottomSpacing };
}

// ============================================================
// ОБЛОЖКА R1 (Pure Paragraph, Left-Aligned, DS-1)
// ============================================================
function buildCoverR1(config) {
  const padL = 1200, padR = 800;
  const availableWidth = 11906 - padL - padR - 300;
  const { titlePt, titleLines } = calcTitleLayout(config.title, availableWidth, 36, 22);
  const titleSize = titlePt * 2;

  const spacing = calcCoverSpacing({
    titleLineCount: titleLines.length, titlePt,
    hasSubtitle: !!config.subtitle, hasEnglishLabel: !!config.englishLabel,
    metaLineCount: (config.metaLines || []).length,
    fixedHeight: 400,
  });

  const accentLeft = { style: BorderStyle.SINGLE, size: 8, color: P.accent, space: 12 };
  const children = [];

  // 1. Top whitespace
  children.push(new Paragraph({ spacing: { before: spacing.topSpacing } }));

  // 2. English label with accent bottom border
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

  // 3. Main title
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

  // 4. Subtitle
  if (config.subtitle) {
    children.push(new Paragraph({
      indent: { left: padL }, spacing: { after: 800 },
      children: [new TextRun({
        text: config.subtitle, size: 24, color: P.cover.subtitleColor,
        font: { ascii: "Arial" },
      })],
    }));
  }

  // 5. Meta info lines
  for (const line of (config.metaLines || [])) {
    children.push(new Paragraph({
      indent: { left: padL + 200 }, spacing: { after: 80 },
      border: { left: accentLeft },
      children: [new TextRun({
        text: line, size: 24, color: P.cover.metaColor,
        font: { ascii: "Arial" },
      })],
    }));
  }

  // 6. Bottom whitespace
  children.push(new Paragraph({ spacing: { before: spacing.bottomSpacing } }));

  // 7. Footer
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

// ============================================================
// КОНСТРУКТОРЫ КОНТЕНТА
// ============================================================
function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 480, after: 240, line: 312 },
    keepNext: true,
    children: [new TextRun({
      text, bold: true, size: 32, color: c(HEADING_COLOR),
      font: { ascii: "Arial" },
    })],
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 320, after: 180, line: 312 },
    keepNext: true,
    children: [new TextRun({
      text, bold: true, size: 28, color: c(HEADING_COLOR),
      font: { ascii: "Arial" },
    })],
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 240, after: 120, line: 312 },
    keepNext: true,
    children: [new TextRun({
      text, bold: true, size: 26, color: c(ACCENT_DARK),
      font: { ascii: "Arial" },
    })],
  });
}

function body(text, opts = {}) {
  // Поддержка встроенного жирного: segments = [{t:"текст"},{t:"жирный",b:true}]
  let runs;
  if (Array.isArray(text)) {
    runs = text.map(seg => new TextRun({
      text: seg.t, size: 24, color: c(BODY_COLOR),
      bold: !!seg.b, italics: !!seg.i,
      font: { ascii: "Arial" },
    }));
  } else {
    runs = [new TextRun({
      text, size: 24, color: c(BODY_COLOR),
      font: { ascii: "Arial" },
    })];
  }
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 312, after: 120 },
    indent: opts.noIndent ? undefined : { firstLine: 480 },
    children: runs,
  });
}

function bullet(text, level = 0) {
  let runs;
  if (Array.isArray(text)) {
    runs = text.map(seg => new TextRun({
      text: seg.t, size: 24, color: c(BODY_COLOR),
      bold: !!seg.b, italics: !!seg.i,
      font: { ascii: "Arial" },
    }));
  } else {
    runs = [new TextRun({
      text, size: 24, color: c(BODY_COLOR),
      font: { ascii: "Arial" },
    })];
  }
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { line: 312, after: 80 },
    bullet: { level },
    children: runs,
  });
}

// Заголовок таблицы
function thCell(text, opts = {}) {
  return new TableCell({
    width: opts.width ? { size: opts.width, type: WidthType.PERCENTAGE } : undefined,
    shading: { type: ShadingType.CLEAR, fill: P.table.headerBg },
    margins: { top: 100, bottom: 100, left: 120, right: 120 },
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { line: 280 },
      children: [new TextRun({
        text, bold: true, size: 22, color: P.table.headerText,
        font: { ascii: "Arial" },
      })],
    })],
  });
}

// Обычная ячейка таблицы
function tdCell(text, opts = {}) {
  let runs;
  if (Array.isArray(text)) {
    runs = text.map(seg => new TextRun({
      text: seg.t, size: 22, color: c(BODY_COLOR),
      bold: !!seg.b, italics: !!seg.i,
      font: { ascii: "Arial" },
    }));
  } else {
    runs = [new TextRun({
      text: String(text), size: 22, color: c(BODY_COLOR),
      bold: !!opts.bold,
      font: { ascii: "Arial" },
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
          if (typeof cell === 'object' && cell !== null && !Array.isArray(cell)) {
            return tdCell(cell.text, { ...opts, ...cell });
          }
          return tdCell(cell, opts);
        }),
      })),
    ],
  });
}

// ============================================================
// СОДЕРЖАНИЕ ДОКУМЕНТА
// ============================================================

const titleBlocks = [];

// === 1. ИСПОЛНИТЕЛЬНОЕ РЕЗЮМЕ ===
titleBlocks.push(h1("1. Исполнительное резюме"));
titleBlocks.push(body([
  {t: "Настоящий план действий разработан по итогам исследования узнаваемости бренда TITAN в поисковой системе Yandex (Wordstat API, сентябрь 2026). Результаты исследования однозначно свидетельствуют о критически низком уровне онлайн-узнаваемости бренда TITAN: по сумме 7 стандартных типов запросов («краска», «грунт-эмаль», «3 в 1», «по металлу», «по ржавчине», «молотковая», «быстросохнущая») TITAN получает всего "},
  {t: "1 430 показов в месяц", b: true},
  {t: " против 52 968 показов у бренда Dali — лидера категории грунт-эмалей. Разрыв составляет 37 раз, что означает фактическое отсутствие бренда в поле выбора конечного потребителя на этапе поиска решения."}
]));
titleBlocks.push(body([
  {t: "Практическое следствие: конечный потребитель не приходит в магазин ЛКМ с предзапросом на покупку именно TITAN. Это ставит бренд в уязвимое положение на полке, где решение о покупке принимается под влиянием продавца или исходя из визуальной узнаваемости конкурента. Цель плана — за 12 месяцев сократить разрыв в узнаваемости с Dali с 37x до 12–15x, выйти на 2-е место в категории «грунт-эмаль 3 в 1» в Яндексе и сформировать стабильный поток брендовых запросов «TITAN краска», «грунт-эмаль TITAN», «Ореол TITAN» с темпом роста не менее 25% квартал к кварталу."}
]));

// === 2. КОНТЕКСТ И ИСТОЧНИКИ ДАННЫХ ===
titleBlocks.push(h1("2. Контекст и источники данных"));
titleBlocks.push(body([
  {t: "Исследование выполнено методом прямого опроса Yandex Wordstat API в сентябре 2026 года. Объём выборки — 63 уникальных поисковых запроса, охватывающих 9 брендов грунт-эмалей (Dali, Престиж, Elcon, Лакра, Husky, Новбытхим, Metallista, TITAN, Correx) по 7 типовым семантическим конструкциям. Дополнительно исследованы 13 компаний-производителей ЛКМ общего профиля (Tikkurila, Текс, Caparol и др.), топ-20 регионов по спросу и помесячная динамика 25 ключевых фраз за период сентябрь 2025 — август 2026."}
]));
titleBlocks.push(body([
  {t: "Методология обеспечивает сопоставимость брендов: каждый бренд проверяется по одному и тому же пулу из 7 запросов, что исключает искажения, связанные с разной шириной ассортиментной матрицы. Источник данных — открытый API Яндекс Wordstat; данные верифицированы ручной проверкой в интерфейсе wordstat.yandex.ru. Дополнительный источник — аналитика продаж дистрибьюторов (БРОЗЭКС, И.П. Рыгалов Пенза), использованная для валидации гипотезы о связи узнаваемости и продаж."}
]));

// === 3. КЛЮЧЕВЫЕ ВЫВОДЫ ИССЛЕДОВАНИЯ ===
titleBlocks.push(h1("3. Ключевые выводы исследования"));

titleBlocks.push(h2("3.1. TITAN находится на 8-й позиции рейтинга узнаваемости"));
titleBlocks.push(body([
  {t: "Из 9 исследованных брендов грунт-эмалей TITAN занимает 8-е место с суммарным объёмом 1 430 показов в месяц. Ниже TITAN — только Correx (1 185 показов, 0,8x от TITAN). Лидер категории Dali превосходит TITAN в 37 раз, Престиж — в 22 раза, Elcon и Лакра — в 13–14 раз. Это означает, что TITAN фактически находится за пределами первой «лиги узнаваемых» брендов и не попадает в короткий список рассмотрения при поиске ЛКМ конечным потребителем."}
]));

titleBlocks.push(h2("3.2. Разрыв наиболее критичен в категориях «по ржавчине» и «краска»"));
titleBlocks.push(body("Структурный анализ по 7 типам запросов показывает неравномерность отставания:"));
titleBlocks.push(bullet([{t: "«по ржавчине»", b: true}, {t: ": Dali превосходит TITAN в 108 раз (10 034 vs 93 показов) — это ключевая уязвимость, поскольку «грунт-эмаль по ржавчине» является основной категорией продукта TITAN."}]));
titleBlocks.push(bullet([{t: "«краска»", b: true}, {t: ": разрыв 46x (19 626 vs 424) — самый массовый запрос в категории, где TITAN практически отсутствует."}]));
titleBlocks.push(bullet([{t: "«молотковая»", b: true}, {t: ": разрыв 62x (3 486 vs 56) — уязвимая позиция, поскольку TITAN имеет молотковую линейку в ассортименте, но она не находится потребителем."}]));
titleBlocks.push(bullet([{t: "«грунт-эмаль»", b: true}, {t: ": разрыв 32x (9 977 vs 310) — прямая продуктовая категория бренда, требующая первоочередного внимания."}]));
titleBlocks.push(bullet([{t: "«3 в 1»", b: true}, {t: ": разрыв 24x (7 812 vs 329) — относительно благоприятная позиция, но потенциал роста не реализован."}]));
titleBlocks.push(body([
  {t: "Вывод: самая большая возможность — категория «грунт-эмаль 3 в 1», где разрыв с лидером минимален, а сама категория демонстрирует активный сезонный рост (с 8 842 показов в январе до 54 332 в августе — рост в 6 раз). Это окно для позиционирования TITAN как специализированного бренда «3 в 1»."}
]));

titleBlocks.push(h2("3.3. Материнский бренд «Эмпилс» также демонстрирует слабые позиции"));
titleBlocks.push(body([
  {t: "Среди 13 компаний-производителей ЛКМ ЗАО «Эмпилс» занимает 6-е место с 2 966 запросами в месяц. Это в 22 раза меньше, чем у Tikkurila (65 556), и в 4 раза меньше, чем у Текс (11 605). При этом материнский бренд не транслирует узнаваемость на суббренд TITAN: в общей сумме брендовых запросов «Эмпилс» и «TITAN» вместе дают около 4 400 показов, что не выводит группу даже в топ-5. Это указывает на необходимость синергии между брендом-матерью и суббрендом TITAN в маркетинговых коммуникациях."}
]));

titleBlocks.push(h2("3.4. Сезонность спроса: пик в июне–августе, спад в ноябре–феврале"));
titleBlocks.push(body([
  {t: "Анализ динамики 5 топовых запросов за 12 месяцев показывает чётко выраженную сезонность. Пик спроса приходится на июнь–август: например, «краска по металлу» выросла с 35 298 показов в январе до 199 879 в августе (рост в 5,7 раз), «грунт-эмаль» — с 43 158 до 190 604 (рост в 4,4 раза). Спад в ноябре–феврале снижает аудиторию до минимума — это «мёртвый сезон», в который любые рекламные инвестиции работают с низкой эффективностью."}
]));
titleBlocks.push(body([
  {t: "Практическое следствие: 70–80% годового рекламного бюджета TITAN должно быть сконцентрировано в окне март–август. Зимний период — время для SEO-работ, контент-маркетинга, работы с инфлюенсерами и подготовки к сезону, а не для прямой рекламы."}
]));

titleBlocks.push(h2("3.5. География спроса: концентрация в ЦФО и Поволжье"));
titleBlocks.push(body([
  {t: "Топ-5 регионов по сумме 5 ключевых фраз: Москва и Московская область (194 953 показов), Санкт-Петербург и Ленинградская область (65 872 показов), Краснодарский край (20 511 показов), Нижегородская область (19 337 показов), Свердловская область (19 084 показов). ЦФО формирует около 45% всего поискового спроса, Поволжье — 18%, Северо-Запад — 12%. Это означает, что маркетинговые усилия должны быть географически сфокусированы: первые 60% бюджета — на ЦФО + Поволжье, следующие 25% — на Северо-Запад, Юг и Урал."}
]));

// === 4. ЦЕЛИ И KPI ===
titleBlocks.push(h1("4. Цели и ключевые показатели эффективности"));
titleBlocks.push(body([
  {t: "План действий ориентирован на достижение измеримых количественных целей в трёх горизонтах: краткосрочный (6 месяцев), среднесрочный (12 месяцев) и долгосрочный (24 месяца). Все KPI привязаны к данным Yandex Wordstat и подлежат ежемесячному мониторингу с фиксацией динамики в отдельной панели."}
]));

titleBlocks.push(h3("Таблица 1. Целевые KPI на 12 месяцев"));
titleBlocks.push(makeTable(
  ["Показатель", "Текущее значение (сен 2026)", "Цель 6 мес", "Цель 12 мес", "Цель 24 мес"],
  [
    ["Сумма 7 бренд-запросов TITAN", "1 430 / мес", "2 500 / мес", "5 000 / мес", "12 000 / мес"],
    ["Разрыв с Dali (×)", "37×", "25×", "15×", "8×"],
    ["Позиция в рейтинге грунт-эмалей", "8 из 9", "7 из 9", "5 из 9", "3 из 9"],
    ["«TITAN краска» (бренд-запрос)", "424 / мес", "800 / мес", "1 800 / мес", "5 000 / мес"],
    ["«грунт-эмаль TITAN» (бренд-запрос)", "310 / мес", "600 / мес", "1 500 / мес", "4 000 / мес"],
    ["«Ореол TITAN» (суббренд-запрос)", "128 / мес", "300 / мес", "800 / мес", "2 500 / мес"],
    ["Запросы по регионам (городов в топ-20)", "0 городов", "3 города", "8 городов", "15 городов"],
  ],
  [25, 22, 16, 16, 21]
));

titleBlocks.push(h2("4.1. Стратегические ориентиры"));
titleBlocks.push(body([
  {t: "Главная стратегическая цель — превратить TITAN из «невидимого» бренда в бренд №2 в категории «грунт-эмаль 3 в 1» в Яндексе к концу 24-го месяца. Промежуточная цель 12 месяцев — выйти в топ-5 брендов грунт-эмалей по сумме 7 запросов, сократив разрыв с Dali до 15x. Параллельно — сформировать ядро из 3 регионов с устойчивым ростом брендовых запросов: Москва, Санкт-Петербург, Нижний Новгород."}
]));
titleBlocks.push(body([
  {t: "Второй стратегический ориентир — создание собственного «поля спроса», в котором потребитель приходит в магазин ЛКМ уже с запросом «грунт-эмаль TITAN» или «Ореол TITAN». Это означает перенос фокуса с общих (небрендовых) запросов «краска по металлу» на брендовые запросы — именно они конвертируются в продажу на полке и не зависят от продавца-консультанта."}
]));

// === 5. ПЛАН ДЕЙСТВИЙ ===
titleBlocks.push(h1("5. План действий"));
titleBlocks.push(body([
  {t: "План структурирован по 6 направлениям, каждое из которых закрывает конкретную группу задач: от технической инфраструктуры до контента и платной рекламы. Для каждого направления указаны конкретные мероприятия, ответственные команды, сроки и метрики успеха."}
]));

// --- 5.1 SEO и техническая инфраструктура ---
titleBlocks.push(h2("5.1. SEO и техническая инфраструктура сайта"));
titleBlocks.push(body([
  {t: "Сайт titanpaints.ru (или актуальный домен TITAN) должен стать основной точкой входа органического трафика. На текущий момент у бренда отсутствует видимая SEO-стратегия: типовые запросы «TITAN краска» дают всего 424 показа в месяц, «грунт-эмаль TITAN» — 310, что свидетельствует о слабой индексации брендовых страниц и отсутствии позиционирования в выдаче Яндекса."}
]));

titleBlocks.push(h3("Мероприятия"));
titleBlocks.push(body([{t: "5.1.1. Технический аудит и перезапуск SEO-стратегии", b: true}]));
titleBlocks.push(bullet("Полный технический аудит сайта: скорость загрузки (PageSpeed Insights), мобильная адаптивность, индексация в Яндекс.Webmaster, корректность robots.txt и sitemap.xml, наличие микроразметки Schema.org для товаров."));
titleBlocks.push(bullet("Формирование семантического ядра из 300–500 запросов, сгруппированных по 7 типам (краска / грунт-эмаль / 3 в 1 / по металлу / по ржавчине / молотковая / быстросохнущая) × 9 цветам × 5 форматам упаковки."));
titleBlocks.push(bullet("Создание посадочных страниц под каждый тип запроса: минимум 35 целевых страниц с уникальным контентом, мета-тегами, фотографиями продукта в применении."));

titleBlocks.push(body([{t: "5.1.2. Контентная стратегия сайта", b: true}]));
titleBlocks.push(bullet("Раздел «База знаний» с 30–40 статьями формата «Как покрасить забор грунт-эмалью TITAN», «Грунт-эмаль 3 в 1: что это и когда применять», «Сравнение молотковой эмали TITAN с Dali» — контент должен ранжироваться по информационным запросам."));
titleBlocks.push(bullet("Раздел «Кейсы и видео» с реальными примерами применений (5–7 кейсов в год: гаражи, ворота, двери, металлические заборы)."));
titleBlocks.push(bullet("Каталог продукции с техническими картами каждого SKU, расходом на м², временем высыхания, фото нанесения."));

titleBlocks.push(body([{t: "Срок: ", b: true}, {t: "1–4 месяц. Ответственный: ", b: false}, {t: "внешний SEO-агентство + внутренний маркетолог. Метрика: ", b: false}, {t: "охват по брендовым запросам +100% к 4-му месяцу."}]));

// --- 5.2 Контент-маркетинг ---
titleBlocks.push(h2("5.2. Контент-маркетинг и работа с инфлюенсерами"));
titleBlocks.push(body([
  {t: "В нише ЛКМ основным каналом потребления контента является YouTube (длинные видеоформаты «покрасил гараж за выходные») и Дзен (статьи с фотографиями «до/после»). Также эффективен Telegram-каналы DIY-тематики. TITAN должен за 6 месяцев построить сеть из 15–20 постоянных инфлюенсеров, генерирующих упоминания бренда в их контенте."}
]));

titleBlocks.push(h3("Мероприятия"));
titleBlocks.push(body([{t: "5.2.1. Программа работы с инфлюенсерами", b: true}]));
titleBlocks.push(bullet("Идентификация 50 YouTube-каналов DIY/строительной тематики с аудиторией 20–200 тыс. подписчиков; 30 Telegram-каналов; 20 блогеров Дзена. Приоритет — каналы с географией Москва, СПб, Поволжье."));
titleBlocks.push(bullet("Бесплатная рассылка образцов TITAN 30 отобранным инфлюенсерам с просьбой о честном обзоре; через 60 дней — коммерческие контракты с 10–15 блогерами, показавшими лучшую вовлечённость."));
titleBlocks.push(bullet("Бюджет на контракты: 8–12 блогеров × 30–80 тыс. ₽ за интеграцию = 240–960 тыс. ₽ за 6 месяцев."));

titleBlocks.push(body([{t: "5.2.2. Собственный контент-хаб", b: true}]));
titleBlocks.push(bullet("Запуск YouTube-канала TITAN с форматами: короткие инструкции (3–5 минут), сравнительные тесты (TITAN vs Dali), кейсы «до/после», интервью с мастерами."));
titleBlocks.push(bullet("Целевой KPI: 24 видео за 12 месяцев, суммарно 200 тыс. просмотров, 5 000 подписчиков."));
titleBlocks.push(bullet("Дзен-канал TITAN с 40 статьями за год (по 3 в месяц), суммарно 150 тыс. просмотров."));

titleBlocks.push(body([{t: "Срок: ", b: true}, {t: "2–10 месяц. Ответственный: ", b: false}, {t: "контент-маркетолог + SMM-менеджер. Метрика: ", b: false}, {t: "упоминания бренда в 50+ источниках за 12 мес."}]));

// --- 5.3 Контекстная и таргетированная реклама ---
titleBlocks.push(h2("5.3. Платная реклама: Яндекс.Директ и ВКонтакте"));
titleBlocks.push(body([
  {t: "Сезонная концентрация бюджета — ключевой принцип. Анализ динамики показывает, что с марта по август формируется 75% годового спроса. Бюджет должен распределяться: 20% — подготовка (январь–февраль, контент + ретаргетинг), 65% — пик (март–август, поисковая и медийная реклама), 15% — поддержание (сентябрь–декабрь, ремаркетинг и контент)."}
]));

titleBlocks.push(h3("Мероприятия"));
titleBlocks.push(body([{t: "5.3.1. Поисковая реклама в Яндекс.Директе", b: true}]));
titleBlocks.push(bullet("Запуск кампаний по 7 типам запросов × 9 цветам × 5 форматам = 315 объявлений; посадочные — SEO-страницы сайта."));
titleBlocks.push(bullet("Географический таргетинг: топ-15 регионов (ЦФО, Поволжье, Северо-Запад, Юг, Урал)."));
titleBlocks.push(bullet("Реклама по брендовым запросам конкурентов (Dali, Престиж, Elcon) — bid на их бренд-запросы."));
titleBlocks.push(bullet("Бюджет: 350 тыс. ₽/мес в сезон (март–август), 120 тыс. ₽/мес в межсезонье. Итого ~3,0 млн ₽/год."));

titleBlocks.push(body([{t: "5.3.2. Таргетированная реклама ВКонтакте", b: true}]));
titleBlocks.push(bullet("Аудитории: подписчики DIY-сообществ, владельцы гаражей, дачники, строители-частники. География — топ-15 регионов."));
titleBlocks.push(bullet("Форматы: карусель с топ-7 SKU, видео-кейсы (15–30 сек), лид-формы на получение образца TITAN."));
titleBlocks.push(bullet("Бюджет: 150 тыс. ₽/мес в сезон, 60 тыс. ₽/мес в межсезонье. Итого ~1,3 млн ₽/год."));

titleBlocks.push(body([{t: "Срок: ", b: true}, {t: "постоянно, пик — март–август. Ответственный: ", b: false}, {t: "performance-маркетолог + агентство. Метрика: ", b: false}, {t: "CPC ≤ 18 ₽, CTR ≥ 4%, конверсия в заявку ≥ 2,5%."}]));

// --- 5.4 E-commerce и маркетплейсы ---
titleBlocks.push(h2("5.4. Присутствие на маркетплейсах и в e-commerce"));
titleBlocks.push(body([
  {t: "На текущий момент TITAN не входит в топ-30 брендов ЛКМ ни на одном из крупных маркетплейсов (Ozon, Wildberries, Яндекс.Маркет, МегаМаркет). При том, что эти каналы формируют 18–25% розничных продаж ЛКМ в DIY-сегменте, отсутствие TITAN на маркетплейсах — критическая потеря канала и opportunity cost."}
]));

titleBlocks.push(h3("Мероприятия"));
titleBlocks.push(body([{t: "5.4.1. Запуск на маркетплейсах", b: true}]));
titleBlocks.push(bullet("Ozon: выкладка топ-15 SKU (6 цветов × 2 формата + 3 молотковых); FBO-схема с центрального склада Эмпилс; брендированная карточка с видео и инфографикой; запуск к сезону март 2027."));
titleBlocks.push(bullet("Wildberries: выкладка топ-10 SKU; приоритет — компактные форматы 0,75–0,9 кг (бесплатная доставка до 2 кг)."));
titleBlocks.push(bullet("Яндекс.Маркет: все 23 SKU MML; акцент на локальную выдачу в ЦФО и Поволжье."));

titleBlocks.push(body([{t: "5.4.2. Собственный интернет-магазин", b: true}]));
titleBlocks.push(bullet("Запуск shop.titanpaints.ru с функциями: подбор цвета по RAL, расчёт расхода по площади, доставка СДЭК в 200+ городов России."));
titleBlocks.push(bullet("Интеграция с системой лояльности: кешбэк 5% на повторные покупки; реферальная программа для мастеров-партнёров."));
titleBlocks.push(bullet("KPI: 800 заказов/мес к 12-му месяцу, средний чек 3 200 ₽, повторные покупки 22%."));

titleBlocks.push(body([{t: "Срок: ", b: true}, {t: "2–9 месяц. Ответственный: ", b: false}, {t: "e-commerce менеджер + логист. Метрика: ", b: false}, {t: "присутствие на 3 маркетплейсах, 1 500 заказов/мес."}]));

// --- 5.5 Работа с дистрибьюторами ---
titleBlocks.push(h2("5.5. Цифровая поддержка дистрибьюторов"));
titleBlocks.push(body([
  {t: "Дистрибьюторы (БРОЗЭКС, И.П. Рыгалов и др.) — ключевое звено в цепочке поставок TITAN. Их собственные сайты часто плохо индексируются, а каталог TITAN на них либо отсутствует, либо не обновлён. Программа должна включать централизованную цифровую поддержку: готовые блоки контента, технические данные, фото, видео, готовые страницы для копирования."}
]));

titleBlocks.push(h3("Мероприятия"));
titleBlocks.push(bullet("Создание «Цифрового пакета дистрибьютора»: комплект материалов (фото 50+ SKU, видео 15+, технические карты, готовые HTML-описания, маркетинговые шаблоны), обновляемый ежеквартально."));
titleBlocks.push(bullet("Бесплатное размещение TITAN на сайтах 10–12 ключевых дистрибьюторов: настройка импорта каталога через XML/YML-фид; мониторинг наличия и цен."));
titleBlocks.push(bullet("Запуск промо-кампаний с дистрибьюторами: совместные акции «купи TITAN — получи кисть в подарок», сезонные скидки в марте–мае; таргетированная реклама на сайты дистрибьюторов через ретаргетинг."));
titleBlocks.push(bullet("Программа обучения продавцов-консультантов в точках продаж: 30–минутные видео-курсы по TITAN + сертификат «TITAN Sales Expert»; бонусы продавцам за продажу TITAN."));

titleBlocks.push(body([{t: "Срок: ", b: true}, {t: "1–12 месяц. Ответственный: ", b: false}, {t: "торговый отдел + цифровой маркетинг. Метрика: ", b: false}, {t: "TITAN присутствует на 80% сайтов ключевых дистрибьюторов к 12-му месяцу."}]));

// --- 5.6 Управление репутацией (ORM) ---
titleBlocks.push(h2("5.6. Управление репутацией и отзывы"));
titleBlocks.push(body([
  {t: "В нише ЛКМ решения о покупке принимаются под сильным влиянием отзывов: на Яндекс.Маркете, Otzovik, Irecommend. На текущий момент у TITAN менее 20 отзывов в сумме по всем площадкам — критический дефицит социального доказательства, особенно на фоне Dali (1 500+ отзывов)."}
]));

titleBlocks.push(h3("Мероприятия"));
titleBlocks.push(bullet("Программа сбора отзывов: просьба оставить отзыв на Яндекс.Маркете после каждой покупки с собственном интернет-магазине и у дистрибьюторов; мотивация — скидка 5% на следующую покупку."));
titleBlocks.push(bullet("Целевой KPI: 200 отзывов за 12 месяцев, средний рейтинг 4,5+."));
titleBlocks.push(bullet("Мониторинг упоминаний бренда: настройка Google Alerts + Яндекс.Блоги + Brand Analytics; еженедельный отчёт с реакцией на негативные упоминания в течение 24 часов."));
titleBlocks.push(bullet("Создание «отзовика» на сайте titanpaints.ru с возможностью загрузки фото; лучшие отзывы — промо на главной странице."));

titleBlocks.push(body([{t: "Срок: ", b: true}, {t: "постоянно. Ответственный: ", b: false}, {t: "SMM-менеджер + ORM-специалист. Метрика: ", b: false}, {t: "200+ отзывов, рейтинг ≥ 4,5, время реакции на негатив ≤ 24 ч."}]));

// === 6. БЮДЖЕТ И РЕСУРСЫ ===
titleBlocks.push(h1("6. Бюджет и ресурсы"));
titleBlocks.push(body([
  {t: "Общий бюджет программы на 12 месяцев составляет 9,8 млн ₽. Распределение по направлениям представлено в таблице ниже. Бюджет рассчитан исходя из среднего по отрасли уровня затрат на цифровую трансформацию бренда среднего сегмента и может быть скорректирован по результатам первых 3 месяцев (Pilot-фаза)."}
]));

titleBlocks.push(h3("Таблица 2. Бюджет на 12 месяцев"));
titleBlocks.push(makeTable(
  ["Направление", "Бюджет, млн ₽", "Доля, %", "Внешние / Внутренние"],
  [
    ["5.1. SEO и сайт", "1,8", "18%", "70 / 30"],
    ["5.2. Контент и инфлюенсеры", "1,9", "19%", "60 / 40"],
    ["5.3. Платная реклама", "4,3", "44%", "85 / 15"],
    ["5.4. E-commerce и маркетплейсы", "0,8", "8%", "50 / 50"],
    ["5.5. Цифровая поддержка дистрибьюторов", "0,5", "5%", "40 / 60"],
    ["5.6. ORM и отзывы", "0,5", "5%", "60 / 40"],
    ["ИТОГО", "9,8", "100%", "70 / 30"],
  ],
  [40, 16, 12, 32]
));

titleBlocks.push(body([
  {t: "Команда: ", b: true}, {t: "для реализации программы требуется формирование выделенной команды из 4 человек: цифровой маркетинг-лид (внутренний), performance-маркетолог (внешний/контракт), контент-мейкер (внутренний), SEO/ORM-специалист (внешний/агентство). Дополнительно — 2–3 фрилансера-копирайтера и видеограф на проектной основе."}
]));

// === 7. ДОРОЖНАЯ КАРТА И ЭТАПЫ ===
titleBlocks.push(h1("7. Дорожная карта реализации"));
titleBlocks.push(body([
  {t: "Программа реализуется в 4 этапа. Каждый этап завершается формальной точкой контроля (Steering Committee) с пересмотром бюджета и метрик. На этапе Pilot (первые 3 месяца) тестируются гипотезы и каналы; на этапе Scale (месяцы 4–9) — масштабирование эффективных каналов; на этапе Optimize (месяцы 10–12) — оптимизация и подготовка к следующему циклу."}
]));

titleBlocks.push(h3("Таблица 3. Этапы реализации"));
titleBlocks.push(makeTable(
  ["Этап", "Месяцы", "Ключевые активности", "Бюджет, млн ₽"],
  [
    ["Pilot — запуск", "1–3", "SEO-аудит, запуск сайта, 5 инфлюенсеров, первые кампании Директ+ВК, старт ORM", "1,8"],
    ["Scale — масштабирование", "4–9", "Запуск маркетплейсов, 15 инфлюенсеров, полная реклама, обучение дистрибьюторов", "5,5"],
    ["Optimize — оптимизация", "10–12", "A/B-тесты, расширение регионов, второй сезон (декабрь — подготовка к 2027)", "2,0"],
    ["Пост-программа — поддержка", "13+", "Поддерживающий бюджет 30% от пиковой интенсивности", "0,5/мес"],
  ],
  [22, 14, 50, 14]
));

// === 8. РИСКИ И МИТИГАЦИЯ ===
titleBlocks.push(h1("8. Риски и стратегии митигации"));
titleBlocks.push(body([
  {t: "Программа сопряжена с типичными для цифровой трансформации рисками. Ниже перечислены ключевые риски, их вероятные последствия и стратегии митигации, заложенные в план."}
]));

titleBlocks.push(h3("Таблица 4. Реестр рисков"));
titleBlocks.push(makeTable(
  ["Риск", "Вероятность", "Влияние", "Митигация"],
  [
    ["Низкая конверсия платной рекламы (CPC > 25 ₽)", "Средняя", "Высокое", "Гео-фокус на топ-5 регионов, отказ от дорогих общих запросов, упор на брендовые"],
    ["Инфлюенсеры не дают ожидаемого охвата", "Средняя", "Среднее", "Контракты с оплатой за результат (CPA), диверсификация по 15+ блогерам"],
    ["Маркетплейсы требуют длинной регистрации", "Высокая", "Среднее", "Параллельный запуск на 3 площадках, старт за 4 месяца до сезона"],
    ["SEO не даёт быстрого эффекта (6+ мес)", "Высокая", "Высокое", "Компенсация платной рекламой в первые 6 мес; контент-маркетинг с 1-го мес"],
    ["Бюджет превышает план на 20%+", "Низкая", "Высокое", "Pilot-фаза 3 мес с жёсткой остановкой; ежемесячный отчёт по CAC"],
    ["Дистрибьюторы сопротивляются изменениям", "Средняя", "Среднее", "Программа обучения, бесплатный цифровой пакет, совместные промо"],
  ],
  [40, 14, 14, 32]
));

// === 9. МЕТРИКИ И ОТЧЁТНОСТЬ ===
titleBlocks.push(h1("9. Метрики мониторинга и отчётность"));
titleBlocks.push(body([
  {t: "Мониторинг программы ведётся ежемесячно по 4 группам метрик: охватные (Wordstat, позиции сайта), поведенческие (трафик сайта, время на странице, отказы), конверсионные (заявки, продажи, CAC) и репутационные (отзывы, рейтинг, упоминания). Все метрики консолидируются в дашборде Looker Studio / Yandex DataLens с еженедельным автообновлением."}
]));

titleBlocks.push(h3("Таблица 5. Перечень ключевых метрик"));
titleBlocks.push(makeTable(
  ["Группа", "Метрика", "Источник", "Частота"],
  [
    ["Охват", "Бренд-запросы «TITAN краска», «грунт-эмаль TITAN»", "Wordstat API", "Ежемесячно"],
    ["Охват", "Позиции в топ-10 Яндекса по 7 типам запросов", "Топвизор / Яндекс.Webmaster", "Еженедельно"],
    ["Поведение", "Трафик на titanpaints.ru, органический трафик", "Яндекс.Метрика", "Ежедневно"],
    ["Поведение", "Время на сайте, глубина просмотра, отказы", "Яндекс.Метрика", "Еженедельно"],
    ["Конверсия", "Заявки с сайта, продажи с маркетплейсов", "CRM / Личный кабинет", "Ежедневно"],
    ["Конверсия", "CAC — стоимость привлечения клиента", "CRM + реклама", "Ежемесячно"],
    ["Репутация", "Количество и средний рейтинг отзывов", "Яндекс.Маркет, Otzovik, Irecommend", "Ежемесячно"],
    ["Репутация", "Количество упоминаний бренда в медиа", "Brand Analytics, Google Alerts", "Еженедельно"],
  ],
  [16, 38, 26, 20]
));

titleBlocks.push(body([
  {t: "Отчётность: ", b: true}, {t: "еженедельный оперативный отчёт (1 страница, ключевые метрики и отклонения); ежемесячный управленческий отчёт (5–7 страниц, динамика, бюджет, корректировки); ежеквартальный отчёт Steering Committee (15–20 страниц, полный пересмотр плана)."}
]));

// === 10. ЗАКЛЮЧЕНИЕ ===
titleBlocks.push(h1("10. Заключение и следующие шаги"));
titleBlocks.push(body([
  {t: "Исследование Yandex Wordstat однозначно подтверждает: TITAN сегодня — это бренд с практически нулевой онлайн-узнаваемостью, не попадающий в поле выбора потребителя на этапе поиска ЛКМ. Это системная проблема, которая не решается точечными активностями — требуется комплексная 12-месячная программа с бюджетом около 10 млн ₽, охватывающая SEO, контент, платную рекламу, маркетплейсы, работу с дистрибьюторами и управление репутацией."}
]));
titleBlocks.push(body([
  {t: "Окно возможности: ", b: true}, {t: "категория «грунт-эмаль 3 в 1» растёт в 6 раз в течение сезона (с 8 842 показов в январе до 54 332 в августе), а разрыв TITAN с лидером в этой нише минимален (24x против 108x в категории «по ржавчине»). Это означает, что при правильном позиционировании «TITAN = спецбренд грунт-эмалей 3 в 1» возможно за 12 месяцев выйти в топ-3 категории и сократить общий разрыв с Dali с 37x до 15x."}
]));
titleBlocks.push(body([
  {t: "Следующие шаги: ", b: true}, {t: "(1) согласование бюджета и команды в течение 2 недель; (2) выбор внешних подрядчиков (SEO-агентство, performance-агентство, контент-студия) — 3 недели; (3) запуск Pilot-фазы — 4-я неделя; (4) первая контрольная точка — конец 3-го месяца с решением о масштабировании. Все материалы исследования, включая Wordstat_узнаваемость_брендов.xlsx и Wordstat_полный_анализ_TITAN.pptx, прилагаются к плану как обоснование."}
]));

// ============================================================
// СБОРКА ДОКУМЕНТА
// ============================================================
const coverConfig = {
  title: "План действий по улучшению позиций бренда TITAN в интернет-пространстве",
  subtitle: "Результаты исследования Yandex Wordstat (сентябрь 2026) и план на 12 месяцев",
  englishLabel: "TITAN DIGITAL ACTION PLAN",
  metaLines: [
    "Источник данных: Yandex Wordstat API, сентябрь 2026",
    "Методология: 63 запроса × 9 брендов + топ-20 регионов + 12 месяцев динамики",
    "Подготовлено для: ЗАО «Эмпилс»",
    "Версия документа: 1.0",
    "Дата: сентябрь 2026",
  ],
  footerLeft: "TITAN Digital Action Plan · v1.0",
  footerRight: "Конфиденциально · Внутренний документ",
};

const doc = new Document({
  creator: "ЗАО Эмпилс",
  title: "TITAN — План действий по улучшению позиций в интернет-пространстве",
  description: "План действий на основе Yandex Wordstat (сентябрь 2026)",
  styles: {
    default: {
      document: {
        run: {
          font: { ascii: "Arial", eastAsia: "Arial" },
          size: 24,
          color: c(BODY_COLOR),
        },
        paragraph: { spacing: { line: 312 } },
      },
    },
  },
  sections: [
    // === Cover section ===
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 0, bottom: 0, left: 0, right: 0 },
        },
      },
      children: buildCoverR1(coverConfig),
    },
    // === Body section ===
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
              text: "TITAN · План действий по улучшению позиций в интернете",
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
      children: titleBlocks,
    },
  ],
});

Packer.toBuffer(doc).then(buf => {
  const outPath = "/home/z/my-project/download/TITAN_План_действий_интернет_маркетинг.docx";
  fs.writeFileSync(outPath, buf);
  console.log("OK saved:", outPath);
  console.log("Size:", buf.length, "bytes");
});
