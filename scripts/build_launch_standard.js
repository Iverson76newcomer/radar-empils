/**
 * Стандарт вывода новинки на рынок — ЗАО «Эмпилс»
 * С визуализацией этапов (стрелочные диаграммы) и матрицей ролей
 */
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, PageNumber, PageBreak,
  BorderStyle, ShadingType, WidthType, TableLayoutType, NumberFormat,
  SectionType,
} = require("docx");
const fs = require("fs");

// Палитра DS-1 (Deep Sea)
const P = {
  bg: "0B1C2C",
  accent: "529286",
  cover: { titleColor: "FFFFFF", subtitleColor: "B0B8C0", metaColor: "90989F", footerColor: "687078" },
  table: { headerBg: "529286", headerText: "FFFFFF", innerLine: "BECFCC", surface: "E8ECEB" },
};

const BODY_COLOR = "1A2B3D";
const HEADING_COLOR = "0B1C2C";
const ACCENT_DARK = "3D6B62";
const ACCENT_ORANGE = "F47629";
const COLOR_GREEN = "16A34A";
const COLOR_RED = "C0392B";
const COLOR_YELLOW = "D97706";
const COLOR_BLUE = "1E3A5F";
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

// === Cover R1 helpers ===
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

function calcTitleLayout(title, maxWidthTwips, preferredPt = 32, minPt = 22) {
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

function calcCoverSpacing({ titleLineCount = 1, titlePt = 32, hasSubtitle = false, hasEnglishLabel = false, metaLineCount = 0, fixedHeight = 800, pageHeight = 16838 }) {
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
  const { titlePt, titleLines } = calcTitleLayout(config.title, availableWidth, 30, 22);
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
          if (typeof cell === 'object' && cell !== null && !Array.isArray(cell)) {
            return tdCell(cell.text, { ...opts, ...cell });
          }
          return tdCell(cell, opts);
        }),
      })),
    ],
  });
}

// === Стрелочная визуализация этапов (через таблицу) ===
function buildStageDiagram(stages) {
  // stages = [{num: 1, name: 'Название', owner: 'Ответственный', days: '7 дней'}]
  // Создаём горизонтальную таблицу с этапами и стрелками между ними
  const rows = [];

  // Первая строка: номера этапов с фоном
  const numRow = new TableRow({
    cantSplit: true,
    children: stages.flatMap((s, i) => {
      const cells = [];
      // Ячейка с номером этапа
      cells.push(new TableCell({
        width: { size: 20, type: WidthType.PERCENTAGE },
        shading: { type: ShadingType.CLEAR, fill: ACCENT_ORANGE },
        margins: { top: 100, bottom: 100, left: 60, right: 60 },
        verticalAlign: "center",
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            text: `ЭТАП ${s.num}`, bold: true, size: 20, color: "FFFFFF", font: { ascii: "Arial" },
          })],
        })],
      }));
      // Стрелка (кроме последнего)
      if (i < stages.length - 1) {
        cells.push(new TableCell({
          width: { size: 5, type: WidthType.PERCENTAGE },
          margins: { top: 100, bottom: 100, left: 30, right: 30 },
          verticalAlign: "center",
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({
              text: "→", bold: true, size: 28, color: ACCENT_ORANGE, font: { ascii: "Arial" },
            })],
          })],
        }));
      }
      return cells;
    }),
  });
  rows.push(numRow);

  // Вторая строка: названия этапов
  const nameRow = new TableRow({
    cantSplit: true,
    children: stages.flatMap((s, i) => {
      const cells = [];
      cells.push(new TableCell({
        width: { size: 20, type: WidthType.PERCENTAGE },
        shading: { type: ShadingType.CLEAR, fill: "F8F9FA" },
        margins: { top: 80, bottom: 80, left: 60, right: 60 },
        verticalAlign: "center",
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            text: s.name, bold: true, size: 20, color: c(HEADING_COLOR), font: { ascii: "Arial" },
          })],
        })],
      }));
      if (i < stages.length - 1) {
        cells.push(new TableCell({
          width: { size: 5, type: WidthType.PERCENTAGE },
          margins: { top: 80, bottom: 80 },
          children: [new Paragraph({ children: [new TextRun({ text: "" })] })],
        }));
      }
      return cells;
    }),
  });
  rows.push(nameRow);

  // Третья строка: ответственный
  const ownerRow = new TableRow({
    cantSplit: true,
    children: stages.flatMap((s, i) => {
      const cells = [];
      cells.push(new TableCell({
        width: { size: 20, type: WidthType.PERCENTAGE },
        shading: { type: ShadingType.CLEAR, fill: "E8F5E9" },
        margins: { top: 80, bottom: 80, left: 60, right: 60 },
        verticalAlign: "center",
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "👤 ", size: 18 }),
            new TextRun({
              text: s.owner, size: 18, color: c(BODY_COLOR), font: { ascii: "Arial" },
            }),
          ],
        })],
      }));
      if (i < stages.length - 1) {
        cells.push(new TableCell({
          width: { size: 5, type: WidthType.PERCENTAGE },
          margins: { top: 80, bottom: 80 },
          children: [new Paragraph({ children: [new TextRun({ text: "" })] })],
        }));
      }
      return cells;
    }),
  });
  rows.push(ownerRow);

  // Четвёртая строка: срок
  const daysRow = new TableRow({
    cantSplit: true,
    children: stages.flatMap((s, i) => {
      const cells = [];
      cells.push(new TableCell({
        width: { size: 20, type: WidthType.PERCENTAGE },
        shading: { type: ShadingType.CLEAR, fill: "FFF3E0" },
        margins: { top: 80, bottom: 80, left: 60, right: 60 },
        verticalAlign: "center",
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "⏱ ", size: 18 }),
            new TextRun({
              text: s.days, bold: true, size: 18, color: c(COLOR_YELLOW), font: { ascii: "Arial" },
            }),
          ],
        })],
      }));
      if (i < stages.length - 1) {
        cells.push(new TableCell({
          width: { size: 5, type: WidthType.PERCENTAGE },
          margins: { top: 80, bottom: 80 },
          children: [new Paragraph({ children: [new TextRun({ text: "" })] })],
        }));
      }
      return cells;
    }),
  });
  rows.push(daysRow);

  const numCols = stages.length * 2 - 1;
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: thinBorders("D0D0D0"),
    rows,
  });
}

// ============================================================
// СОДЕРЖАНИЕ
// ============================================================
const blocks = [];

// === 1. НАЗНАЧЕНИЕ И ОБЛАСТЬ ПРИМЕНЕНИЯ ===
blocks.push(h1("1. Назначение и область применения"));
blocks.push(body([
  {t: "Настоящий стандарт определяет единый порядок действий по выводу новинки на рынок для всех брендов компании «Эмпилс» (TITAN, NEWTONE, ОРЕОЛ, AVANTGARDE и др.). Документ описывает последовательность этапов, роли участников из отдела маркетинга и отдела продаж, а также артефакты (документы, материалы), которые должны быть подготовлены на каждом этапе."},
]));
blocks.push(body([
  {t: "Точка отсчёта: ", b: true},
  {t: "формула продукта разработана и сертифицирована, предприятие готово производить продукцию (технологическая карта утверждена, сырьё закуплено, опытная партия выпущена)."},
]));
blocks.push(body([
  {t: "Финальная точка: ", b: true},
  {t: "новинка представлена целевой аудитории, размещена в точках продаж, получены первые заказы от дистрибьюторов, собрана обратная связь от рынка."},
]));
blocks.push(body([
  {t: "Срок реализации полного цикла: ", b: true},
  {t: "от 8 до 12 недель в зависимости от категории продукта и количества федеральных округов в запуске."},
]));

// === 2. СТРУКТУРА УЧАСТНИКОВ ===
blocks.push(h1("2. Структура участников процесса"));
blocks.push(body("В выводе новинки на рынок участвуют два департамента с чётким разделением зон ответственности:"));

blocks.push(h2("2.1. Отдел маркетинга"));
blocks.push(body([
  {t: "Отвечает за ", b: false},
  {t: "создание спроса", b: true},
  {t: " — стратегическое позиционирование, ценовую политику, маркетинговые материалы, генерацию лидов и подготовку отдела продаж к работе с новинкой."},
]));

blocks.push(h3("Таблица 1. Роли в отделе маркетинга"));
blocks.push(makeTable(
  ["Роль", "Зона ответственности при выводе новинки", "Ключевые артефакты на выходе"],
  [
    [{text: "Директор по маркетингу", bold: true, align: AlignmentType.CENTER},
     "Стратегическое решение о запуске. Утверждение бюджета, позиционирования, ценовой вилки. Согласование календарного плана. Контроль соответствия бренд-стратегии.",
     "Утверждённый календарный план запуска; бюджет маркетинга; бренд-позиционирование"],
    [{text: "Ведущий специалист группы товарной политики", bold: true, align: AlignmentType.CENTER},
     "Главный «движок» запуска. Полностью ведёт направление. Анализирует конкурентное окружение, рассчитывает ценовую политику и прогнозные объёмы, формирует УТП, готовит презентацию новинки, брифы на производство POSM, технические описания. Передаёт вводные данные отделу продаж для генерации спроса.",
     "Competitive-анализ; ценовое обоснование; УТП и позиционирование; презентация новинки; техническое описание; бриф на POSM"],
  ],
  [22, 48, 30]
));

blocks.push(h2("2.2. Отдел продаж"));
blocks.push(body([
  {t: "Отвечает за ", b: false},
  {t: "конвертацию спроса в продажи", b: true},
  {t: " — сбор потребности от региональных менеджеров, обучение эксклюзивных торговых представителей (ЭТП), контроль выкладки в точках продаж, сбор обратной связи с рынка."},
]));

blocks.push(h3("Таблица 2. Иерархия ролей в отделе продаж"));
blocks.push(makeTable(
  ["Уровень", "Роль", "Зона ответственности", "Решения, которые принимает"],
  [
    [{text: "1", bold: true, align: AlignmentType.CENTER},
     {text: "Директор по продажам", bold: true},
     "Общее руководство запуском в каналах продаж. Согласование квот по федеральным округам. Утверждение плана отгрузок новинки. Контроль выполнения плана продаж.",
     "Утверждение квот по ФО; план отгрузок; KPI для региональных директоров"],
    [{text: "2", bold: true, align: AlignmentType.CENTER},
     {text: "Региональный директор (по каждому ФО)", bold: true},
     "Руководство запуском в своём федеральном округе. Распределение квот между областями. Контроль работы региональных менеджеров. Решение конфликтов по ценам и каналам.",
     "Квоты по областям; KPI для региональных менеджеров; локальные промо-акции"],
    [{text: "3", bold: true, align: AlignmentType.CENTER},
     {text: "Региональный менеджер (по каждой области)", bold: true},
     "Прямая работа с ЭТП в области. Обучение ЭТП по новинке. Контроль выкладки в точках продаж. Сбор заказов от ЭТП. Передача обратной связи от рынка наверх.",
     "Заказы от ЭТП; маршрутные листы посещения точек; обратная связь с рынка"],
    [{text: "4", bold: true, align: AlignmentType.CENTER},
     {text: "Эксклюзивный торговый представитель (ЭТП)", bold: true},
     "Прямой контакт с точками продаж (магазинами, дистрибьюторами). Презентация новинки покупателям. Контроль за выкладкой, ценниками, POSM. Сбор первых заказов.",
     "Заказы от точек; фотоверсия выкладки; отчёт о реакциях покупателей"],
  ],
  [6, 22, 42, 30]
));

// === 3. КЛЮЧЕВЫЕ ПРИНЦИПЫ ===
blocks.push(h1("3. Ключевые принципы стандарта"));
blocks.push(body("Стандарт построен на 5 принципах, которые не допускают отклонений при запуске любой новинки:"));

const principles = [
  [{t: "1. Маркетинг создаёт спрос, продажи — конвертируют", b: true}, {t: ". Ведущий специалист товарной политики готовит «упаковку» новинки (анализ, цена, материалы, УТП), но НЕ продаёт напрямую. Отдел продаж продаёт, но НЕ формирует позиционирование. Каждый занимается своей зоной."}],
  [{t: "2. Последовательность этапов обязательна", b: true}, {t: ". Нельзя пропустить Competitive-анализ и сразу начать продажи — это приведёт к ошибке в ценообразовании. Нельзя обучить ЭТП без готовой презентации новинки."}],
  [{t: "3. Передача артефактов — формальная", b: true}, {t: ". Каждый этап заканчивается артефактом (документом, файлом, решением). Артефакт передаётся на следующий этап в формальном виде (файл в общей папке, протокол встречи). «Слово к слову» не считается передачей."}],
  [{t: "4. Сбор обратной связи — с первого дня продаж", b: true}, {t: ". ЭТП начинает собирать реакцию покупателей с первой недели продаж и передаёт её региональному менеджеру. Региональный менеджер — директору. Тактические корректировки возможны с недели 2-й."}],
  [{t: "5. Точки продаж выбираются осознанно", b: true}, {t: ". Новинка выводится сначала в 10-20 пилотных точках на регион, затем масштабируется. Нельзя выложить новинку сразу во все 500+ точек канала — это снизит качество выкладки и контроля."}],
];
for (const principle of principles) {
  const head = principle[0];
  const tail = principle.slice(1);
  blocks.push(bullet([head, ...tail]));
}

// === 4. КАРТА ПРОЦЕССА ===
blocks.push(h1("4. Карта процесса — 8 этапов вывода новинки на рынок"));
blocks.push(body([
  {t: "Процесс разделён на ", b: false},
  {t: "8 последовательных этапов", b: true},
  {t: ". Каждый этап имеет ответственного, артефакт на выходе и срок. Между этапами — формальная передача артефакта (стрелки на диаграмме). Ниже — визуальная карта этапов:"},
]));

// Диаграмма 8 этапов
blocks.push(buildStageDiagram([
  {num: 1, name: 'Анализ рынка\nи конкурентов', owner: 'Ведущий спец.\nтоварной политики', days: '5-7 дней'},
  {num: 2, name: 'Ценообразование\nи прогноз', owner: 'Ведущий спец.\nтоварной политики', days: '3-5 дней'},
  {num: 3, name: 'Маркетинговые\nматериалы', owner: 'Ведущий спец.\nтоварной политики', days: '7-10 дней'},
  {num: 4, name: 'Утверждение\nплана запуска', owner: 'Директор\nпо маркетингу', days: '2-3 дня'},
]));

blocks.push(new Paragraph({ spacing: { before: 120 } }));
blocks.push(body([{t: "↓", b: true}]));

blocks.push(buildStageDiagram([
  {num: 5, name: 'Обучение\nотдела продаж', owner: 'Ведущий спец. +\nДиректор по продажам', days: '3-5 дней'},
  {num: 6, name: 'Сбор\nпотребности', owner: 'Региональные\nменеджеры', days: '7-10 дней'},
  {num: 7, name: 'Запуск\nв точках продаж', owner: 'ЭТП\nв областях', days: '14-21 день'},
  {num: 8, name: 'Обратная связь\nи корректировка', owner: 'Региональные\nдиректора', days: '30 дней'},
]));

blocks.push(body([
  {t: "\nСуммарный срок вывода: ", b: true},
  {t: "от 39 до 81 рабочего дня (≈ 8-12 недель). Этапы 1-4 — подготовительные (маркетинг), этапы 5-8 — исполнительные (продажи)."}
]));

// === 5. ДЕТАЛЬНОЕ ОПИСАНИЕ ЭТАПОВ ===
blocks.push(h1("5. Детальное описание этапов"));

// --- Этап 1 ---
blocks.push(h2("5.1. Этап 1. Анализ рынка и конкурентов"));
blocks.push(body([{t: "Ответственный: ", b: true}, {t: "Ведущий специалист группы товарной политики"}, {t: " | Срок: ", b: true}, {t: "5-7 рабочих дней"}]));
blocks.push(body([{t: "Цель этапа: ", b: true}, {t: "получить полную картину конкурентного окружения новинки, определить её место на рынке и целевые сегменты покупателей."}]));
blocks.push(body([{t: "Действия:", b: true}]));
blocks.push(bullet("Составить список прямых и косвенных конкурентов (минимум 5-7 брендов в категории). Источники: Яндекс.Маркет, ВсеИнструменты.ру, Petrovich, Леруа Мерлен, сайты конкурентов."));
blocks.push(bullet("Собрать данные по каждому конкуренту: цена, фасовка, описание, ключевые преимущества, классы опасности, наличие сертификатов."));
blocks.push(bullet("Построить матрицу позиционирования: где находится новинка относительно конкурентов по цене и качеству."));
blocks.push(bullet("Определить целевые сегменты покупателей (B2C: частные мастера, DIY; B2B: прорабы, бригадиры, дистрибьюторы, дизайнеры)."));
blocks.push(bullet("Оценить объём рынка категории и потенциальную долю новинки (в штуках и рублях за 12 месяцев)."));
blocks.push(body([{t: "Артефакт на выходе: ", b: true}, {t: "Competitive-анализ (Excel/PowerPoint) с матрицей позиционирования и прогнозом объёма рынка."}]));

// --- Этап 2 ---
blocks.push(h2("5.2. Этап 2. Ценообразование и прогноз"));
blocks.push(body([{t: "Ответственный: ", b: true}, {t: "Ведущий специалист группы товарной политики"}, {t: " | Срок: ", b: true}, {t: "3-5 рабочих дней"}]));
blocks.push(body([{t: "Цель этапа: ", b: true}, {t: "рассчитать ценовую политику новинки и согласовать прогнозный объём продаж с отделом продаж."}]));
blocks.push(body([{t: "Действия:", b: true}]));
blocks.push(bullet("Рассчитать себестоимость единицы продукта (с учётом сырья, упаковки, логистики)."));
blocks.push(bullet("Определить рекомендуемую розничную цену (РРЦ) исходя из конкурентного анализа (этап 1) и маржинальности."));
blocks.push(bullet("Установить вилку цен: оптовая цена для дистрибьюторов, рекомендованная мелкооптовая, РРЦ для розницы."));
blocks.push(bullet("Подготовить прогноз продаж на 3/6/12 месяцев: по федеральным округам, по каналам продаж (DIY, B2B, маркетплейсы)."));
blocks.push(bullet("Согласовать прогноз с директором по продажам (через директора по маркетингу) — встреча с региональными директорами."));
blocks.push(body([{t: "Артефакт на выходе: ", b: true}, {t: "Ценовое обоснование + прогноз продаж (Excel) с разбивкой по ФО и каналам."}]));

// --- Этап 3 ---
blocks.push(h2("5.3. Этап 3. Подготовка маркетинговых материалов"));
blocks.push(body([{t: "Ответственный: ", b: true}, {t: "Ведущий специалист группы товарной политики"}, {t: " | Срок: ", b: true}, {t: "7-10 рабочих дней"}]));
blocks.push(body([{t: "Цель этапа: ", b: true}, {t: "создать полный комплект материалов для отдела продаж и точек продаж. Без этих материалов обучение ЭТП и запуск в точках невозможны."}]));
blocks.push(body([{t: "Действия:", b: true}]));
blocks.push(bullet("Сформировать УТП (уникальное торговое предложение) — 1-2 предложения, объясняющие, чем новинка лучше конкурентов."));
blocks.push(bullet("Подготовить презентацию новинки для отдела продаж (PowerPoint, 10-15 слайдов): продукт, преимущества, цена, конкуренты, целевая аудитория, точки упора при продаже."));
blocks.push(bullet("Разработать техническое описание продукта (PDF, 2-3 страницы): характеристики, расход, способы нанесения, сертификаты."));
blocks.push(bullet("Создать бриф на производство POSM (воблеры, шелфтокеры, ценникодержатели, листовки)."));
blocks.push(bullet("Подготовить фото и видео продукта для использования в презентациях, на сайте, в каталоге."));
blocks.push(bullet("Сформировать «точки упора при продаже» — 5-7 ключевых аргументов для ЭТП, которые они будут использовать в разговоре с покупателем."));
blocks.push(body([{t: "Артефакты на выходе: ", b: true}, {t: "комплект из 5 документов: презентация новинки + техническое описание + УТП + бриф на POSM + точки упора."}]));

// --- Этап 4 ---
blocks.push(h2("5.4. Этап 4. Утверждение плана запуска"));
blocks.push(body([{t: "Ответственный: ", b: true}, {t: "Директор по маркетингу"}, {t: " | Срок: ", b: true}, {t: "2-3 рабочих дня"}]));
blocks.push(body([{t: "Цель этапа: ", b: true}, {t: "получить формальное решение о запуске с утверждёнными бюджетом, календарём и ответственными."}]));
blocks.push(body([{t: "Действия:", b: true}]));
blocks.push(bullet("Собрать итоги этапов 1-3 в единый документ «План запуска новинки» (Word/PDF)."));
blocks.push(bullet("Провести встречу с директором по продажам для согласования прогноза и квот."));
blocks.push(bullet("Утвердить календарный план запуска (сроки по 8 этапам)."));
blocks.push(bullet("Утвердить бюджет маркетинга и коммерческих промо."));
blocks.push(bullet("Назначить ответственных за каждый этап формально (приказом)."));
blocks.push(body([{t: "Артефакт на выходе: ", b: true}, {t: "Утверждённый «План запуска новинки» (Word/PDF) с подписями директора по маркетингу и директора по продажам."}]));

// --- Этап 5 ---
blocks.push(h2("5.5. Этап 5. Обучение отдела продаж"));
blocks.push(body([{t: "Ответственный: ", b: true}, {t: "Ведущий специалист товарной политики + Директор по продажам"}, {t: " | Срок: ", b: true}, {t: "3-5 рабочих дней"}]));
blocks.push(body([{t: "Цель этапа: ", b: true}, {t: "обучить всю цепочку продаж — от региональных директоров до ЭТП — работать с новинкой. Без обучения продажи запрещены."}]));
blocks.push(body([{t: "Действия:", b: true}]));
blocks.push(bullet("Провести онлайн-встречу (Zoom/Teams) с региональными директорами и региональными менеджерами: презентация новинки, обзор рынка, цена, прогноз, квоты по ФО."));
blocks.push(bullet("Передать региональным менеджерам комплект материалов (презентация, тех. описание, точки упора) для обучения ЭТП в областях."));
blocks.push(bullet("Региональные менеджеры проводят обучение ЭТП в каждой области (очно или онлайн): как продавать, какие аргументы использовать, какие точки упора."));
blocks.push(bullet("ЭТП проходят тест на знание новинки (минимум 80% правильных ответов)."));
blocks.push(bullet("Подготовить образцы новинки для отправки ЭТП — по 2-3 штуки на ЭТП для демонстрации в точках."));
blocks.push(body([{t: "Артефакт на выходе: ", b: true}, {t: "Протокол обучения + результаты теста ЭТП + подтверждение отправки образцов."}]));

// --- Этап 6 ---
blocks.push(h2("5.6. Этап 6. Сбор потребности от регионов"));
blocks.push(body([{t: "Ответственный: ", b: true}, {t: "Региональные менеджеры (под руководством региональных директоров)"}, {t: " | Срок: ", b: true}, {t: "7-10 рабочих дней"}]));
blocks.push(body([{t: "Цель этапа: ", b: true}, {t: "собрать от ЭТП первичную потребность в новинке: сколько штук и в какие точки они планируют отгрузить в первый месяц."}]));
blocks.push(body([{t: "Действия:", b: true}]));
blocks.push(bullet("ЭТП посещают ключевые точки продаж в своих областях и проводят презентацию новинки (по материалам этапа 3)."));
blocks.push(bullet("ЭТП собирают предварительные заказы от каждой точки (в штуках, с указанием даты поставки)."));
blocks.push(bullet("Региональные менеджеры сводят заказы по областям и передают региональным директорам."));
blocks.push(bullet("Региональные директора сводят заказы по ФО и передают директору по продажам."));
blocks.push(bullet("Директор по продажам сверяет собранную потребность с прогнозом (этап 2). При значительном расхождении — корректировка квот."));
blocks.push(body([{t: "Артефакт на выходе: ", b: true}, {t: "Сводный заказ на новинку (Excel) с разбивкой по ФО, областям, точкам."}]));

// --- Этап 7 ---
blocks.push(h2("5.7. Этап 7. Запуск в точках продаж"));
blocks.push(body([{t: "Ответственный: ", b: true}, {t: "ЭТП (под контролем региональных менеджеров)"}, {t: " | Срок: ", b: true}, {t: "14-21 рабочий день"}]));
blocks.push(body([{t: "Цель этапа: ", b: true}, {t: "физически разместить новинку в пилотных точках продаж, обеспечить выкладку, POSM, ценники, и начать продажи."}]));
blocks.push(body([{t: "Рекомендации по выбору точек продаж:", b: true}]));
blocks.push(bullet([{t: "Принцип ", b: true}, {t: "«10-20 пилотных точек на регион»"}]));
blocks.push(bullet("Точки должны быть разных типов: 50% — DIY-сети (Леруа Мерлен, Петрович, ВсеИнструменты), 30% — региональные сети, 20% — одиночные магазины."));
blocks.push(bullet("Точка должна иметь высокий трафик целевой аудитории (мастера, DIY-покупатели)."));
blocks.push(bullet("В точке должен быть компетентный продавец, прошедший обучение (через ЭТП)."));
blocks.push(bullet("Точка должна agreeing на выкладку новинки с POSM."));
blocks.push(body([{t: "Действия по запуску:", b: true}]));
blocks.push(bullet("ЭТП доставляет товар и POSM в пилотные точки."));
blocks.push(bullet("ЭТП контролирует выкладку: новинка размещается на уровне глаз, POSM (воблеры, шелфтокеры) — рядом, ценник с ценой."));
blocks.push(bullet("ЭТП фотографирует выкладку (фото до и фото после) и передаёт региональному менеджеру."));
blocks.push(bullet("Региональный менеджер проверяет 20% точек лично в течение первой недели."));
blocks.push(bullet("ЭТП собирает первые заказы и передаёт их региональному менеджеру для отгрузки."));
blocks.push(body([{t: "Артефакт на выходе: ", b: true}, {t: "Фотоотчёт выкладки + первые заказы + журнал посещений."}]));

// --- Этап 8 ---
blocks.push(h2("5.8. Этап 8. Обратная связь и корректировка"));
blocks.push(body([{t: "Ответственный: ", b: true}, {t: "Региональные директора (сводят обратную связь по ФО)"}, {t: " | Срок: ", b: true}, {t: "30 дней после запуска"}]));
blocks.push(body([{t: "Цель этапа: ", b: true}, {t: "оценить результаты первого месяца продаж, собрать обратную связь от рынка и внести корректировки в стратегию."}]));
blocks.push(body([{t: "Действия:", b: true}]));
blocks.push(bullet("ЭТП собирают обратную связь от продавцов и покупателей: что нравится, что не нравится, какие вопросы возникают."));
blocks.push(bullet("Региональные менеджеры сводят обратную связь по областям (раз в неделю)."));
blocks.push(bullet("Региональные директора сводят обратную связь по ФО (раз в 2 недели) и передают директору по продажам."));
blocks.push(bullet("Директор по продажам и директор по маркетингу проводят совместную встречу (через 30 дней после запуска): результаты vs прогноз, корректировки."));
blocks.push(bullet("При необходимости — корректировки цены, POSM, обучения ЭТП. Решение о масштабировании на все точки канала (или возврат на доработку)."));
blocks.push(body([{t: "Артефакт на выходе: ", b: true}, {t: "Отчёт о первых 30 днях продаж (Word/Excel) с выводами и решениями по корректировкам."}]));

// === 6. МАТРИЦА РОЛЕЙ И ЭТАПОВ (RACI) ===
blocks.push(h1("6. Матрица ролей и этапов (RACI)"));
blocks.push(body([
  {t: "Матрица RACI показывает, кто ", b: false},
  {t: "отвечает (R — Responsible)", b: true},
  {t: ", кто ", b: false},
  {t: "согласовывает (A — Accountable)", b: true},
  {t: ", кого ", b: false},
  {t: "консультируют (C — Consulted)", b: true},
  {t: " и кого ", b: false},
  {t: "информируют (I — Informed)", b: true},
  {t: " на каждом этапе. Это исключает дублирование и пропуски ответственности."}
]));

blocks.push(makeTable(
  ["Этап", "Директор\nпо маркетингу", "Ведущий спец.\nтов. политики", "Директор\nпо продажам", "Региональный\nдиректор", "Региональный\nменеджер", "ЭТП"],
  [
    [{text: "1. Анализ рынка и конкурентов", bold: true},
     {text: "I", align: AlignmentType.CENTER},
     {text: "R", align: AlignmentType.CENTER, bold: true},
     {text: "—", align: AlignmentType.CENTER},
     {text: "—", align: AlignmentType.CENTER},
     {text: "—", align: AlignmentType.CENTER},
     {text: "—", align: AlignmentType.CENTER}],
    [{text: "2. Ценообразование и прогноз", bold: true},
     {text: "A", align: AlignmentType.CENTER, bold: true},
     {text: "R", align: AlignmentType.CENTER, bold: true},
     {text: "C", align: AlignmentType.CENTER},
     {text: "C", align: AlignmentType.CENTER},
     {text: "—", align: AlignmentType.CENTER},
     {text: "—", align: AlignmentType.CENTER}],
    [{text: "3. Маркетинговые материалы", bold: true},
     {text: "A", align: AlignmentType.CENTER, bold: true},
     {text: "R", align: AlignmentType.CENTER, bold: true},
     {text: "I", align: AlignmentType.CENTER},
     {text: "—", align: AlignmentType.CENTER},
     {text: "—", align: AlignmentType.CENTER},
     {text: "—", align: AlignmentType.CENTER}],
    [{text: "4. Утверждение плана запуска", bold: true},
     {text: "R", align: AlignmentType.CENTER, bold: true},
     {text: "C", align: AlignmentType.CENTER},
     {text: "A", align: AlignmentType.CENTER, bold: true},
     {text: "I", align: AlignmentType.CENTER},
     {text: "—", align: AlignmentType.CENTER},
     {text: "—", align: AlignmentType.CENTER}],
    [{text: "5. Обучение отдела продаж", bold: true},
     {text: "A", align: AlignmentType.CENTER, bold: true},
     {text: "R", align: AlignmentType.CENTER, bold: true},
     {text: "R", align: AlignmentType.CENTER, bold: true},
     {text: "I", align: AlignmentType.CENTER},
     {text: "I", align: AlignmentType.CENTER},
     {text: "—", align: AlignmentType.CENTER}],
    [{text: "6. Сбор потребности", bold: true},
     {text: "I", align: AlignmentType.CENTER},
     {text: "—", align: AlignmentType.CENTER},
     {text: "A", align: AlignmentType.CENTER, bold: true},
     {text: "R", align: AlignmentType.CENTER, bold: true},
     {text: "R", align: AlignmentType.CENTER, bold: true},
     {text: "R", align: AlignmentType.CENTER, bold: true}],
    [{text: "7. Запуск в точках продаж", bold: true},
     {text: "I", align: AlignmentType.CENTER},
     {text: "C", align: AlignmentType.CENTER},
     {text: "A", align: AlignmentType.CENTER, bold: true},
     {text: "C", align: AlignmentType.CENTER},
     {text: "R", align: AlignmentType.CENTER, bold: true},
     {text: "R", align: AlignmentType.CENTER, bold: true}],
    [{text: "8. Обратная связь и корректировка", bold: true},
     {text: "C", align: AlignmentType.CENTER},
     {text: "C", align: AlignmentType.CENTER},
     {text: "A", align: AlignmentType.CENTER, bold: true},
     {text: "R", align: AlignmentType.CENTER, bold: true},
     {text: "C", align: AlignmentType.CENTER},
     {text: "R", align: AlignmentType.CENTER, bold: true}],
  ],
  [25, 13, 13, 13, 12, 12, 12]
));

blocks.push(body([{t: "Легенда: ", b: true}, {t: "R — Responsible (отвечает за выполнение) · A — Accountable (согласовывает, принимает решение) · C — Consulted (консультируют) · I — Informed (информируют)"}]));

// === 7. АРТЕФАКТЫ ЭТАПОВ ===
blocks.push(h1("7. Артефакты этапов — чек-лист"));
blocks.push(body("Полный перечень документов и материалов, которые должны быть готовы на каждом этапе. Чек-лист обязателен к заполнению перед переходом на следующий этап."));

blocks.push(makeTable(
  ["Этап", "Артефакт", "Формат", "Согласует"],
  [
    [{text: "1", align: AlignmentType.CENTER, bold: true},
     {text: "Competitive-анализ с матрицей позиционирования", bold: true},
     "Excel + PowerPoint", "Директор по маркетингу"],
    [{text: "2", align: AlignmentType.CENTER, bold: true},
     {text: "Ценовое обоснование + прогноз продаж", bold: true},
     "Excel", "Директор по маркетингу + Директор по продажам"],
    [{text: "3", align: AlignmentType.CENTER, bold: true},
     {text: "Презентация новинки для отдела продаж", bold: true},
     "PowerPoint (10-15 слайдов)", "Директор по маркетингу"],
    [{text: "3", align: AlignmentType.CENTER, bold: true},
     {text: "Техническое описание продукта", bold: true},
     "PDF (2-3 стр.)", "Директор по маркетингу"],
    [{text: "3", align: AlignmentType.CENTER, bold: true},
     {text: "УТП + точки упора при продаже", bold: true},
     "Word (1 стр.)", "Директор по маркетингу + Директор по продажам"],
    [{text: "3", align: AlignmentType.CENTER, bold: true},
     {text: "Бриф на производство POSM", bold: true},
     "Word", "Директор по маркетингу"],
    [{text: "4", align: AlignmentType.CENTER, bold: true},
     {text: "Утверждённый План запуска новинки", bold: true},
     "Word/PDF", "Директор по маркетингу + Директор по продажам"],
    [{text: "5", align: AlignmentType.CENTER, bold: true},
     {text: "Протокол обучения + результаты теста ЭТП", bold: true},
     "Excel", "Директор по продажам"],
    [{text: "6", align: AlignmentType.CENTER, bold: true},
     {text: "Сводный заказ на новинку", bold: true},
     "Excel (по ФО и точкам)", "Директор по продажам"],
    [{text: "7", align: AlignmentType.CENTER, bold: true},
     {text: "Фотоотчёт выкладки + первые заказы", bold: true},
     "Фото + Excel", "Региональный менеджер"],
    [{text: "8", align: AlignmentType.CENTER, bold: true},
     {text: "Отчёт о первых 30 днях продаж", bold: true},
     "Word/Excel", "Директор по продажам + Директор по маркетингу"],
  ],
  [6, 38, 25, 31]
));

// === 8. СВЯЗЬ С ГЕНЕРАЦИЕЙ СПРОСА ===
blocks.push(h1("8. Связь с генерацией спроса у отдела продаж"));
blocks.push(body([
  {t: "Ключевая особенность стандарта: ", b: true},
  {t: "Ведущий специалист товарной политики ", b: false},
  {t: "не только готовит материалы, но и рассчитывает, ", b: false},
  {t: "как через отдел продаж генерировать спрос", b: true},
  {t: " на новинку. Это значит, что маркетинг отвечает за «что сказать» и «кому сказать», а продажи — за «как сказать» и «кому дойти»."}
]));
blocks.push(body([{t: "Механизм генерации спроса:", b: true}]));
blocks.push(bullet([{t: "Этап 3: ", b: true}, {t: "Ведущий специалист формирует «точки упора» — 5-7 аргументов для ЭТП, которые они будут использовать в разговоре с покупателем."}]));
blocks.push(bullet([{t: "Этап 5: ", b: true}, {t: "ЭТП обучают этим точкам упора и проверяют через тест (минимум 80% правильных ответов)."}]));
blocks.push(bullet([{t: "Этап 6: ", b: true}, {t: "ЭТП идут в точки продаж с готовыми аргументами и собирают первичную потребность — это и есть генерация спроса на уровне канала."}]));
blocks.push(bullet([{t: "Этап 7: ", b: true}, {t: "ЭТП продолжают генерацию в точках продаж, опираясь на точки упора, POSM и образцы."}]));
blocks.push(bullet([{t: "Этап 8: ", b: true}, {t: "Если спрос ниже прогноза — ведущий специалист корректирует точки упора, цену или POSM. Цикл повторяется."}]));

blocks.push(body([
  {t: "Маркетинговая поддержка спроса (опционально): ", b: true},
  {t: "параллельно с этапом 7 можно запустить таргетированную рекламу, контент-маркетинг (статьи в Дзене, видео на Rutube) и интеграции у блогеров — это усиливает действия ЭТП на местах."}
]));

// === 9. СРОКИ И КОНТРОЛЬНЫЕ ТОЧКИ ===
blocks.push(h1("9. Сроки и контрольные точки"));
blocks.push(body("Минимальные и максимальные сроки по этапам и контрольные точки для руководства:"));

blocks.push(makeTable(
  ["Этап", "Мин. срок", "Макс. срок", "Контрольная точка", "Кто контролирует"],
  [
    [{text: "1. Анализ рынка", bold: true}, "5 дней", "7 дней", "Готов Competitive-анализ", "Директор по маркетингу"],
    [{text: "2. Ценообразование", bold: true}, "3 дня", "5 дней", "Цена + прогноз согласованы", "Директор по маркетингу"],
    [{text: "3. Маркетинговые материалы", bold: true}, "7 дней", "10 дней", "Комплект материалов готов", "Директор по маркетингу"],
    [{text: "4. Утверждение плана", bold: true}, "2 дня", "3 дня", "План запуска подписан", "Директор по маркетингу"],
    [{text: "5. Обучение", bold: true}, "3 дня", "5 дней", "Протокол обучения + тесты", "Директор по продажам"],
    [{text: "6. Сбор потребности", bold: true}, "7 дней", "10 дней", "Сводный заказ готов", "Директор по продажам"],
    [{text: "7. Запуск в точках", bold: true}, "14 дней", "21 день", "Фотоотчёт + первые заказы", "Региональный директор"],
    [{text: "8. Обратная связь", bold: true}, "30 дней", "30 дней", "Отчёт о 30 днях продаж", "Директор по продажам"],
    [{text: "ИТОГО", bold: true}, {text: "71 день", bold: true}, {text: "91 день", bold: true}, {text: "Полный цикл вывода", bold: true}, "—"],
  ],
  [22, 12, 12, 32, 22]
));

// === 10. КОНТРОЛЬ КАЧЕСТВА ===
blocks.push(h1("10. Контроль качества запуска"));
blocks.push(body("Стандарт включает 5 контрольных точек, где процесс может быть остановлен при невыполнении критериев:"));

const control_points = [
  ["CP1 (после этапа 1)", "Competitive-анализ содержит минимум 5 конкурентов, матрицу позиционирования и прогноз объёма рынка. Если нет —回到 этап 1."],
  ["CP2 (после этапа 2)", "Цена новинки не выше топ-3 конкурентов в категории. Прогноз продаж согласован с директором по продажам. Если нет —回到 этап 2."],
  ["CP3 (после этапа 3)", "Комплект из 5 материалов готов и утверждён директором по маркетингу. Если нет —回到 этап 3."],
  ["CP4 (после этапа 5)", "Все ЭТП прошли тест на знание новинки (≥80% правильных ответов). Если нет — обучение повторяется."],
  ["CP5 (после этапа 7)", "Фотоотчёт выкладки минимум из 80% пилотных точек. Первые заказы получены минимум из 60% пилотных точек. Если нет —回到 этап 7."],
];
for (const [cp, desc] of control_points) {
  blocks.push(bullet([{t: cp + ": ", b: true}, {t: desc}]));
}

// === ЗАКЛЮЧЕНИЕ ===
blocks.push(h1("11. Заключение"));
blocks.push(body([
  {t: "Стандарт обеспечивает единый порядок вывода новинок на рынок для всех брендов «Эмпилс». ", b: true},
  {t: "Соблюдение стандарта снижает риск ошибки в ценообразовании, упущенной выкладки в точках, слабого обучения ЭТП и — главное — невыполнения плана продаж новинки."}
]));
blocks.push(body([
  {t: "Ключевой принцип: ", b: true},
  {t: "маркетинг отвечает за «что» (анализ, цена, материалы, точки упора), продажи — за «как» (обучение, выкладка, заказы). Связку обеспечивает ведущий специалист товарной политики — он же «владелец» запуска со стороны маркетинга."}
]));
blocks.push(body([
  {t: "Стандарт пересматривается ", b: true},
  {t: "раз в год по итогам запусков. Внесение изменений утверждается директором по маркетингу и директором по продажам."}
]));

// ============================================================
// СБОРКА ДОКУМЕНТА
// ============================================================
const coverConfig = {
  title: "Стандарт вывода новинки на рынок",
  subtitle: "Поэтапная система взаимодействия отдела маркетинга и отдела продаж",
  englishLabel: "EMPILS PRODUCT LAUNCH STANDARD",
  metaLines: [
    "Заказчик: ЗАО «Эмпилс» (TITAN, NEWTONE, ОРЕОЛ, AVANTGARDE)",
    "Точка отсчёта: формула разработана, продукция готова к производству",
    "Тип документа: Внутренний стандарт (обязателен к исполнению)",
    "Версия: 1.0",
    "Дата: сентябрь 2026",
  ],
  footerLeft: "Эмпилс · Стандарт вывода новинки · v1.0",
  footerRight: "Конфиденциально · Внутренний документ",
};

const doc = new Document({
  creator: "ЗАО Эмпилс",
  title: "Стандарт вывода новинки на рынок",
  description: "Поэтапная система взаимодействия отдела маркетинга и отдела продаж",
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
          size: { width: 16838, height: 11906 },
          margin: { top: 1440, bottom: 1440, left: 1417, right: 1417 },
          pageNumbers: { start: 1, formatType: NumberFormat.DECIMAL },
        },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({
              text: "Эмпилс · Стандарт вывода новинки на рынок",
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
  const outPath = "/home/z/my-project/download/Эмпилс_Стандарт_вывода_новинки.docx";
  fs.writeFileSync(outPath, buf);
  console.log("OK saved:", outPath);
  console.log("Size:", buf.length, "bytes");
});
