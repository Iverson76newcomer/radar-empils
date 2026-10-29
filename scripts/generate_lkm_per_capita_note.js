/**
 * Краткая Word-заметка: "Потребление ЛКМ на душу населения — мировой контекст"
 * Без обложки, без TOC. 1-2 страницы A4, лаконичный стиль.
 */

const {
  Document, Packer, Paragraph, TextRun,
  AlignmentType, HeadingLevel, BorderStyle,
  Table, TableRow, TableCell, WidthType, ShadingType,
  TableLayoutType,
} = require("docx");
const fs = require("fs");

// === Палитра Empils ===
const P = {
  primary:  "1A2330",
  body:     "1F2937",
  secondary:"6B7280",
  accent:   "D4875A",  // графитово-оранжевый
  accentLight: "F8F0EB",
  surface:  "F5F5F5",
  positive: "2E7D32",
  danger:   "C62828",
  white:    "FFFFFF",
};
const c = (hex) => hex.replace("#", "");

const FONT = { ascii: "Calibri", eastAsia: "Microsoft YaHei" };
const NB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };

// === Хелперы ===

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 240, after: 120, line: 280 },
    children: [new TextRun({
      text, bold: true, size: 28, color: c(P.primary),
      font: FONT,
    })],
  });
}

function eyebrow(text) {
  return new Paragraph({
    spacing: { before: 0, after: 80 },
    children: [new TextRun({
      text, size: 16, color: c(P.accent), bold: true,
      font: FONT, characterSpacing: 30,
    })],
  });
}

function body(text) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 280, after: 120 },
    children: [new TextRun({
      text, size: 22, color: c(P.body),
      font: FONT,
    })],
  });
}

function bodyBold(parts) {
  // parts: [{text, bold?, color?}, ...]
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { line: 280, after: 120 },
    children: parts.map(p => new TextRun({
      text: p.text, bold: !!p.bold,
      color: c(p.color || P.body),
      size: 22, font: FONT,
    })),
  });
}

function bullet(text) {
  return new Paragraph({
    indent: { left: 360, hanging: 200 },
    spacing: { line: 260, after: 80 },
    children: [
      new TextRun({ text: "▸ ", size: 22, color: c(P.accent), bold: true, font: FONT }),
      new TextRun({ text, size: 22, color: c(P.body), font: FONT }),
    ],
  });
}

function bulletBold(lead, rest, color = P.body) {
  return new Paragraph({
    indent: { left: 360, hanging: 200 },
    spacing: { line: 260, after: 80 },
    children: [
      new TextRun({ text: "▸ ", size: 22, color: c(P.accent), bold: true, font: FONT }),
      new TextRun({ text: lead, size: 22, color: c(P.primary), bold: true, font: FONT }),
      new TextRun({ text: " — " + rest, size: 22, color: c(color), font: FONT }),
    ],
  });
}

function divider() {
  return new Paragraph({
    spacing: { before: 120, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: c(P.accent), space: 1 } },
    children: [],
  });
}

// === Таблица мирового потребления ЛКМ ===

function buildWorldTable() {
  const headerCells = ["Регион / Страна", "кг/чел", "Комментарий"];
  const rows = [
    ["🇺🇸 США (всего)",          "≈ 22",  "Абсолютный лидер — культура DIY, крупные дома"],
    ["🇩🇪 Германия",             "≈ 21",  "Европейский лидер — сильная промышленность"],
    ["🇯🇵 Япония",              "11,5–14,7", "Самый высокий в Азии"],
    ["🌍 МИРОВОЕ СРЕДНЕЕ",       "13–15", "Базовый бенчмарк"],
    ["🇷🇺 Россия (2014)",       "10,3",  "Старые данные — до стагнации"],
    ["🇺🇸 США (архитектурные)", "10,4",  "Только декоративные ЛКМ"],
    ["🇷🇺 Россия (2024)",       "8,1",   "Падение -21% за 10 лет"],
    ["🇨🇦 Канада",              "7,8",   "Близко к ЕС"],
    ["🇪🇺 ЕС (среднее)",        "7,2",   "Зрелый европейский рынок"],
    ["🇨🇳 Китай",               "6,4–15", "Разные оценки"],
    ["🇰🇿 Казахстан",           "7,0",   "Текущий рынок Эмпилс"],
    ["🇦🇿 Азербайджан",         "4,6",   "Самый недонасыщенный"],
    ["🇹🇲 Туркменистан",        "4,7",   "Малый рынок"],
    ["🇮🇳 Индия",               "4–5",   "Низший среди крупных рынков"],
  ];

  // Заголовок
  const headerRow = new TableRow({
    tableHeader: true, cantSplit: true,
    children: headerCells.map((text, idx) => new TableCell({
      shading: { type: ShadingType.CLEAR, fill: c(P.primary) },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 4, color: c(P.primary) },
        bottom: { style: BorderStyle.SINGLE, size: 4, color: c(P.primary) },
        left: NB, right: NB,
      },
      children: [new Paragraph({
        alignment: idx === 1 ? AlignmentType.CENTER : AlignmentType.LEFT,
        spacing: { line: 240 },
        children: [new TextRun({
          text, bold: true, size: 18, color: c(P.white),
          font: FONT,
        })],
      })],
    })),
  });

  // Данные с цветовой подсветкой для России
  const dataRows = rows.map((row, idx) => {
    const isRussia = row[0].includes("Россия");
    const isWorld = row[0].includes("МИРОВОЕ");
    const isHighlight = isRussia || isWorld;

    const fillColor = isRussia && row[0].includes("2024")
      ? c(P.accentLight)  // актуальная Россия — оранжевый
      : isWorld
        ? c(P.surface)  // мировое среднее — серое
        : idx % 2 === 0
          ? c("FAFAFA")
          : c(P.white);

    return new TableRow({
      cantSplit: true,
      children: row.map((text, ci) => new TableCell({
        shading: { type: ShadingType.CLEAR, fill: fillColor },
        margins: { top: 60, bottom: 60, left: 120, right: 120 },
        borders: {
          top: NB,
          bottom: { style: BorderStyle.SINGLE, size: 2, color: "E0E0E0" },
          left: NB, right: NB,
        },
        children: [new Paragraph({
          alignment: ci === 1 ? AlignmentType.CENTER : AlignmentType.LEFT,
          spacing: { line: 240 },
          children: [new TextRun({
            text,
            size: 18,
            bold: ci === 0 || ci === 1 || isHighlight,
            color: c(ci === 1 ? P.accent : P.body),
            font: ci === 1 ? { ascii: "Consolas", eastAsia: "Microsoft YaHei" } : FONT,
          })],
        })],
      })),
    });
  });

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: c(P.primary) },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: c(P.primary) },
      left: NB, right: NB, insideHorizontal: NB, insideVertical: NB,
    },
    rows: [headerRow, ...dataRows],
  });
}

// === Сборка документа ===

const children = [];

// Шапка-заголовок (без обложки)
children.push(eyebrow("АНАЛИЗ · ЭМПИЛС · СЕНТЯБРЬ 2026"));
children.push(h1("Потребление ЛКМ на душу населения — мировой контекст"));
children.push(body(
  "Краткая справка для оценки потенциала экспортных рынков Эмпилс. Базовый показатель — кг ЛКМ на 1 жителя в год. Источники: European Coatings, Coatings World, Kusumgar, Mordor Intelligence (2021-2026)."
));
children.push(divider());

// Таблица
children.push(new Paragraph({
  spacing: { before: 80, after: 80 },
  children: [new TextRun({
    text: "Мировой рейтинг потребления ЛКМ (кг/чел, год)",
    bold: true, size: 20, color: c(P.primary), font: FONT,
  })],
}));
children.push(buildWorldTable());

children.push(new Paragraph({
  spacing: { before: 80 },
  children: [new TextRun({
    text: "Данные для РФ 2024 г. — расчёт из файлов Эмпилс (1 181 тыс. т / 146 млн чел).",
    italics: true, size: 16, color: c(P.secondary), font: FONT,
  })],
}));

children.push(divider());

// === Выводы ===
children.push(h1("Ключевые выводы"));

children.push(bulletBold(
  "Россия (8,1 кг) ниже среднемирового (13–15 кг)",
  "рынок недонасыщен на 40–45%. Падение -21% за 10 лет (с 10,3 кг в 2014).",
  P.danger
));
children.push(bulletBold(
  "США (22 кг) — абсолютный лидер",
  "за счёт культуры DIY (Home Depot, Lowe's), крупных домов, сильной промышленности."
));
children.push(bulletBold(
  "Германия (21 кг) — европейский лидер",
  "сильная промышленность (Mercedes, BMW, BASF), традиция качества и ремонта."
));
children.push(bulletBold(
  "Текущие экспортные рынки Эмпилс — самые недонасыщенные в мире",
  "Азербайджан 4,6 кг, Туркменистан 4,7 кг, Узбекистан 5,2 кг, Таджикистан 9,1 кг."
));

children.push(divider());

// Стратегический вывод
children.push(h1("Стратегические возможности для Эмпилс"));

children.push(bodyBold([
  { text: "Внутренний рынок РФ: ", bold: true, color: P.primary },
  { text: "потенциал роста с 8 до 10–12 кг/чел (+25–50%). Драйверы: импортозамещение, рост водных ЛКМ (+3,9% в 2024), сдвиг в премиум (AvantGarde)." },
]));

children.push(bodyBold([
  { text: "Экспортные рынки: ", bold: true, color: P.primary },
  { text: "Азербайджан, Туркменистан, Узбекистан, Таджикистан — рост с 4-5 до 7-8 кг/чел = " },
  { text: "двукратное увеличение рынков", bold: true, color: P.accent },
  { text: " без изменения доли Эмпилс. Это стратегическое окно возможностей." },
]));

children.push(bodyBold([
  { text: "Эталон для сравнения: ", bold: true, color: P.primary },
  { text: "при достижении уровня РФ (8 кг/чел) — рынки Ц.Азии удвоятся. При уровне ЕС (7,2 кг) — growth +50%." },
]));

children.push(divider());

// Финал
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 200, after: 0 },
  children: [new TextRun({
    text: "ЗАО «Эмпилс» · Аналитический отдел · 2026",
    size: 16, color: c(P.secondary), italics: true, font: FONT,
  })],
}));

// === Документ ===

const doc = new Document({
  creator: "ЗАО Эмпилс",
  title: "Потребление ЛКМ на душу населения — мировой контекст",
  styles: {
    default: {
      document: {
        run: { font: FONT, size: 22, color: c(P.body) },
        paragraph: { spacing: { line: 280 } },
      },
    },
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1200, bottom: 1200, left: 1200, right: 1000 },
      },
    },
    children,
  }],
});

const OUT = "/home/z/my-project/download/Эмпилс_Потребление_ЛКМ_мировой_контекст.docx";

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log("✅ Документ сохранён:", OUT);
  console.log("   Размер:", buf.length.toLocaleString(), "байт");
}).catch(err => {
  console.error("❌", err);
  process.exit(1);
});
