// Персонализированное приглашение для Алексея Анатольевича (Олимп)
// посетить завод Эмпилс. Лаконичный формат делового письма.

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, PageNumber, AlignmentType, HeadingLevel,
  WidthType, BorderStyle, ShadingType, PageBreak,
} = require("docx");
const fs = require("fs");

// === Палитра ===
const P = {
  primary: "#1F4E79",
  accent: "#C00000",
  body: "#1A1A1A",
  secondary: "#595959",
  muted: "#808080",
};
const c = (hex) => hex.replace("#", "");

const FONT = { ascii: "Calibri", eastAsia: "Calibri" };
const FONT_HEAD = { ascii: "Calibri", eastAsia: "Calibri" };

// === Утилиты ===
function p(text, opts = {}) {
  const runs = Array.isArray(text) ? text : [new TextRun({
    text, size: 22, color: c(P.body), font: FONT,
  })];
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 312, after: 160 },
    ...opts,
    children: runs,
  });
}

// Параграф с inline-форматированием
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
    spacing: { line: 312, after: 160 },
    ...opts,
    children: runs,
  });
}

// Пункт списка с маркером
function li(text, opts = {}) {
  const runs = Array.isArray(text) ? text : [new TextRun({
    text, size: 22, color: c(P.body), font: FONT,
  })];
  return new Paragraph({
    spacing: { line: 312, after: 100 },
    bullet: { level: 0 },
    ...opts,
    children: runs,
  });
}

// Пустая строка
function empty() {
  return new Paragraph({
    spacing: { line: 240, after: 0 },
    children: [new TextRun({ text: "" })],
  });
}

// Горизонтальная линия
function hr() {
  return new Paragraph({
    spacing: { before: 120, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: c(P.primary) } },
    children: [new TextRun({ text: "", size: 2 })],
  });
}

// === СОДЕРЖАНИЕ ПИСЬМА ===
const children = [];

// Шапка письма — слева отправитель, справа дата
children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: {
    top: { style: BorderStyle.NONE },
    bottom: { style: BorderStyle.NONE },
    left: { style: BorderStyle.NONE },
    right: { style: BorderStyle.NONE },
    insideHorizontal: { style: BorderStyle.NONE },
    insideVertical: { style: BorderStyle.NONE },
  },
  rows: [
    new TableRow({
      cantSplit: true,
      children: [
        new TableCell({
          width: { size: 60, type: WidthType.PERCENTAGE },
          margins: { top: 0, bottom: 0, left: 0, right: 0 },
          children: [new Paragraph({
            spacing: { line: 280 },
            children: [
              new TextRun({ text: "ЗАО «ЭМПИЛС»", bold: true, size: 24, color: c(P.primary), font: FONT_HEAD }),
            ],
          }),
          new Paragraph({
            spacing: { line: 260, after: 0 },
            children: [
              new TextRun({ text: "Ростов-на-Дону · empils.ru", size: 18, color: c(P.secondary), font: FONT }),
            ],
          })],
        }),
        new TableCell({
          width: { size: 40, type: WidthType.PERCENTAGE },
          margins: { top: 0, bottom: 0, left: 0, right: 0 },
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            spacing: { line: 280 },
            children: [
              new TextRun({ text: "Октябрь 2026", size: 22, color: c(P.body), font: FONT }),
            ],
          })],
        }),
      ],
    }),
  ],
}));

children.push(hr());

// Обращение
children.push(new Paragraph({
  spacing: { line: 312, before: 240, after: 240 },
  children: [
    new TextRun({ text: "Алексею Анатольевичу", size: 22, color: c(P.body), font: FONT, bold: true }),
  ],
}));

children.push(new Paragraph({
  spacing: { line: 280, after: 60 },
  children: [
    new TextRun({ text: "компания «Олимп»", size: 20, color: c(P.secondary), font: FONT, italics: true }),
  ],
}));

// Заголовок письма
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 240, after: 240 },
  children: [
    new TextRun({
      text: "Приглашение посетить завод «Эмпилс»",
      bold: true, size: 28, color: c(P.primary), font: FONT_HEAD,
    }),
  ],
}));

// Основной текст — обращение и контекст
children.push(pRich([
  ["Уважаемый Алексей Анатольевич!", { bold: true }],
]));

children.push(p(
  "Благодарю Вас за встречи и конструктивное обсуждение коммерческих условий сотрудничества. " +
  "Чтобы Вы лично оценили мощь производства и убедились в надёжности «Эмпилс» как долгосрочного партнёра, " +
  "приглашаю Вас посетить наш завод в Ростове-на-Дону в удобное для Вас время."
));

// Ключевое условие — выделенный блок
children.push(new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: {
    top: { style: BorderStyle.SINGLE, size: 8, color: c(P.primary) },
    bottom: { style: BorderStyle.SINGLE, size: 8, color: c(P.primary) },
    left: { style: BorderStyle.SINGLE, size: 2, color: c(P.primary) },
    right: { style: BorderStyle.SINGLE, size: 2, color: c(P.primary) },
    insideHorizontal: { style: BorderStyle.NONE },
    insideVertical: { style: BorderStyle.NONE },
  },
  rows: [
    new TableRow({
      cantSplit: true,
      children: [new TableCell({
        shading: { type: ShadingType.CLEAR, fill: "F2F7FC" },
        margins: { top: 200, bottom: 200, left: 280, right: 280 },
        children: [
          new Paragraph({
            spacing: { line: 312, after: 80 },
            children: [
              new TextRun({ text: "Эмпилс берёт на себя все расходы:", bold: true, size: 24, color: c(P.primary), font: FONT }),
            ],
          }),
          new Paragraph({
            spacing: { line: 312, after: 60 },
            children: [
              new TextRun({ text: "✓  ", size: 22, color: c(P.primary), bold: true, font: FONT }),
              new TextRun({ text: "авиабилеты туда-обратно", size: 22, color: c(P.body), font: FONT }),
            ],
          }),
          new Paragraph({
            spacing: { line: 312, after: 60 },
            children: [
              new TextRun({ text: "✓  ", size: 22, color: c(P.primary), bold: true, font: FONT }),
              new TextRun({ text: "проживание в гостинице на время визита", size: 22, color: c(P.body), font: FONT }),
            ],
          }),
          new Paragraph({
            spacing: { line: 312, after: 0 },
            children: [
              new TextRun({ text: "✓  ", size: 22, color: c(P.primary), bold: true, font: FONT }),
              new TextRun({ text: "трансфер из аэропорта и обратно", size: 22, color: c(P.body), font: FONT }),
            ],
          }),
        ],
      })],
    }),
  ],
}));

children.push(empty());

// Программа визита — лаконично
children.push(pRich([
  ["Что Вы увидите за один день:", { bold: true, color: P.primary, size: 24 }],
]));

children.push(li([
  new TextRun({ text: "Экскурсия по производству", size: 22, bold: true, color: c(P.body), font: FONT }),
  new TextRun({ text: " — лаборатории, линии розлива, контроль качества от сырья до готовой банки.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(li([
  new TextRun({ text: "Презентация ассортимента под Ваш регион", size: 22, bold: true, color: c(P.body), font: FONT }),
  new TextRun({ text: " — фокусные SKU по TITAN, PROFIWOOD, ОРЕОЛ, AvantGarde, NEWTONE с расчётом маржинальности под формат Вашей сети.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(li([
  new TextRun({ text: "Практический тест ЛКМ", size: 22, bold: true, color: c(P.body), font: FONT }),
  new TextRun({ text: " — лично сравните укрывистость, стойкость к мытью и истиранию с конкурентами.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(li([
  new TextRun({ text: "Обсуждение условий и фиксация договорённостей", size: 22, bold: true, color: c(P.body), font: FONT }),
  new TextRun({ text: " — закрепим то, что уже обсудили на прошлых встречах, и наметим стартовые отгрузки.", size: 22, color: c(P.body), font: FONT }),
]));

children.push(empty());

// Призыв к действию
children.push(pRich([
  ["Подтвердите, пожалуйста, удобные Вам даты визита ", {}],
  ["(в диапазоне ноябрь 2026 — февраль 2027)", { italic: true, color: P.secondary }],
  [" — и мы сразу забронируем билеты и проживание.", {}],
]));

children.push(p(
  "Буду рад лично показать Вам завод и ответить на все вопросы. " +
  "Уверен, что после визита Вы примете взвешенное решение о партнёрстве — " +
  "а мы со своей стороны гарантируем поддержку на каждом этапе."
));

children.push(empty());

// Подпись
children.push(new Paragraph({
  spacing: { line: 312, after: 60 },
  children: [
    new TextRun({ text: "С уважением,", size: 22, color: c(P.body), font: FONT }),
  ],
}));

children.push(new Paragraph({
  spacing: { line: 312, after: 0 },
  children: [
    new TextRun({ text: "________________________", size: 22, color: c(P.muted), font: FONT }),
  ],
}));

children.push(new Paragraph({
  spacing: { line: 312, after: 40 },
  children: [
    new TextRun({ text: "Региональный директор ЗАО «Эмпилс»", size: 22, bold: true, color: c(P.primary), font: FONT }),
  ],
}));

children.push(new Paragraph({
  spacing: { line: 280, after: 240 },
  children: [
    new TextRun({ text: "Ответственный за работу с компанией «Олимп»", size: 18, color: c(P.secondary), italics: true, font: FONT }),
  ],
}));

// Контактный блок — мелким шрифтом в подвале
children.push(hr());

children.push(new Paragraph({
  spacing: { line: 260, after: 40 },
  children: [
    new TextRun({ text: "Для подтверждения даты и согласования деталей:", size: 18, bold: true, color: c(P.secondary), font: FONT }),
  ],
}));

children.push(new Paragraph({
  spacing: { line: 260, after: 40 },
  children: [
    new TextRun({ text: "Телефон: ", size: 18, color: c(P.secondary), font: FONT }),
    new TextRun({ text: "+7 (___) ___-__-__", size: 18, color: c(P.body), font: FONT, bold: true }),
    new TextRun({ text: "   ·   E-mail: ", size: 18, color: c(P.secondary), font: FONT }),
    new TextRun({ text: "_________@empils.ru", size: 18, color: c(P.body), font: FONT, bold: true }),
  ],
}));

children.push(new Paragraph({
  spacing: { line: 260, after: 40 },
  children: [
    new TextRun({ text: "Адрес завода: ", size: 18, color: c(P.secondary), font: FONT }),
    new TextRun({ text: "344090, г. Ростов-на-Дону, 1-й Машиностроительный пер., 21", size: 18, color: c(P.body), font: FONT }),
  ],
}));

// === Документ ===
const doc = new Document({
  creator: "Z.ai",
  title: "Приглашение посетить завод Эмпилс — Алексей Анатольевич (Олимп)",
  description: "Персонализированное приглашение для дистрибьютора Олимп",
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
    children: children,
  }],
});

Packer.toBuffer(doc).then(buffer => {
  const outPath = "/home/z/my-project/download/Приглашение_Олимп_Алексей_Анатольевич.docx";
  fs.writeFileSync(outPath, buffer);
  console.log(`✓ Документ сохранён: ${outPath}`);
  console.log(`  Размер: ${buffer.length.toLocaleString()} байт`);
});
