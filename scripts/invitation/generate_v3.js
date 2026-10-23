// Приглашение v3: добавлен гала-ужин с неформальным общением

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, ShadingType, WidthType,
} = require("docx");
const fs = require("fs");

const P = {
  primary: "#1F4E79",
  body: "#1A1A1A",
  secondary: "#595959",
};
const c = (hex) => hex.replace("#", "");

const FONT = { ascii: "Calibri", eastAsia: "Calibri" };

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

function empty() {
  return new Paragraph({
    spacing: { line: 240, after: 0 },
    children: [new TextRun({ text: "" })],
  });
}

const children = [];

// Заголовок письма
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 0, after: 240 },
  children: [
    new TextRun({
      text: "Приглашение посетить завод «Эмпилс»",
      bold: true, size: 28, color: c(P.primary), font: FONT,
    }),
  ],
}));

// Обращение
children.push(pRich([
  ["Уважаемый Алексей Анатольевич!", { bold: true }],
]));

// Вводный абзац
children.push(p(
  "Благодарю Вас за встречи и конструктивное обсуждение коммерческих условий сотрудничества. " +
  "Чтобы Вы лично оценили мощь производства и убедились в надёжности «Эмпилс» как долгосрочного партнёра, " +
  "приглашаю Вас посетить наш завод в Ростове-на-Дону в удобное для Вас время."
));

// === Выделенный блок расходов (с гала-ужином) ===
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
            spacing: { line: 312, after: 60 },
            children: [
              new TextRun({ text: "✓  ", size: 22, color: c(P.primary), bold: true, font: FONT }),
              new TextRun({ text: "трансфер из аэропорта и обратно", size: 22, color: c(P.body), font: FONT }),
            ],
          }),
          new Paragraph({
            spacing: { line: 312, after: 40 },
            children: [
              new TextRun({ text: "✓  ", size: 22, color: c(P.primary), bold: true, font: FONT }),
              new TextRun({ text: "гала-ужин", size: 22, bold: true, color: c(P.primary), font: FONT }),
              new TextRun({ text: " в ресторане в вечер визита — неформальное общение с руководством «Эмпилс»", size: 22, color: c(P.body), font: FONT }),
            ],
          }),
        ],
      })],
    }),
  ],
}));

children.push(empty());

// Программа визита
children.push(pRich([
  ["Что Вы увидите за один день:", { bold: true, color: P.primary, size: 24 }],
]));

children.push(li([
  new TextRun({ text: "Экскурсия по производству", size: 22, bold: true, color: c(P.body), font: FONT }),
]));

children.push(li([
  new TextRun({ text: "Презентация ассортимента под Ваш регион", size: 22, bold: true, color: c(P.body), font: FONT }),
]));

children.push(li([
  new TextRun({ text: "Практический тест ЛКМ", size: 22, bold: true, color: c(P.body), font: FONT }),
]));

children.push(li([
  new TextRun({ text: "Обсуждение условий и фиксация договорённостей", size: 22, bold: true, color: c(P.body), font: FONT }),
]));

children.push(empty());

// Призыв к действию
children.push(p(
  "Подтвердите, пожалуйста, удобные Вам даты визита в период ноября месяца 2026 года."
));

// Заключительный абзац
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
    new TextRun({ text: "________________________", size: 22, color: c(P.secondary), font: FONT }),
  ],
}));

children.push(new Paragraph({
  spacing: { line: 312, after: 0 },
  children: [
    new TextRun({ text: "Директор по продажам ЗАО «Эмпилс»", size: 22, bold: true, color: c(P.primary), font: FONT }),
  ],
}));

// === Документ ===
const doc = new Document({
  creator: "Z.ai",
  title: "Приглашение посетить завод Эмпилс — Алексей Анатольевич (Олимп) v3",
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
  const outPath = "/home/z/my-project/download/Приглашение_Олимп_Алексей_Анатольевич_v3.docx";
  fs.writeFileSync(outPath, buffer);
  console.log(`✓ Документ сохранён: ${outPath}`);
  console.log(`  Размер: ${buffer.length.toLocaleString()} байт`);
});
