/**
 * Компактная версия документа "Цели, задачи и концепция практической встречи в НИЛ Эмпилс"
 * Цель: уместить в 2 страницы A4 (один лист, двусторонняя печать).
 */

const {
  Document, Packer, Paragraph, TextRun, Header, Footer,
  AlignmentType, HeadingLevel, PageNumber,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  TableLayoutType, SectionType,
} = require("docx");
const fs = require("fs");

// === Палитра (Graphite Orange — близко к оранжевому Эмпилс) ===
const P = {
  bg: "1A2330",
  titleColor: "FFFFFF",
  subtitleColor: "B0B8C0",
  metaColor: "90989F",
  footerColor: "687078",
  accent: "D4875A",
  primary: "1A2330",
  body: "1F2937",
  secondary: "6B7280",
  surfaceLight: "F8F0EB",
  innerLine: "DDD0C8",
};

const c = (hex) => hex.replace("#", "");

const NB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: NB, bottom: NB, left: NB, right: NB };
const allNoBorders = { top: NB, bottom: NB, left: NB, right: NB, insideHorizontal: NB, insideVertical: NB };

// === Компактные хелперы ===

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 140, after: 60, line: 240 },
    children: [new TextRun({
      text, bold: true, size: 24, color: c(P.primary),
      font: { ascii: "Calibri", eastAsia: "SimHei" },
    })],
  });
}

function body(text) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    indent: { firstLine: 280 },
    spacing: { line: 240, after: 40 },
    children: [new TextRun({
      text, size: 18, color: c(P.body),
      font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" },
    })],
  });
}

function bullet(lead, rest) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    indent: { left: 240, hanging: 180 },
    spacing: { line: 220, after: 30 },
    children: [
      new TextRun({ text: "▸ ", size: 18, color: c(P.accent), bold: true,
        font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" } }),
      new TextRun({ text: lead, size: 18, color: c(P.primary), bold: true,
        font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" } }),
      new TextRun({ text: " — " + rest, size: 18, color: c(P.body),
        font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" } }),
    ],
  });
}

// === Таблица SKU (компактная) ===

function buildSkuTable() {
  const headerCells = ["Бренд", "Сегмент", "Назначение"];
  const data = [
    ["AvantGarde", "Премиум", "Гибридные технологии, высокий стандарт"],
    ["NewTone", "Средний +", "Передовые технологии, современный дизайн"],
    ["Ореол", "Средний", "Надёжная краска для массового ремонта"],
    ["Ореол Дисконт", "Эконом", "Доступная цена для DIY-сегмента"],
  ];

  const headerRow = new TableRow({
    tableHeader: true,
    cantSplit: true,
    children: headerCells.map(text => new TableCell({
      shading: { type: ShadingType.CLEAR, fill: c(P.accent) },
      margins: { top: 40, bottom: 40, left: 100, right: 100 },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 4, color: c(P.accent) },
        bottom: { style: BorderStyle.SINGLE, size: 4, color: c(P.accent) },
        left: { style: BorderStyle.NONE },
        right: { style: BorderStyle.NONE },
      },
      children: [new Paragraph({
        spacing: { line: 220 },
        children: [new TextRun({
          text, bold: true, size: 16, color: "FFFFFF",
          font: { ascii: "Calibri", eastAsia: "SimHei" },
        })],
      })],
    })),
  });

  const dataRows = data.map((row, idx) => new TableRow({
    cantSplit: true,
    children: row.map((text, ci) => new TableCell({
      shading: idx % 2 === 0
        ? { type: ShadingType.CLEAR, fill: c(P.surfaceLight) }
        : { type: ShadingType.CLEAR, fill: "FFFFFF" },
      margins: { top: 40, bottom: 40, left: 100, right: 100 },
      borders: {
        top: { style: BorderStyle.NONE },
        bottom: { style: BorderStyle.SINGLE, size: 2, color: c(P.innerLine) },
        left: { style: BorderStyle.NONE },
        right: { style: BorderStyle.NONE },
      },
      children: [new Paragraph({
        spacing: { line: 220 },
        children: [new TextRun({
          text, size: 16, bold: ci === 0,
          color: ci === 0 ? c(P.primary) : c(P.body),
          font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" },
        })],
      })],
    })),
  }));

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: c(P.accent) },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: c(P.accent) },
      left: { style: BorderStyle.NONE },
      right: { style: BorderStyle.NONE },
      insideHorizontal: { style: BorderStyle.NONE },
      insideVertical: { style: BorderStyle.NONE },
    },
    rows: [headerRow, ...dataRows],
  });
}

// === Таблица программы (компактная) ===

function buildProgramTable() {
  const headerCells = ["Время", "Этап", "Деятельность"];
  const data = [
    ["10:00–10:30", "Вводная", "Знакомство, цели дня, ТБ в НИЛ"],
    ["10:30–11:30", "Теория", "Обзор 4 SKU, отличия от органики"],
    ["11:30–13:00", "Укрывистость", "Нанесение 4 SKU на контрастную полосу"],
    ["13:00–14:00", "Обед", "—"],
    ["14:00–15:30", "Моющие свойства", "Контрольная мойка высохших образцов"],
    ["15:30–16:30", "Практика партнёра", "Самостоятельное нанесение, Q&A"],
    ["16:30–17:00", "Итоги", "Согласование следующих шагов, протокол"],
  ];

  const headerRow = new TableRow({
    tableHeader: true,
    cantSplit: true,
    children: headerCells.map(text => new TableCell({
      shading: { type: ShadingType.CLEAR, fill: c(P.accent) },
      margins: { top: 40, bottom: 40, left: 100, right: 100 },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 4, color: c(P.accent) },
        bottom: { style: BorderStyle.SINGLE, size: 4, color: c(P.accent) },
        left: { style: BorderStyle.NONE },
        right: { style: BorderStyle.NONE },
      },
      children: [new Paragraph({
        spacing: { line: 220 },
        children: [new TextRun({
          text, bold: true, size: 16, color: "FFFFFF",
          font: { ascii: "Calibri", eastAsia: "SimHei" },
        })],
      })],
    })),
  });

  const dataRows = data.map((row, idx) => new TableRow({
    cantSplit: true,
    children: row.map((text, ci) => new TableCell({
      shading: idx % 2 === 0
        ? { type: ShadingType.CLEAR, fill: c(P.surfaceLight) }
        : { type: ShadingType.CLEAR, fill: "FFFFFF" },
      margins: { top: 40, bottom: 40, left: 100, right: 100 },
      borders: {
        top: { style: BorderStyle.NONE },
        bottom: { style: BorderStyle.SINGLE, size: 2, color: c(P.innerLine) },
        left: { style: BorderStyle.NONE },
        right: { style: BorderStyle.NONE },
      },
      children: [new Paragraph({
        spacing: { line: 220 },
        children: [new TextRun({
          text, size: 16, bold: ci === 0,
          color: ci === 0 ? c(P.primary) : c(P.body),
          font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" },
        })],
      })],
    })),
  }));

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: c(P.accent) },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: c(P.accent) },
      left: { style: BorderStyle.NONE },
      right: { style: BorderStyle.NONE },
      insideHorizontal: { style: BorderStyle.NONE },
      insideVertical: { style: BorderStyle.NONE },
    },
    rows: [headerRow, ...dataRows],
  });
}

// === Тело документа (компактное) ===

function buildBody() {
  const children = [];

  // Шапка документа (вместо отдельной обложки)
  children.push(new Paragraph({
    spacing: { after: 80, line: 240 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: c(P.accent), space: 4 } },
    children: [
      new TextRun({ text: "ПРАКТИЧЕСКАЯ ВСТРЕЧА В НИЛ ЭМПИЛС", size: 22, bold: true, color: c(P.primary),
        font: { ascii: "Calibri", eastAsia: "SimHei" } }),
      new TextRun({ text: "    ·    ", size: 22, color: c(P.accent),
        font: { ascii: "Calibri" } }),
      new TextRun({ text: "Концепция, цели и задачи", size: 22, color: c(P.secondary),
        font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" } }),
    ],
  }));

  children.push(new Paragraph({
    spacing: { after: 140, line: 220 },
    children: [
      new TextRun({ text: "Дистрибьютор, работавший по органике, открывает розничную сеть с водно-дисперсионными красками. 2 сотрудника партнёра посещают НИЛ Эмпилс для практической демонстрации.",
        size: 18, color: c(P.secondary), italics: true,
        font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" } }),
    ],
  }));

  // 1. Контекст
  children.push(h1("1. Контекст"));
  children.push(body(
    "Партнёр, исторически работающий с органическими ЛКМ, инициировал открытие собственной розничной сети с водно-дисперсионными красками четырёх брендов — AvantGarde, NewTone, Ореол, Ореол Дисконт. Команда партнёра не имеет опыта продаж водных систем: без целевого обучения запуск сети рискует обернуться возвратами и потерей доверия к брендам. Предлагаемый формат — практическая встреча в НИЛ Эмпилс, где специалист в реальных условиях демонстрирует свойства красок на подготовленных образцах."
  ));

  // 2. Цели
  children.push(h1("2. Цели встречи"));
  children.push(new Paragraph({
    spacing: { after: 60, line: 220 },
    children: [
      new TextRun({ text: "Главная цель: ", size: 18, bold: true, color: c(P.primary),
        font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" } }),
      new TextRun({ text: "вселить уверенность в команду партнёра по водно-дисперсионному направлению и передать свойства и аргументы интерьерных моющихся красок для трансляции покупателю.",
        size: 18, color: c(P.body),
        font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" } }),
    ],
  }));
  children.push(bullet("Знания", "особенности водных ЛКМ vs органических: плёнкообразование, поверхность, сушка, расход"));
  children.push(bullet("Доказательство", "укрывистость на контрастной полосе, моющие свойства, сравнение 4 SKU"));
  children.push(bullet("Аргументы", "для каждого SKU: целевой клиент, помещение, бюджет, обоснование цены"));
  children.push(bullet("Обратная связь", "по ассортименту, ценам, упаковке, POSM; план следующих шагов"));
  children.push(bullet("Отношения", "партнёр видит производство и НИЛ, прямой контакт со специалистом"));

  // 3. Задачи
  children.push(h1("3. Задачи"));

  children.push(new Paragraph({
    spacing: { before: 60, after: 30, line: 220 },
    children: [new TextRun({ text: "Подготовка (до встречи):", size: 18, bold: true, color: c(P.primary),
      font: { ascii: "Calibri", eastAsia: "SimHei" } })],
  }));
  children.push(bullet("Состав", "2 чел. от партнёра (руководитель + продавец/технолог) + менеджер и специалист НИЛ Эмпилс"));
  children.push(bullet("Образцы", "4 панели шпаклёванного гипсокартона 60×40 см с контрастной чёрно-белой полосой"));
  children.push(bullet("Материалы", "банки 4 SKU одной партии, валики, кисти, лотки, скотч, средства защиты, расходники для мойки"));
  children.push(bullet("Раздатка", "листовка по 4 SKU, чек-лист продажных аргументов, сравнение с органикой"));

  children.push(new Paragraph({
    spacing: { before: 80, after: 30, line: 220 },
    children: [new TextRun({ text: "Демонстрация (в ходе практической встречи):", size: 18, bold: true, color: c(P.primary),
      font: { ascii: "Calibri", eastAsia: "SimHei" } })],
  }));
  children.push(bullet("Теория 30 мин", "обзор 4 SKU, ценовое позиционирование и краткие отличительные характеристики"));
  children.push(bullet("Укрывистость", "специалист наносит каждый SKU на контрастную полосу"));
  children.push(bullet("Сравнение", "партнёр оценивает укрывистость, вид плёнки, степень блеска"));
  children.push(bullet("Мойка", "после сушки — демонстрация моющих свойств"));
  children.push(bullet("Практика", "по желанию — сотрудники партнёра сами наносят краску под контролем специалиста"));
  children.push(bullet("Q&A", "открытые вопросы"));

  children.push(new Paragraph({
    spacing: { before: 80, after: 30, line: 220 },
    children: [new TextRun({ text: "Переговоры (завершение):", size: 18, bold: true, color: c(P.primary),
      font: { ascii: "Calibri", eastAsia: "SimHei" } })],
  }));
  children.push(bullet("Ассортимент", "какие SKU в стартовую матрицу розничной сети"));
  children.push(bullet("Условия", "минимальная партия, отсрочка, логистика, резервный запас"));
  children.push(bullet("Следующая встреча", "обучение продавцов розницы, колеровка, POSM"));
  children.push(bullet("Протокол", "договорённости с ответственными и сроками — в течение 24 часов"));

  // 4. Концепция
  children.push(h1("4. Концепция"));
  children.push(bullet("Видимое доказательство", "партнёр видит свойства краски, а не верит на слово — маркетинг превращается в личный опыт"));
  children.push(bullet("Сравнение, не утверждение", "4 SKU параллельно на одинаковых образцах — партнёр сам видит разницу, это снимает возражение «почему AvantGarde дороже»"));
  children.push(bullet("Фокус свойства", "укрывистость — визуально очевидна, влияет на расход, обосновывает цену премиума"));

  // 5. Продуктовый фокус — таблица SKU
  children.push(new Paragraph({
    spacing: { before: 120, after: 40, line: 220 },
    children: [new TextRun({ text: "Продуктовый фокус — 4 SKU:", size: 18, bold: true, color: c(P.primary),
      font: { ascii: "Calibri", eastAsia: "SimHei" } })],
  }));
  children.push(buildSkuTable());

  // 6. Программа — таблица
  children.push(new Paragraph({
    spacing: { before: 120, after: 40, line: 220 },
    children: [new TextRun({ text: "Программа (10:00–17:00):", size: 18, bold: true, color: c(P.primary),
      font: { ascii: "Calibri", eastAsia: "SimHei" } })],
  }));
  children.push(buildProgramTable());

  // 7. Результаты
  children.push(h1("5. Ожидаемые результаты"));
  children.push(bullet("Для партнёра", "знания для самостоятельной консультации, раздаточные материалы, согласованная стартовая матрица, прямые контакты НИЛ"));
  children.push(bullet("Для Эмпилс", "новый розничный канал, расширение категории водных ЛКМ, обратная связь, эталонный кейс онбординга"));
  children.push(bullet("Критерий успеха", "оба сотрудника прошли всю программу, самостоятельно нанесли краску, могут перечислить отличия 4 SKU, согласовано время следующей встречи"));

  return children;
}

// === Сборка ===

const doc = new Document({
  creator: "ЗАО Эмпилс",
  title: "Практическая встреча в НИЛ Эмпилс — компактная концепция",
  styles: {
    default: {
      document: {
        run: {
          font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" },
          size: 18,
          color: c(P.body),
        },
        paragraph: { spacing: { line: 240 } },
      },
    },
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 900, bottom: 900, left: 1000, right: 900 },
        },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({
              text: "Эмпилс · НИЛ · Концепция встречи",
              size: 14, color: "A0A0A0",
              font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" },
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
              size: 14, color: "A0A0A0",
              font: { ascii: "Calibri" },
            })],
          })],
        }),
      },
      children: buildBody(),
    },
  ],
});

const OUT = "/home/z/my-project/download/Эмпилс_Концепция_встречи_НИЛ_компакт.docx";

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log("✅ Документ сохранён:", OUT);
  console.log("   Размер:", buf.length.toLocaleString(), "байт");
}).catch(err => {
  console.error("❌ Ошибка:", err);
  process.exit(1);
});
