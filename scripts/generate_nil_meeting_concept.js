/**
 * Генерация Word-документа:
 * "Цели, задачи и концепция практической встречи в НИЛ Эмпилс"
 *
 * Сценарий: дистрибьютор, работавший только по органике, хочет открыть
 * розничную сеть с водно-дисперсионными красками 4 SKU.
 * 2 сотрудника партнёра приезжают в НИЛ Эмпилс для практической демонстрации.
 */

const {
  Document, Packer, Paragraph, TextRun, Header, Footer,
  AlignmentType, HeadingLevel, PageNumber, PageBreak,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  TableLayoutType, SectionType, NumberFormat,
} = require("docx");
const fs = require("fs");

// === Палитра GO-1 (Graphite Orange) — близко к фирменному оранжевому Эмпилс ===
const P = {
  bg: "1A2330",           // тёмный фон обложки
  titleColor: "FFFFFF",
  subtitleColor: "B0B8C0",
  metaColor: "90989F",
  footerColor: "687078",
  accent: "D4875A",       // графитово-оранжевый акцент

  // Для тела документа
  primary: "1A2330",
  body: "1F2937",
  secondary: "6B7280",
  surfaceLight: "F8F0EB",
  accentLine: "D4875A",
  innerLine: "DDD0C8",
};

const c = (hex) => hex.replace("#", "");

// === Утилиты ===
const NB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: NB, bottom: NB, left: NB, right: NB };
const allNoBorders = { top: NB, bottom: NB, left: NB, right: NB, insideHorizontal: NB, insideVertical: NB };

// === Хелперы для контента ===

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 180, line: 312 },
    children: [new TextRun({
      text, bold: true, size: 32, color: c(P.primary),
      font: { ascii: "Calibri", eastAsia: "SimHei" },
    })],
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 140, line: 312 },
    children: [new TextRun({
      text, bold: true, size: 26, color: c(P.primary),
      font: { ascii: "Calibri", eastAsia: "SimHei" },
    })],
  });
}

function body(text) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    indent: { firstLine: 480 },
    spacing: { line: 312, after: 120 },
    children: [new TextRun({
      text, size: 24, color: c(P.body),
      font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" },
    })],
  });
}

function bullet(text, level = 0) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    indent: { left: 360 + level * 360, hanging: 240 },
    spacing: { line: 312, after: 80 },
    children: [
      new TextRun({ text: "▸ ", size: 24, color: c(P.accent), bold: true,
        font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" } }),
      new TextRun({ text, size: 24, color: c(P.body),
        font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" } }),
    ],
  });
}

function bulletWithLead(lead, rest, level = 0) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    indent: { left: 360 + level * 360, hanging: 240 },
    spacing: { line: 312, after: 80 },
    children: [
      new TextRun({ text: "▸ ", size: 24, color: c(P.accent), bold: true,
        font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" } }),
      new TextRun({ text: lead, size: 24, color: c(P.primary), bold: true,
        font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" } }),
      new TextRun({ text: " — " + rest, size: 24, color: c(P.body),
        font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" } }),
    ],
  });
}

// === Обложка R1 (Pure Paragraph Left) ===

function buildCoverR1(config) {
  const padL = 1200, padR = 800;
  const accentLeft = { style: BorderStyle.SINGLE, size: 8, color: c(P.accent), space: 12 };
  const children = [];

  // 1. Верхний отступ
  children.push(new Paragraph({ spacing: { before: 2400 } }));

  // 2. Английская метка с нижней границей-акцентом
  children.push(new Paragraph({
    indent: { left: padL, right: padR },
    spacing: { after: 500 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: c(P.accent), space: 8 } },
    children: [new TextRun({
      text: "P A R T N E R   T R A I N I N G   /   N I L   D E M O",
      size: 18, color: c(P.accent),
      font: { ascii: "Calibri", eastAsia: "SimHei" },
      characterSpacing: 40,
    })],
  }));

  // 3. Главный заголовок (3 строки)
  const titleLines = ["Практическая встреча", "в НИЛ Эмпилс:", "концепция и цели"];
  for (let i = 0; i < titleLines.length; i++) {
    children.push(new Paragraph({
      indent: { left: padL },
      spacing: { after: i < titleLines.length - 1 ? 100 : 300, line: 920, lineRule: "atLeast" },
      children: [new TextRun({
        text: titleLines[i], size: 72, bold: true,
        color: c(P.titleColor),
        font: { eastAsia: "SimHei", ascii: "Arial" },
      })],
    }));
  }

  // 4. Подзаголовок
  children.push(new Paragraph({
    indent: { left: padL },
    spacing: { after: 800 },
    children: [new TextRun({
      text: "Демонстрация водно-дисперсионных красок для дистрибьютора, открывающего розничную сеть",
      size: 24, color: c(P.subtitleColor),
      font: { eastAsia: "Microsoft YaHei", ascii: "Arial" },
    })],
  }));

  // 5. Мета-строки с левым акцентом
  const metaLines = [
    "Продуктовый фокус: 4 SKU интерьерных моющихся красок",
    "Состав: 2 сотрудника партнёра + специалист Эмпилс",
    "Формат: практическая демонстрация в лаборатории",
    "Документ: цели · задачи · концепция · программа",
  ];
  for (const line of metaLines) {
    children.push(new Paragraph({
      indent: { left: padL + 200 },
      spacing: { after: 80 },
      border: { left: accentLeft },
      children: [new TextRun({
        text: line, size: 24, color: c(P.metaColor),
        font: { eastAsia: "Microsoft YaHei", ascii: "Arial" },
      })],
    }));
  }

  // 6. Нижний отступ
  children.push(new Paragraph({ spacing: { before: 2200 } }));

  // 7. Футер с верхней границей-акцентом
  children.push(new Paragraph({
    indent: { left: padL, right: padR },
    border: { top: { style: BorderStyle.SINGLE, size: 2, color: c(P.accent), space: 8 } },
    spacing: { before: 200 },
    children: [
      new TextRun({ text: "ЗАО Эмпилс  ·  Ростов-на-Дону", size: 16, color: c(P.footerColor), font: { ascii: "Arial" } }),
      new TextRun({ text: "                                        " }),
      new TextRun({ text: "Конфиденциально · 2026", size: 16, color: c(P.footerColor), font: { ascii: "Arial" } }),
    ],
  }));

  // Обёрточная таблица 16838
  return [new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: allNoBorders,
    rows: [new TableRow({
      height: { value: 16838, rule: "exact" },
      children: [new TableCell({
        shading: { type: ShadingType.CLEAR, fill: c(P.bg) },
        borders: noBorders,
        children,
      })],
    })],
  })];
}

// === Таблица SKU ===

function buildSkuTable() {
  const headerCells = ["Бренд", "Категория", "Ценовой сегмент", "Назначение"];
  const data = [
    ["AvantGarde", "Интерьерная моющаяся", "Премиум", "Гибридные технологии для demanding-объектов"],
    ["NewTone", "Интерьерная моющаяся", "Средний +", "Передовые технологии, современный дизайн"],
    ["Ореол", "Интерьерная моющаяся", "Средний", "Надёжная краска для массового ремонта"],
    ["Ореол Дисконт", "Интерьерная моющаяся", "Эконом", "Доступная цена для DIY-сегмента"],
  ];

  const headerRow = new TableRow({
    tableHeader: true,
    cantSplit: true,
    children: headerCells.map(text => new TableCell({
      shading: { type: ShadingType.CLEAR, fill: c(P.accent) },
      margins: { top: 100, bottom: 100, left: 140, right: 140 },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 4, color: c(P.accent) },
        bottom: { style: BorderStyle.SINGLE, size: 4, color: c(P.accent) },
        left: { style: BorderStyle.NONE },
        right: { style: BorderStyle.NONE },
      },
      children: [new Paragraph({
        spacing: { line: 280 },
        children: [new TextRun({
          text, bold: true, size: 22, color: "FFFFFF",
          font: { ascii: "Calibri", eastAsia: "SimHei" },
        })],
      })],
    })),
  });

  const dataRows = data.map((row, idx) => new TableRow({
    cantSplit: true,
    children: row.map(text => new TableCell({
      shading: idx % 2 === 0
        ? { type: ShadingType.CLEAR, fill: c(P.surfaceLight) }
        : { type: ShadingType.CLEAR, fill: "FFFFFF" },
      margins: { top: 100, bottom: 100, left: 140, right: 140 },
      borders: {
        top: { style: BorderStyle.NONE },
        bottom: { style: BorderStyle.SINGLE, size: 2, color: c(P.innerLine) },
        left: { style: BorderStyle.NONE },
        right: { style: BorderStyle.NONE },
      },
      children: [new Paragraph({
        spacing: { line: 280 },
        children: [new TextRun({
          text, size: 22, color: c(P.body),
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

// === Таблица программы встречи ===

function buildProgramTable() {
  const headerCells = ["Время", "Этап", "Деятельность", "Ответственный"];
  const data = [
    ["10:00–10:30", "Вводная часть", "Знакомство, кофе, цели дня, регламент, техника безопасности в НИЛ", "Менеджер Эмпилс"],
    ["10:30–11:30", "Теоретический блок", "Обзор 4 SKU, позиционирование по сегментам, отличия от органики", "Специалист НИЛ"],
    ["11:30–13:00", "Демонстрация укрывистости", "Подготовка контрастной полосы, нанесение 4 SKU, наблюдение результата", "Специалист НИЛ"],
    ["13:00–14:00", "Обед", "—", "—"],
    ["14:00–15:30", "Тест моющих свойств", "Контрольная мойка высохших образцов через 24 часа (ускоренная сушка)", "Специалист НИЛ"],
    ["15:30–16:30", "Практика партнёра", "Самостоятельное нанесение под контролем специалиста, вопросы", "Сотрудники партнёра"],
    ["16:30–17:00", "Подведение итогов", "Резюме преимуществ, согласование следующих шагов, обратная связь", "Менеджер Эмпилс"],
  ];

  const headerRow = new TableRow({
    tableHeader: true,
    cantSplit: true,
    children: headerCells.map(text => new TableCell({
      shading: { type: ShadingType.CLEAR, fill: c(P.accent) },
      margins: { top: 100, bottom: 100, left: 120, right: 120 },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 4, color: c(P.accent) },
        bottom: { style: BorderStyle.SINGLE, size: 4, color: c(P.accent) },
        left: { style: BorderStyle.NONE },
        right: { style: BorderStyle.NONE },
      },
      children: [new Paragraph({
        spacing: { line: 280 },
        children: [new TextRun({
          text, bold: true, size: 22, color: "FFFFFF",
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
      margins: { top: 100, bottom: 100, left: 120, right: 120 },
      borders: {
        top: { style: BorderStyle.NONE },
        bottom: { style: BorderStyle.SINGLE, size: 2, color: c(P.innerLine) },
        left: { style: BorderStyle.NONE },
        right: { style: BorderStyle.NONE },
      },
      children: [new Paragraph({
        spacing: { line: 280 },
        children: [new TextRun({
          text, size: 21, bold: ci === 0,
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

// === Содержимое тела документа ===

function buildBody() {
  const children = [];

  // ────────────────────────────────────────────────────────────────────
  // 1. ИСХОДНЫЙ КОНТЕКСТ
  // ────────────────────────────────────────────────────────────────────
  children.push(h1("1. Исходный контекст и предпосылки"));

  children.push(body(
    "Действующий дистрибьютор Эмпилс, исторически работающий исключительно с лакокрасочными материалами на органической основе, инициировал открытие собственной розничной сети с ассортиментом водно-дисперсионных красок. Партнёр планирует закупать у Эмпилс четыре SKU в категории «интерьерная моющаяся краска»: AvantGarde, NewTone, Ореол и Ореол Дисконт. Это решение открывает для Эмпилс новый канал сбыта в розничном сегменте и одновременно требует от партнёра перестройки операционной модели — от знания свойств алкидных эмалей к пониманию специфики водных систем."
  ));

  children.push(body(
    "Ключевая сложность перехода: команда партнёра не имеет практического опыта продаж водно-дисперсионных красок. Менеджеры и продавцы не смогут уверенно консультровать розничного клиента по свойствам, расходу, подготовке поверхности и эксплуатации моющихся интерьерных красок. Без целенаправленного обучения запуск розничной сети рискует обернуться возвратами, отрицательными отзывами и потерей доверия к брендам Эмпилс."
  ));

  children.push(body(
    "Для решения этой задачи предлагается формат практической встречи в научно-исследовательской лаборатории (НИЛ) Эмпилс. Два сотрудника партнёра посещают производство, где специалист Эмпилс в реальных условиях демонстрирует свойства четырёх SKU на подготовленных образцах. Концепция встречи — «обучение через видимое доказательство»: партнёр видит свойства краски своими глазами, а не из презентации, и получает готовые аргументы для розничных продаж."
  ));

  // ────────────────────────────────────────────────────────────────────
  // 2. ЦЕЛИ ВСТРЕЧИ
  // ────────────────────────────────────────────────────────────────────
  children.push(h1("2. Цели встречи"));

  children.push(body(
    "Цели сформулированы по принципу SMART: конкретные, измеримые, достижимые в рамках одного рабочего дня, релевантные бизнес-задаче партнёра и ограниченные по времени. Достижение целей оценивается по итогам встречи через устную обратную связь от сотрудников партнёра и согласование следующих шагов."
  ));

  children.push(h2("2.1. Главная цель"));
  children.push(body(
    "Подготовить команду партнёра к самостоятельной розничной продаже четырёх SKU водно-дисперсионных интерьерных моющихся красок Эмпилс — AvantGarde, NewTone, Ореол, Ореол Дисконт — с уверенным знанием ключевых потребительских свойств и аргументов для конечного покупателя."
  ));

  children.push(h2("2.2. Сопутствующие цели"));

  children.push(bulletWithLead(
    "Передать прикладные знания",
    "об особенностях водно-дисперсионных красок в сравнении с органическими эмалями: механика плёнкообразования, требования к поверхности, время высыхания, условия эксплуатации, расход на квадратный метр"
  ));
  children.push(bulletWithLead(
    "Доказать ключевые свойства визуально",
    "укрывистость на контрастной полосе, моющая способность после высыхания, сравнение четырёх SKU между собой в одинаковых условиях — партнёр видит результат, а не верит на слово"
  ));
  children.push(bulletWithLead(
    "Сформировать продажные аргументы",
    "по каждому SKU: для какого клиента, для какого помещения, при каком бюджете, чем обосновать более высокую цену AvantGarde по сравнению с Ореол Дисконт"
  ));
  children.push(bulletWithLead(
    "Получить обратную связь",
    "от партнёра по ассортименту, ценам, упаковке, маркетинговым материалам и наметить следующие шаги по запуску розничной сети"
  ));
  children.push(bulletWithLead(
    "Закрепить отношения",
    "через личный контакт со специалистом НИЛ и менеджером Эмпилс — партнёр видит производство, лабораторию, уровень компетенций команды"
  ));

  // ────────────────────────────────────────────────────────────────────
  // 3. ЗАДАЧИ ВСТРЕЧИ
  // ────────────────────────────────────────────────────────────────────
  children.push(h1("3. Задачи встречи"));

  children.push(body(
    "Задачи — это конкретные шаги, выполняемые в ходе встречи для достижения поставленных целей. Каждая задача имеет ответственного, ожидаемый результат и привязку к этапу программы (см. раздел 5). Задачи сгруппированы по трём блокам: подготовительный, демонстрационный, переговорный."
  ));

  children.push(h2("3.1. Подготовительные задачи (до встречи)"));

  children.push(bulletWithLead(
    "Согласовать состав делегации",
    "от партнёра — 2 человека (рекомендуется: руководитель направления ЛКМ + старший продавец/технолог); от Эмпилс — менеджер по работе с партнёрами + специалист НИЛ"
  ));
  children.push(bulletWithLead(
    "Подготовить образцы",
    "4 панели из шпаклёванного гипсокартона размером не менее 60×40 см с контрастной полосой чёрно-белой окраски для теста укрывистости по каждому SKU"
  ));
  children.push(bulletWithLead(
    "Подготовить материалы и инструмент",
    "банки краски 4 SKU одной партии, валики и кисти, лотки, малярный скотч, средства защиты, расходные материалы для мойки образцов"
  ));
  children.push(bulletWithLead(
    "Подготовить раздаточные материалы",
    "листовка по 4 SKU с расходом, ценой, областью применения; чек-лист «аргументы для розничной продажи»; таблица сравнения с органическими эмалями"
  ));
  children.push(bulletWithLead(
    "Согласовать логистику",
    "дата, время, транспорт, парковка, пропуск на территорию предприятия, обед"
  ));

  children.push(h2("3.2. Демонстрационные задачи (в ходе встречи)"));

  children.push(bulletWithLead(
    "Провести теоретический блок",
    "30 минут: обзор 4 SKU, ценовое позиционирование, отличия водных систем от органических, типичные ошибки при нанесении"
  ));
  children.push(bulletWithLead(
    "Демонстрация укрывистости",
    "специалист наносит каждый SKU на контрастную полосу одним слоем при свидетелях; партнёр наблюдает процесс и фиксирует визуальный результат"
  ));
  children.push(bulletWithLead(
    "Сравнительный анализ",
    "партнёр сравнивает 4 образца между собой: укрывистость, внешний вид плёнки, плотность, глянец/матовость; специалист комментирует различия"
  ));
  children.push(bulletWithLead(
    "Тест моющих свойств",
    "после сушки образцов (ускоренной) провести контрольную мойку жёсткой губкой с моющим средством — продемонстрировать устойчивость плёнки к истиранию"
  ));
  children.push(bulletWithLead(
    "Практика партнёра",
    "сотрудники партнёра самостоятельно наносят краску под контролем специалиста — учатся правильной технике, задают вопросы, чувствуют материал"
  ));
  children.push(bulletWithLead(
    "Q&A сессия",
    "открытые вопросы от партнёра по любым аспектам: от подбора цвета до гарантийных обязательств"
  ));

  children.push(h2("3.3. Переговорные задачи (завершение встречи)"));

  children.push(bulletWithLead(
    "Согласовать ассортиментный план",
    "какие SKU из четырёх включаются в стартовую матрицу розничной сети, какие добавляются позже"
  ));
  children.push(bulletWithLead(
    "Обсудить условия поставок",
    "минимальная партия, отсрочка, логистика, резервный запас, сроки"
  ));
  children.push(bulletWithLead(
    "Спланировать следующую встречу",
    "обучение продавцов розничных точек, установка колеровочного оборудования, поставка POS-материалов"
  ));
  children.push(bulletWithLead(
    "Зафиксировать договорённости",
    "протокол встречи с ответственными и сроками — отправляется партнёру в течение 24 часов"
  ));

  // ────────────────────────────────────────────────────────────────────
  // 4. КОНЦЕПЦИЯ ВСТРЕЧИ
  // ────────────────────────────────────────────────────────────────────
  children.push(h1("4. Концепция встречи"));

  children.push(h2("4.1. Принципы построения"));

  children.push(body(
    "Концепция встречи строится на принципе «обучение через видимое доказательство». Партнёр не слушает лекцию о свойствах краски — он видит, как краска укрывает контрастную полосу одним слоем, как плёнка выдерживает мойку, как четыре SKU визуально отличаются между собой. Это превращает абстрактные маркетинговые обещания в конкретный визуальный опыт, который сотрудник партнёра сможет воспроизвести в разговоре с розничным клиентом."
  ));

  children.push(body(
    "Второй принцип — «практика важнее теории». Из семи часов программы на теорию отводится не более 60 минут, остальное время — демонстрации, тесты, самостоятельная работа партнёра с материалом. Сотрудник, который сам нанёс краску валиком на panel, в десять раз увереннее расскажет о ней клиенту, чем сотрудник, прочитавший брошюру."
  ));

  children.push(body(
    "Третий принцип — «сравнение, а не утверждение». Все четыре SKU демонстрируются параллельно на одинаковых образцах в одинаковых условиях. Партнёр сам видит, чем AvantGarde отличается от Ореол Дисконт — не из прайс-листа, а из поведения краски. Это снимает возражение «почему так дорого» и даёт продавцу готовый аргумент для каждого ценового сегмента."
  ));

  children.push(h2("4.2. Продуктовый фокус"));

  children.push(body(
    "В фокусе встречи — четыре SKU водно-дисперсионных интерьерных моющихся красок. Все четыре относятся к одной категории, но позиционированы в разных ценовых сегментах — от премиума до экономики. Это позволяет партнёру выстроить продуктовую матрицу под разные сегменты розничных клиентов."
  ));

  children.push(buildSkuTable());

  children.push(body(
    "Ключевое свойство категории «интерьерная моющаяся» — устойчивость сухой плёнки к влажной уборке с применением бытовых моющих средств. Это свойство выносится в заголовок всех продажных аргументов и становится центральной осью демонстрации: партнёр видит, что все четыре SKU выдерживают мойку, но различаются по укрывистости, внешнему виду и плотности плёнки."
  ));

  children.push(h2("4.3. Ключевое свойство для демонстрации: укрывистость"));

  children.push(body(
    "Укрывистость — способность краски за один слой полностью скрыть цвет подложки — выбрана центральным демонстрационным свойством по трём причинам. Во-первых, это свойство визуально очевидно: партнёр видит результат сразу после нанесения, без ожидания сушки. Во-вторых, оно напрямую влияет на расход материала и, следовательно, на экономику покраски для конечного клиента. В-третьих, разница между SKU по укрывистости — это главный аргумент для обоснования более высокой цены премиальных продуктов."
  ));

  children.push(body(
    "Для теста укрывистости используется стандартизированная панель: лист шпаклёванного гипсокартона с контрастной полосой чёрно-белой окраски посередине. Специалист наносит краску одним слоем при свидетелях, после чего партнёр визуально оценивает степень укрывания контрастной полосы. Четыре SKU наносятся на одну панель в идентичных условиях — это обеспечивает корректность сравнения."
  ));

  children.push(h2("4.4. Формат и логистика"));

  children.push(bulletWithLead(
    "Состав участников",
    "2 сотрудника от партнёра (руководитель направления + старший продавец/технолог); 2 представителя Эмпилс (менеджер по работе с партнёрами + специалист НИЛ)"
  ));
  children.push(bulletWithLead(
    "Место проведения",
    "научно-исследовательская лаборатория ЗАО Эмпилс, г. Ростов-на-Дону"
  ));
  children.push(bulletWithLead(
    "Длительность",
    "7 часов (10:00–17:00) с перерывом на обед"
  ));
  children.push(bulletWithLead(
    "Язык",
    "русский"
  ));
  children.push(bulletWithLead(
    "Стоимость для партнёра",
    "бесплатно (Эмпилс покрывает материалы, лабораторное время, обед); партнёр оплачивает свой проезд и проживание"
  ));
  children.push(bulletWithLead(
    "Результат",
    "протокол встречи с согласованными следующими шагами + комплект раздаточных материалов"
  ));

  // ────────────────────────────────────────────────────────────────────
  // 5. ПРОГРАММА ВСТРЕЧИ
  // ────────────────────────────────────────────────────────────────────
  children.push(h1("5. Программа встречи"));

  children.push(body(
    "Программа построена по принципу «от теории — к демонстрации — к практике — к договорённостям». Первая треть дня — теоретический фундамент, вторая треть — наблюдение за работой специалиста, третья треть — самостоятельная работа партнёра и подведение итогов."
  ));

  children.push(buildProgramTable());

  // ────────────────────────────────────────────────────────────────────
  // 6. ОЖИДАЕМЫЕ РЕЗУЛЬТАТЫ
  // ────────────────────────────────────────────────────────────────────
  children.push(h1("6. Ожидаемые результаты"));

  children.push(h2("6.1. Для партнёра"));

  children.push(bulletWithLead(
    "Знания",
    "сотрудники партнёра смогут самостоятельно консультировать розничного клиента по 4 SKU, подбирать продукт под задачу и бюджет, отвечать на типичные возражения"
  ));
  children.push(bulletWithLead(
    "Материалы",
    "раздаточные материалы по 4 SKU, сравнительная таблица, чек-лист продажных аргументов — всё готово к использованию в розничной точке"
  ));
  children.push(bulletWithLead(
    "Решение",
    "согласованный ассортиментный план стартовой матрицы розничной сети, понимание ценовых позиций и целевых сегментов"
  ));
  children.push(bulletWithLead(
    "Контакты",
    "прямые контакты специалиста НИЛ и менеджера Эмпилс для оперативных консультаций"
  ));

  children.push(h2("6.2. Для Эмпилс"));

  children.push(bulletWithLead(
    "Запуск нового канала",
    "розничная сеть партнёра открывается с ассортиментом 4 SKU водно-дисперсионных красок Эмпилс"
  ));
  children.push(bulletWithLead(
    "Расширение категории",
    "партнёр, ранее работавший только по органике, начинает закупать водно-дисперсионные краски — рост категории в обороте"
  ));
  children.push(bulletWithLead(
    "Обратная связь",
    "прямая обратная связь от партнёра по продукту, упаковке, ценам, маркетинговым материалам — основа для улучшений"
  ));
  children.push(bulletWithLead(
    "Эталонный кейс",
    "успешный сценарий онбординга органического дистрибьютора в водный сегмент — методологию можно масштабировать на других партнёров"
  ));

  children.push(h2("6.3. Критерии успеха"));

  children.push(body(
    "Встреча считается успешной при одновременном выполнении следующих условий: оба сотрудника партнёра присутствовали на всей программе; они самостоятельно нанесли краску на образцы под контролем специалиста; они могут устно перечислить ключевые отличия четырёх SKU; согласовано время следующей встречи (обучение продавцов розничных точек); протокол встречи с договорённостями отправлен партнёру в течение 24 часов."
  ));

  return children;
}

// === Сборка документа ===

const doc = new Document({
  creator: "ЗАО Эмпилс",
  title: "Практическая встреча в НИЛ Эмпилс — концепция и цели",
  description: "Концепция, цели и задачи практической встречи для дистрибьютора, открывающего розничную сеть с водно-дисперсионными красками",
  styles: {
    default: {
      document: {
        run: {
          font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" },
          size: 24,
          color: c(P.body),
        },
        paragraph: { spacing: { line: 312 } },
      },
    },
  },
  sections: [
    // Обложка
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 0, bottom: 0, left: 0, right: 0 },
        },
      },
      children: buildCoverR1({}),
    },
    // Тело
    {
      properties: {
        type: SectionType.NEXT_PAGE,
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, bottom: 1440, left: 1701, right: 1417 },
        },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({
              text: "Эмпилс · Практическая встреча в НИЛ",
              size: 18, color: "808080",
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
              size: 18, color: "808080",
              font: { ascii: "Calibri" },
            })],
          })],
        }),
      },
      children: buildBody(),
    },
  ],
});

const OUT = "/home/z/my-project/download/Эмпилс_Концепция_встречи_НИЛ.docx";

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log("✅ Документ сохранён:", OUT);
  console.log("   Размер:", buf.length.toLocaleString(), "байт");
}).catch(err => {
  console.error("❌ Ошибка:", err);
  process.exit(1);
});
