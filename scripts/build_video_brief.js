/**
 * Бриф-вопросник для агентства по раскрутке видеоканалов Эмпилс
 * (Rutube, YouTube, VK Видео) с использованием различных видов рекламы
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

// Нумерованный вопрос
function numberedQuestion(num, question, context) {
  const children = [];
  // Вопрос
  children.push(new Paragraph({
    spacing: { before: 240, after: 80, line: 312 },
    keepNext: true,
    children: [
      new TextRun({ text: `${num}. `, bold: true, size: 24, color: c(ACCENT_DARK), font: { ascii: "Arial" } }),
      new TextRun({ text: question, bold: true, size: 24, color: c(BODY_COLOR), font: { ascii: "Arial" } }),
    ],
  }));
  // Контекст
  if (context) {
    children.push(new Paragraph({
      spacing: { after: 120, line: 312 },
      indent: { left: 480 },
      children: [
        new TextRun({ text: "Зачем спрашиваем: ", bold: true, italics: true, size: 22, color: c(ACCENT_DARK), font: { ascii: "Arial" } }),
        new TextRun({ text: context, italics: true, size: 22, color: "5B6B7D", font: { ascii: "Arial" } }),
      ],
    }));
  }
  return children;
}

// ============================================================
// СОДЕРЖАНИЕ
// ============================================================
const blocks = [];

// === 1. Контекст задачи ===
blocks.push(h1("1. Контекст задачи"));
blocks.push(body([
  {t: "ЗАО «Эмпилс» (бренды TITAN, NEWTONE, ОРЕОЛ) планирует запустить собственный видеоканал для продвижения ЛКМ в B2B и B2C сегментах. Цель — системная работа с видео-контентом: обучающие ролики по нанесению, обзоры продуктов, кейсы с реальных объектов, мастер-классы, видео-инструкции для мастеров и прорабов."}
]));
blocks.push(body([
  {t: "Рассматриваются три видеоплатформы: ", b: false},
  {t: "Rutube, YouTube, VK Видео", b: true},
  {t: ". У YouTube есть ограничения для пользователей из РФ, поэтому фокус может смещаться на Rutube и VK Видео. Эмпилс ищет профессиональное агентство, которое возьмёт на себя: (1) стратегию развития канала, (2) техническую настройку и ведение, (3) закупку рекламы разных типов (таргетированная, контекстная, ретаргетинг, посевы, интеграции у блогеров), (4) аналитику и отчётность."}
]));
blocks.push(body([
  {t: "Настоящий документ — ", b: false},
  {t: "бриф-вопросник для агентства", b: true},
  {t: ". Ответы на эти вопросы позволят Эмпилс оценить экспертизу подрядчика, сопоставить предложения на рынке и принять обоснованное решение о выборе партнёра."}
]));

// === 2. Разделы брифа ===
blocks.push(h1("2. Блок 1. Экспертиза и опыт агентства"));
blocks.push(body("Цель блока — понять, чем реально занимается агентство, какой у него опыт в нише ЛКМ/строительной химии/DIY, и есть ли у него кейсы, релевантные нашему продукту."));

[
  ["Расскажите о вашем агентстве: год основания, размер команды, специализация, ключевые клиенты за последние 2 года.",
   "Базовая оценка устойчивости агентства и его специализации."],
  ["Какой процент ваших клиентов — производители строительных материалов, ЛКМ, DIY-брендов? Назовите 2-3 релевантных кейса.",
   "Проверяем, есть ли у агентства опыт именно в нашей нише. ЛКМ — узкая категория, общий опыт FMCG не подходит."],
  ["Какие из платформ — Rutube, YouTube, VK Видео — вы считаете приоритетными для брендов ЛКМ в 2026-2027 годах и почему?",
   "Хотим понять, как агентство видит рынок видеохостинга в РФ и какие у него прогнозы по платформам."],
  ["Есть ли у вас сертификации партнёров платформ (YouTube Premier Partner, Rutube Media Partner, VK Official Partner)?",
   "Сертификации подтверждают прямой доступ к рекламным инструментам платформ, без посредников."],
  ["Какие у вас есть собственные каналы или проекты, которые вы ведёте? Это покажет ваш реальный уровень, а не только клиентский.",
   "Лучший индикатор экспертизы — собственные проекты, а не только клиентские кейсы."],
].forEach((item, i) => {
  numberedQuestion(i + 1, item[0], item[1]).forEach(p => blocks.push(p));
});

// === Блок 2 ===
blocks.push(h1("3. Блок 2. Стратегия и формат канала"));
blocks.push(body("Цель блока — понять, как агентство подходит к разработке контент-стратегии канала: тематические рубрики, частота выхода, формат роликов, длительность, целевая аудитория."));

[
  ["Как вы построите контент-стратегию для канала производителя ЛКМ (TITAN, NEWTONE, ОРЕОЛ)? Какие рубрики предложите?",
   "Хотим увидеть конкретные рубрики: обучение, кейсы, обзоры, тесты, экспертный контент."],
  ["Какой формат роликов вы считаете приоритетным: длинные (8-15 мин), средние (3-5 мин), короткие (Shorts/Reels/Клипы 30-60 сек)? Объясните логику.",
   "Стратегия зависит от целей: длинные для доверия, короткие для охвата."],
  ["Какая оптимальная частота выхода роликов на старте? (1 раз в неделю / 2 раза / чаще?)",
   "Понять, какой объём контента агентство рекомендует производить ежемесячно."],
  ["Кто будет разрабатывать сценарии? С вашей стороны — сценарист с опытом в B2B-сегменте ЛКМ?",
   "Сценарий — ключевое. Хотим понять, кто пишет, какой у него опыт, и есть ли у него понимание продукта."],
  ["Как вы видите роль лица бренда в кадре? Нужен ли нам ведущий (эксперт) или формат без ведущего?",
   "Лицо бренда (эксперт завода, технолог, мастер) — отдельная тема. Это влияет на узнаваемость."],
  ["Будете ли вы привлекать блогеров-интеграторов (маляров, прорабов) для гостевых роликов? Если да — как отбираете?",
   "Гостевые ролики с экспертами повышают доверие. Хотим понять подход агентства к их подбору."],
  ["Как вы подходите к локализации контента — нужны ли ролики на разных языках (узбекский, казахский) для рынков СНГ?",
   "Эмпилс работает в экспорте. Если агентство может делать локализацию — это плюс."],
].forEach((item, i) => {
  numberedQuestion(i + 1, item[0], item[1]).forEach(p => blocks.push(p));
});

// === Блок 3 ===
blocks.push(h1("4. Блок 3. Платформы и техническая настройка"));
blocks.push(body("Цель блока — понять, как агентство технически настраивает каналы на трёх платформах, какие инструменты аналитики использует и как работает с ограничениями YouTube в РФ."));

[
  ["Как вы настраиваете каналы на трёх платформах одновременно (Rutube, YouTube, VK Видео)? Один контент публикуется везде или есть адаптация?",
   "Важно понять подход: единый контент-стратегия или разные стратегии для каждой платформы."],
  ["Как вы работаете с ограничениями YouTube для пользователей из РФ? Какие функции (монетизация, реклама, Shorts) недоступны и чем вы это компенсируете?",
   "Российские бренды не могут монетизировать YouTube напрямую. Хотим понять, как агентство это решает."],
  ["Какие инструменты аналитики вы используете для каждой платформы? (YouTube Studio, Rutube Аналитика, VK Статистика, внешние — Social Blade, TGStat, VidIQ, TubeBuddy)",
   "Аналитика — основа оптимизации. Хотим понять глубину ваших метрик."],
  ["Как вы работаете с SEO видеороликов (заголовки, описания, теги, хештеги)? Есть ли у вас процесс оптимизации под поисковые алгоритмы платформ?",
   "SEO видео часто недооценивают. Но именно оно приносит органический трафик годами."],
  ["Какие технические требования к качеству видео (разрешение, битрейт, формат)? Снимаете ли вы сами или только монтируете?",
   "Понять, нужен ли нам отдельный продакшн или агентство закрывает всю цепочку."],
  ["Как вы подходите к обложкам (превью) роликов? Кто их делает и как вы их тестируете (A/B)?",
   "Превью влияет на CTR на 30-50%. Хотим понять подход."],
].forEach((item, i) => {
  numberedQuestion(i + 1, item[0], item[1]).forEach(p => blocks.push(p));
});

// === Блок 4 ===
blocks.push(h1("5. Блок 4. Рекламные инструменты и закупка трафика"));
blocks.push(body("Цель блока — понять, какие рекламные инструменты агентство использует для привлечения трафика на видеоканал, как закупает и оптимизирует рекламу, и какие бюджеты нужны на старте. Это ключевой блок для Эмпилс — мы хотим использовать все доступные типы рекламы."));

[
  ["Перечислите все рекламные инструменты, которые вы используете для привлечения трафика на видеоканал: таргетированная реклама, контекстная реклама, ретаргетинг, посевы у блогеров, тизерные сети, нативные интеграции, Shorts-продвижение. По каждому — коротко механизм и средний CPM/CPC.",
   "Хотим увидеть полную карту инструментов, а не только те, которые агентству удобнее продавать."],
  ["Как работает таргетированная реклама в VK Видео и Rutube? Какие аудитории можно таргетировать (пол, возраст, гео, интересы, поведения)?",
   "Таргет на этих платформах отличается от классического VK. Хотим понять глубину."],
  ["Как вы используете ретаргетинг на зрителей канала? Какие сценарии — догоняющая реклама, призыв подписаться, переход в каталог товаров на сайте Эмпилс?",
   "Ретаргетинг — самый дешёвый канал для повторных касаний. Хотим понять подход."],
  ["Как вы работаете с контекстной рекламой в Яндекс.Директ для видео-целей? (Может ли реклама вести на ролик Rutube / VK Видео? Или только на сайт?)",
   "Понять, как контекст приводит трафик на видео, а не только на лендинг."],
  ["Как вы закупаете интеграции у блогеров? У вас есть собственная база или работаете через биржи (GetBlogger, EasyPRbot, LabelUp)?",
   "Понять, как агентство выбирает блогеров и контролирует качество."],
  ["Используете ли вы посевы в Telegram-каналах и VK-сообществах DIY/строительной тематики для привлечения трафика на видео? Как отбираете площадки?",
   "Посевы — дешёвый способ стартового трафика. Хотим понять методологию."],
  ["Как вы относитесь к Shorts/Reels/VK Клипам как к инструменту привлечения трафика на основные длинные ролики? Какая доля бюджета должна идти на короткие форматы?",
   "Короткие форматы — основной канал органического роста в 2026. Хотим понять долю бюджета."],
  ["Какой минимальный месячный рекламный бюджет вы считаете рабочим для канала на старте? Какое соотношение органика / платный трафик вы прогнозируете через 3, 6, 12 месяцев?",
   "Понять вилку бюджета и реалистичные KPI."],
].forEach((item, i) => {
  numberedQuestion(i + 1, item[0], item[1]).forEach(p => blocks.push(p));
});

// === Блок 5 ===
blocks.push(h1("6. Блок 5. Контент-продакшн и креативы"));
blocks.push(body("Цель блока — понять, кто и как делает креативы для рекламы: снимает ли агентство видеоролики для рекламы, делает ли анимацию, какова стоимость производства."));

[
  ["Что входит в ваши услуги по производству контента: съёмка, монтаж, цветокоррекция, саунд-дизайн, субтитры, превью? По каждому — есть ли в штате или аутсорс?",
   "Понять, закрыта ли вся цепочка в агентстве или часть уходит на субподряд."],
  ["Кто пишет сценарии рекламных роликов? Есть ли у вас копирайтер/сценарист с опытом в B2B ЛКМ?",
   "Сценарий — 70% успеха ролика. Хотим понять, кто его пишет."],
  ["Делаете ли вы анимацию (2D, моушн-графика) для объяснения сложных технических моментов (например, «как грунт-эмаль проникает в ржавчину»)?",
   "Анимация критична для объяснения технологии ЛКМ. Если агентство её не делает — нужен доп. подрядчик."],
  ["Какая средняя стоимость производства одного ролика (минуты): обучение, обзор, кейс, реклама? Дайте ориентир по категориям.",
   "Понять вилку стоимости, чтобы заложить в бюджет."],
  ["Как вы работаете с правами на музыку, шрифты, стоковые материалы? Все ли креативы чистые юридически?",
   "Юридическая чистота — частая проблема у небольших агентств."],
  ["Делаете ли вы адаптации одного ролика под разные платформы (вертикальные/горизонтальные, разная длительность для Shorts/основного/клипов)?",
   "Один ролик = 4-5 адаптаций. Это умножает охват."],
].forEach((item, i) => {
  numberedQuestion(i + 1, item[0], item[1]).forEach(p => blocks.push(p));
});

// === Блок 6 ===
blocks.push(h1("7. Блок 6. Аналитика и KPI"));
blocks.push(body("Цель блока — понять, как агентство измеряет результаты и какие метрики считает ключевыми для оценки эффективности канала."));

[
  ["Какие ключевые метрики вы отслеживаете: просмотры, удержание (retention), CTR, подписчики, конверсия в покупку, ROI? Назовите топ-5 для канала бренда ЛКМ.",
   "Понять, какие метрики агентство считает главными — это показывает зрелость подхода."],
  ["Как вы связываете просмотры видео с продажами ЛКМ (на сайте, у дистрибьюторов, на маркетплейсах)? Есть ли у вас опыт сквозной аналитики?",
   "Сквозная аналитика «видео → продажа» — самое сложное. Если агентство умеет — это сильный плюс."],
  ["Как часто вы предоставляете отчёты и в каком формате? (Excel, Looker Studio, Power BI, собственный дашборд?)",
   "Понять формат отчётности и её регулярность."],
  ["Какие KPI вы готовы взять на себя при работе с каналом Эмпилс? (рост подписчиков за N месяцев, средние просмотры на ролик, доля органики и т. д.)",
   "Если агентство готово к обязательствам по KPI — это серьёзный сигнал зрелости."],
  ["Как вы тестируете гипотезы (A/B-тесты обложек, заголовков, форматов)? Опишите процесс.",
   "A/B-тесты — основа роста. Хотим понять подход."],
  ["Как вы работаете с негативом и комментариями на канале? Есть ли у вас модерация?",
   "Модерация важна для репутации бренда. Особенно в DIY, где аудитория активная."],
].forEach((item, i) => {
  numberedQuestion(i + 1, item[0], item[1]).forEach(p => blocks.push(p));
});

// === Блок 7 ===
blocks.push(h1("8. Блок 7. Бюджет и коммерческие условия"));
blocks.push(body("Цель блока — понять структуру вознаграждения агентства и сравнить предложения разных подрядчиков по единой методологии."));

[
  ["Какая у вас модель оплаты: фикс + % от рекламного бюджета, только фикс, только %, гонорар за результат? Объясните плюсы вашей модели для клиента.",
   "Понять, как агентство зарабатывает. Это влияет на мотивацию подрядчика."],
  ["Какая вилка ежемесячного гонорара за ведение канала (стратегия + аналитика + отчётность), без учёта рекламного бюджета?",
   "Базовая услуга — стратегия и ведение. Это фикс."],
  ["Какой процент от рекламного бюджета вы берёте за медиабаинг и оптимизацию? (стандарт рынка — 10-20%)",
   "Медиабаинг-фewire стандартен, но варьируется."],
  ["Какой минимальный рекламный бюджет на старте вы считаете рабочим? (например, 200 тыс. ₽/мес, 500 тыс. ₽/мес?)",
   "Минимальный порог показывает реалистичность агентства."],
  ["Какая стоимость производства одного ролика? (диапазон за 1 минуту готового контента — без рекламы, только производство)",
   "Сравнить с рыночными ставками продакшена."],
  ["Предусмотрен ли trial-период (например, 1 месяц)? Какие условия выхода из договора, если нас не устроит результат?",
   "Trial и условия выхода — критично для долгосрочного партнёрства."],
  ["Какие KPI вы готовы гарантировать в первые 3-6 месяцев? Что происходит, если они не достигнуты?",
   "Если агентство не готово к KPI — это тревожный сигнал."],
  ["Есть ли у вас реферальные скидки или пакетные предложения при работе с несколькими брендами одного клиента? (Эмпилс — 3 бренда: TITAN, NEWTONE, ОРЕОЛ)",
   "Может быть отдельный канал под каждый бренд или один общий."],
].forEach((item, i) => {
  numberedQuestion(i + 1, item[0], item[1]).forEach(p => blocks.push(p));
});

// === Блок 8 ===
blocks.push(h1("9. Блок 8. Команда и процессы"));
blocks.push(body("Цель блока — понять, кто конкретно будет работать с Эмпилс, как устроено взаимодействие и какие процессы внутри агентства."));

[
  ["Кто будет аккаунт-менеджером проекта Эмпилс? Какой у него опыт в B2B сегменте строительных материалов?",
   "Аккаунт — лицо агентства. Важно, чтобы он понимал специфику ЛКМ."],
  ["Какая команда закрепляется за проектом (стратег, продюсер, сценарист, монтажёр, дизайнер)? Все ли в штате?",
   "Понять состав команды и долю субподряда."],
  ["Как выстроен процесс коммуникации с клиентом? (еженедельные созвоны, чат в Telegram, Trello/Asana, отчёты?)",
   "Чёткий процесс коммуникации — основа работы."],
  ["Как быстро вы реагируете на срочные задачи? (SLA на правки, срочные публикации, кризисные ситуации)",
   "SLA — ключевой параметр для оценки зрелости агентства."],
  ["Как вы работаете с обратной связью? (сколько раундов правок включено в стоимость?)",
   "Правки — частый источник конфликтов. Лучше проговорить заранее."],
  ["Как вы передаёте исходники? (Кто владеет правами на контент, если мы решим сменить агентство?)",
   "Права на исходники — критично для смены подрядчика."],
].forEach((item, i) => {
  numberedQuestion(i + 1, item[0], item[1]).forEach(p => blocks.push(p));
});

// === Блок 9 ===
blocks.push(h1("10. Блок 9. Кейсы и референции"));
blocks.push(body("Цель блока — проверить реальный опыт агентства на конкретных проектах, а не на общих словах."));

[
  ["Предоставьте 3-5 кейсов с каналами брендов B2B (необязательно ЛКМ — это может быть DIY, инструменты, стройматериалы). По каждому: ссылка на канал, какие задачи ставились, какие результаты достигнуты, какой бюджет.",
   "Кейсы — главный критерий выбора. Просим конкретику, а не общие слова."],
  ["Можно ли связаться с 2-3 вашими текущими или прошлыми клиентами для рекомендаций?",
   "Прямой отзыв клиента — самый надёжный источник."],
  ["Расскажите о провальном проекте: что не получилось, почему, какие выводы вы сделали? (Это вопрос про зрелость и способность к рефлексии.)",
   "Способность говорить о неудачах — признак зрелого агентства."],
  ["Какой кейс вы считаете своим лучшим и почему? Что в нём было сделано нестандартно?",
   "Лучший кейс показывает реальный потолок агентства."],
  ["Работали ли вы с брендами, у которых был «нулевой» канал на старте? (Если да — покажите, как вывели с 0 до первых 10 000 подписчиков.)",
   "Если у Эмпилс сейчас канала нет — нужен опыт именно с запуском с нуля."],
].forEach((item, i) => {
  numberedQuestion(i + 1, item[0], item[1]).forEach(p => blocks.push(p));
});

// === Блок 10 ===
blocks.push(h1("11. Блок 10. Дополнительные услуги"));
blocks.push(body("Цель блока — понять, какие сопутствующие услуги агентство может предложить и нужны ли они Эмпилс в дополнение к основному контракту."));

[
  ["Помимо видеопродакшна и закупки рекламы — какие ещё услуги вы можете предоставить? (SMM, работа с отзывами, инфлюенс-маркетинг, мониторинг конкурентов, SERM, кризисные коммуникации)",
   "Дополнительные услуги часто упрощают жизнь клиенту — один подрядчик вместо пяти."],
  ["Можете ли вы помочь с разработкой дизайна-кода канала (единый стиль обложек, заставки, переходы, плашки)?",
   "Единый дизайн-код критичен для узнаваемости бренда в ленте."],
  ["Делаете ли вы локализацию контента для других рынков (СНГ, Казахстан, Узбекистан)?",
   "Эмпилс работает в экспорте — локализация контента может быть нужна."],
  ["Есть ли у вас опыт работы с B2B-партнёрами (дистрибьюторами)? Например, обучение партнёров через видео-контент.",
   "B2B-партнёры — отдельная аудитория. Если агентство умеет с ними работать — это плюс."],
  ["Можете ли вы консультировать по продуктовому позиционированию в видео? (Как донести преимущества TITAN 3-в-1 vs ПФ-115; NEWTONE vs масс-маркет?)",
   "Агентство с сильной стратегической экспертизой ценнее, чем просто исполнители."],
  ["Можете ли вы провести аудит существующих конкурентов на этих платформах (Dali, Tikkurila, Pinotex и др.)? Какие выводы вы уже делали?",
   "Конкурентный анализ — стандартное стартовое действие. Если агентство его уже делало — это ускорит старт."],
].forEach((item, i) => {
  numberedQuestion(i + 1, item[0], item[1]).forEach(p => blocks.push(p));
});

// === Блок 11: Сроки и формат ответа ===
blocks.push(h1("12. Формат и сроки ответа"));
blocks.push(body([
  {t: "Просим агентство подготовить ответы в формате: ", b: false},
  {t: "PDF-презентация (15-25 слайдов) + Excel-приложение с бюджетом на 6 и 12 месяцев.", b: true},
  {t: " По возможности — приложить 2-3 кейса в формате отдельных PDF с конкретными цифрами (просмотры, рост подписчиков, стоимость привлечения)."}
]));
blocks.push(body([
  {t: "Срок подготовки ответа: ", b: true},
  {t: "7 рабочих дней с момента получения брифа."}
]));
blocks.push(body([
  {t: "Формат очной встречи: ", b: true},
  {t: "после получения ответов планируем презентацию агентства в офисе Эмпилс (Ростов-на-Дону) или в Zoom — 60-90 минут. Просим подготовить 15-минутную презентацию с ключевыми тезисами из брифа."}
]));

// === Заключение ===
blocks.push(h1("13. Критерии выбора агентства"));
blocks.push(body("Эмпилс будет оценивать предложения по следующим критериям (в порядке убывания веса):"));

const criteria = [
  ["1", "Релевантный опыт в нише (ЛКМ / DIY / строительная химия)", "30%"],
  ["2", "Понимание специфики B2B-аудитории (мастера, прорабы, дистрибьюторы)", "20%"],
  ["3", "Комплексный подход: стратегия + продакшн + закупка + аналитика", "15%"],
  ["4", "Готовность к KPI и обоснованность бюджета", "15%"],
  ["5", "Качество кейсов и рекомендации прошлых клиентов", "10%"],
  ["6", "Стоимость услуг (при прочих равных)", "5%"],
  ["7", "Химия с командой (после очной встречи)", "5%"],
];

const table = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  layout: TableLayoutType.FIXED,
  borders: thinBorders(P.table.innerLine),
  rows: [
    new TableRow({
      tableHeader: true, cantSplit: true,
      children: [
        new TableCell({
          width: { size: 5, type: WidthType.PERCENTAGE },
          shading: { type: ShadingType.CLEAR, fill: P.table.headerBg },
          margins: { top: 100, bottom: 100, left: 120, right: 120 },
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: "№", bold: true, size: 22, color: P.table.headerText, font: { ascii: "Arial" } })],
          })],
        }),
        new TableCell({
          width: { size: 75, type: WidthType.PERCENTAGE },
          shading: { type: ShadingType.CLEAR, fill: P.table.headerBg },
          margins: { top: 100, bottom: 100, left: 120, right: 120 },
          children: [new Paragraph({
            alignment: AlignmentType.LEFT,
            children: [new TextRun({ text: "Критерий", bold: true, size: 22, color: P.table.headerText, font: { ascii: "Arial" } })],
          })],
        }),
        new TableCell({
          width: { size: 20, type: WidthType.PERCENTAGE },
          shading: { type: ShadingType.CLEAR, fill: P.table.headerBg },
          margins: { top: 100, bottom: 100, left: 120, right: 120 },
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: "Вес", bold: true, size: 22, color: P.table.headerText, font: { ascii: "Arial" } })],
          })],
        }),
      ],
    }),
    ...criteria.map((row, ri) => new TableRow({
      cantSplit: true,
      children: row.map((cell, ci) => new TableCell({
        width: { size: [5, 75, 20][ci], type: WidthType.PERCENTAGE },
        shading: ri % 2 === 1 ? { type: ShadingType.CLEAR, fill: P.table.surface } : undefined,
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({
          alignment: ci === 1 ? AlignmentType.LEFT : AlignmentType.CENTER,
          children: [new TextRun({ text: cell, size: 22, color: c(BODY_COLOR), font: { ascii: "Arial" } })],
        })],
      })),
    })),
  ],
});
blocks.push(table);

blocks.push(body([
  {t: "\nДополнительные требования: ", b: true},
  {t: "агентство должно быть готово к долгосрочному партнёрству (12+ месяцев), поскольку развитие видеоканала — это марафон, а не спринт. Ожидаемая длительность первого контракта: 6 месяцев с возможностью продления."}
]));

blocks.push(body([
  {t: "Контакт для связи со стороны Эмпилс: ", b: true},
  {t: "[Имя Фамилия], [Должность] · +7 [телефон] · [email]. Просим присылать материалы в срок до [дата] с темой письма «Ответ на бриф — Видеоканал Эмпилс»."}
]));

// ============================================================
// СБОРКА ДОКУМЕНТА
// ============================================================
const coverConfig = {
  title: "Бриф для агентства по раскрутке видеоканала Эмпилс",
  subtitle: "Список вопросов для агентства — Rutube, YouTube, VK Видео",
  englishLabel: "EMPILS VIDEO BRIEF",
  metaLines: [
    "Заказчик: ЗАО «Эмпилс» (TITAN, NEWTONE, ОРЕОЛ)",
    "Платформы: Rutube, YouTube, VK Видео",
    "Тип документа: Бриф-вопросник для агентства",
    "Версия: 1.0",
    "Дата: сентябрь 2026",
  ],
  footerLeft: "Эмпилс · Бриф для видеоагентства · v1.0",
  footerRight: "Конфиденциально · Внутренний документ",
};

const doc = new Document({
  creator: "ЗАО Эмпилс",
  title: "Бриф для агентства по раскрутке видеоканала Эмпилс",
  description: "Список вопросов для агентства по раскрутке видеоканалов на Rutube, YouTube, VK Видео",
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
              text: "Эмпилс · Бриф для видеоагентства",
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
  const outPath = "/home/z/my-project/download/Эмпилс_Бриф_для_видеоагентства.docx";
  fs.writeFileSync(outPath, buf);
  console.log("OK saved:", outPath);
  console.log("Size:", buf.length, "bytes");
});
