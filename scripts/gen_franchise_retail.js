/**
 * Документ 2: «Франшиза TITAN — Программа работы с федеральными сетями»
 *
 * Для: Лемана Про (бывш. Леруа Мерлен), Петрович, Максидом, OBI, Мегаполис
 * Отличия от дистрибьюторского:
 * - Фокус на планограммы по классам ТТ (A/B/C/D)
 * - Промо-механики для сетей (подарок за покупку, кешбэк, мастер-классы)
 * - Реклама на экранах города, таргет по ТТ
 * - Брендирование торгового оборудования
 */

const {
  Document, Packer, Paragraph, TextRun, Header, Footer,
  AlignmentType, HeadingLevel, PageNumber,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  TableLayoutType,
} = require("docx");
const fs = require("fs");

const P = {
  primary: "1A2330", body: "1F2937", secondary: "6B7280",
  accent: "D4875A", accentLight: "F8F0EB",
  surface: "F5F5F5", white: "FFFFFF",
  positive: "2E7D32", danger: "C62828",
};
const c = (hex) => hex.replace("#", "");
const FONT = { ascii: "Calibri", eastAsia: "Microsoft YaHei" };
const NB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };

function h1(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 180, line: 280 },
    children: [new TextRun({ text, bold: true, size: 28, color: c(P.primary), font: FONT })] });
}
function h2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 280, after: 120, line: 260 },
    children: [new TextRun({ text, bold: true, size: 22, color: c(P.primary), font: FONT })] });
}
function body(text) {
  return new Paragraph({ alignment: AlignmentType.JUSTIFIED, indent: { firstLine: 420 }, spacing: { line: 280, after: 100 },
    children: [new TextRun({ text, size: 22, color: c(P.body), font: FONT })] });
}
function bullet(text) {
  return new Paragraph({ indent: { left: 360, hanging: 200 }, spacing: { line: 260, after: 60 },
    children: [new TextRun({ text: "\u25B8 ", size: 22, color: c(P.accent), bold: true, font: FONT }),
               new TextRun({ text, size: 22, color: c(P.body), font: FONT })] });
}
function bulletBold(lead, rest) {
  return new Paragraph({ indent: { left: 360, hanging: 200 }, spacing: { line: 260, after: 60 },
    children: [new TextRun({ text: "\u25B8 ", size: 22, color: c(P.accent), bold: true, font: FONT }),
               new TextRun({ text: lead, size: 22, bold: true, color: c(P.primary), font: FONT }),
               new TextRun({ text: " \u2014 " + rest, size: 22, color: c(P.body), font: FONT })] });
}
function divider() {
  return new Paragraph({ spacing: { before: 120, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: c(P.accent), space: 1 } }, children: [] });
}
function buildTable(headerCells, rows, colWidths) {
  const headerRow = new TableRow({ tableHeader: true, cantSplit: true,
    children: headerCells.map((text, ci) => new TableCell({
      shading: { type: ShadingType.CLEAR, fill: c(P.primary) },
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      width: colWidths ? { size: colWidths[ci], type: WidthType.PERCENTAGE } : undefined,
      borders: { top: { style: BorderStyle.SINGLE, size: 4, color: c(P.primary) }, bottom: { style: BorderStyle.SINGLE, size: 4, color: c(P.primary) }, left: NB, right: NB },
      children: [new Paragraph({ spacing: { line: 240 }, children: [new TextRun({ text, bold: true, size: 18, color: c(P.white), font: FONT })] })],
    })),
  });
  const dataRows = rows.map((row, idx) => new TableRow({ cantSplit: true,
    children: row.map((text, ci) => new TableCell({
      shading: { type: ShadingType.CLEAR, fill: idx % 2 === 0 ? "FAFAFA" : c(P.white) },
      margins: { top: 50, bottom: 50, left: 100, right: 100 },
      width: colWidths ? { size: colWidths[ci], type: WidthType.PERCENTAGE } : undefined,
      borders: { top: NB, bottom: { style: BorderStyle.SINGLE, size: 2, color: "E0E0E0" }, left: NB, right: NB },
      children: [new Paragraph({ spacing: { line: 240 }, children: [new TextRun({ text, size: 20, bold: ci === 0, color: c(P.body), font: FONT })] })],
    })),
  }));
  return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, layout: TableLayoutType.FIXED,
    borders: { top: { style: BorderStyle.SINGLE, size: 4, color: c(P.primary) }, bottom: { style: BorderStyle.SINGLE, size: 4, color: c(P.primary) }, left: NB, right: NB, insideHorizontal: NB, insideVertical: NB },
    rows: [headerRow, ...dataRows] });
}

const children = [];

// Обложка
children.push(new Paragraph({ spacing: { before: 600, after: 200 }, alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "ЗАО \u00ABЭМПИЛС\u00BB", size: 24, bold: true, color: c(P.accent), font: FONT, characterSpacing: 40 })] }));
children.push(new Paragraph({ spacing: { after: 100 }, alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "ФРАНШИЗА TITAN", size: 44, bold: true, color: c(P.primary), font: FONT })] }));
children.push(new Paragraph({ spacing: { after: 400 }, alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "Программа работы с федеральными сетями", size: 26, color: c(P.secondary), font: FONT, italics: true })] }));
children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 },
  border: { top: { style: BorderStyle.SINGLE, size: 8, color: c(P.accent), space: 8 } }, children: [] }));
children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 },
  children: [new TextRun({ text: "\u00ABКрасишь 1 раз \u2014 служит 10 лет\u00BB", size: 22, color: c(P.accent), font: FONT, italics: true })] }));
children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 },
  children: [new TextRun({ text: "Ростов-на-Дону \u00B7 2026", size: 18, color: c(P.secondary), font: FONT })] }));

// 1. Базовые условия для федеральных сетей
children.push(h1("1. Базовые условия для федеральных сетей"));
children.push(body("ЗАО \u00ABЭмпилс\u00BB предлагает сотрудничество по продукции TITAN для федеральных DIY-сетей на специальных условиях, адаптированных под специфику работы с крупным ритейлом. Программа учитывает требования сетей к планограммам, промо-механикам, маркетинговой поддержке и отчётности."));

children.push(buildTable(
  ["Параметр", "Условие"],
  [
    ["Тип договора", "Поставочный / рамочный с сетью"],
    ["Тип цен", "Сетевой прайс (ниже дистрибьюторского)"],
    ["Отгрузка", "Через распределительный центр сети или напрямую в ТТ"],
    ["Отсрочка платежа", "По стандарту сети (до 60-90 дней)"],
    ["MML TITAN", "По классам точек: A (16-21 SKU), B (12-16), C (2-4), D (0)"],
    ["Брендирование", "Стеллажи TITAN, подиумы, stopper-бочки \u2014 по согласованию"],
    ["Промо-активности", "Подарок за покупку, кешбэк, мастер-классы \u2014 совместно с сетью"],
    ["Цифровое продвижение", "Я.Директ по геолокации ТТ, RuTube, VK таргетинг"],
    ["Отчётность", "CIS-Link + отчётность по стандарту сети (sell-in / sell-out)"],
    ["Бонус за выполнение плана", "Ретро-бонус по итогам периода (согласуется)"],
  ],
  [30, 70]
));

// 2. MML по классам точек
children.push(h1("2. MML TITAN по классам торговых точек"));
children.push(body("Ассортиментная матрица TITAN адаптирована под 4 класса торговых точек. Это обеспечивает оптимальную глубину выкладки в зависимости от формата магазина."));

children.push(buildTable(
  ["Класс ТТ", "Канал", "SKU TITAN", "3\u04321", "Быстросох.", "Молотк."],
  [
    ["A \u2014 Федеральная DIY", "Лемана, Петрович, Максидом", "16 \u2192 21", "10", "2", "4"],
    ["B \u2014 Локальная сеть", "Региональные сети", "12 \u2192 16", "9", "1", "2"],
    ["C \u2014 Супермаркет", "Магазины у дома", "2 \u2192 4", "0", "0", "2"],
    ["D \u2014 Хозмаг", "Мелкие точки", "0", "0", "0", "0"],
  ],
  [20, 25, 15, 13, 13, 14]
));

// 3. Промо-механики для сетей
children.push(h1("3. Промо-механики для федеральных сетей"));
children.push(body("Набор инструментов для стимулирования продаж в торговых точках сети. Протестировано в 11 ТТ Лемана Про (ЮФО) в июле-августе 2026 г."));

children.push(buildTable(
  ["\u2116", "Инструмент", "Описание", "Эффект"],
  [
    ["1", "Промо \u00ABПодарок за покупку\u00BB", "При покупке TITAN от 2000 \u20BD \u2014 кисть + перчатки; от 4000 \u20BD \u2014 сертификат 500 \u20BD. Промоутеры у полки и в закассовой зоне", "Инициация пробной покупки, рост среднего чека"],
    ["2", "Кешбэк 5% на TITAN", "Кешбэк на карту лояльности сети при покупке TITAN во всех ТТ программы", "Стимулирование повторных покупок"],
    ["3", "Мастер-классы Эмпилс", "Выезд команды Эмпилс в 3-5 ТТ. Демонстрация выкрасов с покупателем, презентация свойств TITAN. 2 дня на город", "Обучение покупателей + сбор обратной связи"],
    ["4", "Реклама на экранах города", "Размещение рекламы TITAN на топовых экранах города присутствия ТТ (6-10 экранов на 1 мес)", "Привлечение трафика в ТТ"],
    ["5", "Таргет VK по геолокации ТТ", "Реклама в VK с геотаргетингом по 5-15 ТТ (радиус 5 км от магазина). Аудитория: \u00ABобустройство и ремонт\u00BB", "Ловля горячего спроса рядом с ТТ"],
    ["6", "Жёлтый ценник / спеццена", "Скидка 15-20% на TITAN в течение 2-4 недель. Жёлтый ценник на полке", "Стимулирование пробной покупки"],
    ["7", "Веерные выкрасы B2B", "Изготовление 13+ веерных выкрасов для B2B-клиентов сети (оптовые покупатели)", "Развитие B2B-канала через сеть"],
  ],
  [5, 20, 50, 25]
));

// 4. POS-материалы
children.push(h1("4. POS-материалы на полке"));
children.push(body("Эмпилс предоставляет полный комплект POS-материалов для выкладки TITAN в торговых точках сети \u2014 бесплатно."));

children.push(bulletBold("Шелфтокер", "На полку с выкладкой TITAN. Слоган \u00ABКрасишь 1 раз \u2014 служит 10 лет!\u00BB + QR-код на titanpaints.ru"));
children.push(bulletBold("Воблер", "Подвижный элемент на полке. Привлекает внимание покупателя к продукту TITAN."));
children.push(bulletBold("Брендированный ценник", "Переводит цену банки в \u00ABцену за год\u00BB. Снимает ценовое возражение."));
children.push(bulletBold("Листовка / портфолио объектов", "Расчёт совокупной стоимости владения. Ворота, дверь, гараж с экономическими параметрами."));
children.push(bulletBold("MML + планограмма", "Унифицированная выкладка TITAN с учётом класса ТТ. 3 полки \u00B7 22+ SKU \u00B7 3 линейки."));
children.push(bulletBold("Ролл-ап", "Для специализированных мероприятий (мастер-классы, промо-дни в ТТ)."));

// 5. Цифровое продвижение
children.push(h1("5. Цифровое продвижение TITAN"));
children.push(body("Эмпилс обеспечивает привлечение внешнего трафика в торговые точки сети через digital-каналы."));

children.push(buildTable(
  ["Канал", "Механика", "Бюджет", "Охват"],
  [
    ["Яндекс.Директ", "Поиск + РСЯ по запросам \u00ABкраска по ржавчине\u00BB, \u00ABгрунт-эмаль 3\u04321\u00BB. Геотаргетинг по ТТ сети", "30 000 \u20BD/мес на регион", "50 000+ показов/мес"],
    ["VK (геотаргетинг)", "Реклама по геолокации ТТ (радиус 5 км). Аудитория: \u00ABобустройство и ремонт\u00BB", "25 000 \u20BD/мес на регион", "30 000+ показов/мес"],
    ["RuTube", "Пре-роллы 15-30 сек перед видео о ремонте, даче, авто. Таргет: мужчины 25-55", "40 000 \u20BD/мес", "100 000+ просмотров/мес"],
    ["Блогер / инфлюенсер", "Сотрудничество с блогером в сфере ЛКМ для продвижения TITAN", "Согласуется", "Доверительная рекомендация"],
    ["Экстремальный тест 365", "TITAN vs 365 дней под открытым небом. Онлайн-камера 24/7 на titanpaints.ru", "Включено", "Публичное доказательство \u00AB10 лет\u00BB"],
  ],
  [15, 40, 20, 25]
));

// 6. KPI и контроль
children.push(h1("6. KPI и контроль"));
children.push(buildTable(
  ["Метрика", "Целевое значение", "Периодичность"],
  [
    ["Доля TITAN в категории \u00ABгрунт-эмали\u00BB сети", "Не менее ___% от категории", "Ежеквартально"],
    ["Количество ТТ с TITAN на полке", "___ ТТ из общего количества сети", "Ежемесячно"],
    ["Точки с POS-материалами", "___% точек с TITAN", "Ежемесячно"],
    ["Sell-out TITAN (шт/мес)", "+___% к базовому периоду", "Ежемесячно"],
    ["Сканирования QR-кода с полки", "___ сканов/мес", "Ежемесячно"],
    ["Уходимость TITAN (оборачиваемость)", "Не менее ___ раз/мес", "Ежемесячно"],
  ],
  [40, 35, 25]
));

// 7. Обязательства
children.push(h1("7. Обязательства сторон"));

children.push(h2("Эмпилс обязуется:"));
children.push(bullet("Поставлять продукцию TITAN в РЦ сети или напрямую в ТТ в согласованные сроки"));
children.push(bullet("Предоставлять POS-материалы, выкрасы, планограммы \u2014 бесплатно"));
children.push(bullet("Обеспечивать цифровое продвижение с геопривязкой к ТТ сети"));
children.push(bullet("Проводить мастер-классы силами команды Эмпилс в приоритетных ТТ"));
children.push(bullet("Финансировать промо-акции совместно с сетью (по согласованию)"));
children.push(bullet("Назначить персонального менеджера для работы с сетью"));

children.push(h2("Сеть обязуется:"));
children.push(bullet("Разместить TITAN на полке согласно согласованной планограмме (MML)"));
children.push(bullet("Обеспечить наличие POS-материалов в ___% ТТ с TITAN"));
children.push(bullet("Предоставлять отчётность sell-out по TITAN (ежемесячно через CIS-Link или систему сети)"));
children.push(bullet("Соблюдать рекомендованную розничную цену (допустимое отклонение \u00B1__%)"));
children.push(bullet("Согласовывать промо-активности с Эмпилс не менее чем за 2 недели до старта"));
children.push(bullet("Обеспечить присутствие промоутеров в ТТ в дни проведения промо-акций"));

children.push(divider());

children.push(new Paragraph({ spacing: { before: 400, after: 100 }, alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "\u00ABКрасишь 1 раз \u2014 служит 10 лет\u00BB", size: 24, bold: true, color: c(P.accent), font: FONT })] }));
children.push(new Paragraph({ spacing: { after: 200 }, alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "Партнёрство, которое увеличивает трафик и средний чек", size: 18, italics: true, color: c(P.secondary), font: FONT })] }));
children.push(new Paragraph({ spacing: { after: 100 }, alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "ЗАО \u00ABЭмпилс\u00BB \u00B7 г. Ростов-на-Дону \u00B7 2026", size: 16, color: c(P.secondary), font: FONT })] }));

const doc = new Document({
  creator: "ЗАО Эмпилс",
  title: "Франшиза TITAN — Программа работы с федеральными сетями",
  styles: { default: { document: { run: { font: FONT, size: 22, color: c(P.body) }, paragraph: { spacing: { line: 280 } } },
    heading1: { run: { font: FONT, size: 28, bold: true, color: c(P.primary) } },
    heading2: { run: { font: FONT, size: 22, bold: true, color: c(P.primary) } },
  }},
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1200, bottom: 1200, left: 1417, right: 1134 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
      children: [new TextRun({ text: "Эмпилс \u00B7 TITAN \u00B7 Франшиза \u2014 Федеральные сети", size: 16, color: "808080", font: FONT })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: [PageNumber.CURRENT], size: 16, color: "808080", font: FONT })] })] }) },
    children,
  }],
});

Packer.toBuffer(doc).then(buf => {
  const OUT = "/home/z/my-project/download/TITAN_Франшиза_Федеральные_сети.docx";
  fs.writeFileSync(OUT, buf);
  console.log("\u2705 Документ 2 сохранён:", OUT);
  console.log("   Размер:", buf.length.toLocaleString(), "байт");
}).catch(err => { console.error("\u274C", err); process.exit(1); });
