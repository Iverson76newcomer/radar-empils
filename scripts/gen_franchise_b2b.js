/**
 * Документ 3: «Франшиза TITAN — Программа работы с B2B-партнёрами»
 *
 * Для: производители металлоконструкций, строительные компании, промышленные предприятия
 * Отличия:
 * - Фокус на совокупной стоимости владения (ССВ/TCO)
 * - Пробники 0,4 л для тестирования
 * - Крупная фасовка (8,5 л / 10 кг) — максимальная маржа
 * - Портфолио B2B-объектов
 * - Региональный B2B-инфлюенсер
 * - Прямой контакт Эмпилс с B2B-клиентом
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
  children: [new TextRun({ text: "Программа работы с B2B-партнёрами", size: 26, color: c(P.secondary), font: FONT, italics: true })] }));
children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 },
  border: { top: { style: BorderStyle.SINGLE, size: 8, color: c(P.accent), space: 8 } }, children: [] }));
children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 },
  children: [new TextRun({ text: "\u00ABДороже банка \u2014 дешевле покраска\u00BB", size: 22, color: c(P.accent), font: FONT, italics: true })] }));
children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 },
  children: [new TextRun({ text: "Ростов-на-Дону \u00B7 2026", size: 18, color: c(P.secondary), font: FONT })] }));

// 1. О программе
children.push(h1("1. О программе TITAN B2B"));
children.push(body("Программа TITAN B2B \u2014 это система работы с производителями металлоконструкций, строительными компаниями и промышленными предприятиями, которые используют грунт-эмали для защиты металла от коррозии. Программа построена на принципе совокупной стоимости владения (ССВ): TITAN стоит дороже за банку, но дешевле за год эксплуатации."));

children.push(bulletBold("Главный принцип", "Мы продаём не банку краски, а решение, которое окупается за 1 год и служит 10 лет."));
children.push(bulletBold("Целевая аудитория", "Производители металлоконструкций (беседки, заборы, ангары), строительные компании, промышленные предприятия (трубопроводы, резервуары), автосервисы."));

// 2. Почему TITAN для B2B — расчёт ССВ
children.push(h1("2. Совокупная стоимость владения (ССВ)"));
children.push(body("Главный аргумент для B2B-клиента \u2014 расчёт совокупной стоимости владения. TITAN дороже при покупке, но дешевле при эксплуатации."));

children.push(buildTable(
  ["Параметр", "ПФ-115 (конкурент)", "TITAN 3\u04321"],
  [
    ["Цена за банку (0,75 л)", "~250 \u20BD", "~494 \u20BD"],
    ["Срок службы покрытия", "3 года", "10 лет"],
    ["Стоимость 1 года эксплуатации", "~83 \u20BD/год", "~49 \u20BD/год"],
    ["Экономия за 10 лет", "\u2014", "~340 \u20BD на 1 банку"],
    ["Затраты на перекраску за 10 лет", "3 перекраски", "0 перекрасок"],
    ["Стоимость простоя/демонтажа", "3 \u00D7 стоимость работ", "0"],
    ["Итог: TITAN дешевле в ___ раз", "\u2014", "\u00D72,6 по стоимости года"],
  ],
  [35, 30, 35]
));

// 3. Инструменты для B2B-партнёров
children.push(h1("3. Инструменты для привлечения B2B-партнёров"));
children.push(body("Что Эмпилс делает, чтобы B2B-партнёры захотели работать с TITAN."));

children.push(buildTable(
  ["\u2116", "Инструмент", "Описание", "Эффект"],
  [
    ["1", "Доходность для дистрибьютора", "Расчёт маржинальности B2B-канала. Демонстрируем, что работа с B2B-партнёрами выгодна \u2014 доходность выше, чем в рознице", "Мотивация дистрибьютора продвигать B2B"],
    ["2", "Пробник TITAN 0,4 л (глянец)", "Тестовый образец для B2B-партнёра. Малая фасовка \u2014 низкий порог входа. Партнёр пробует продукт без больших затрат", "Снятие страха \u00ABкуплю и не подойдёт\u00BB"],
    ["3", "Листовка для B2B-сегмента", "Отдельная листовка с расчётом выгоды TITAN для B2B-клиента: срок службы 10 лет, ССВ, экономия vs ПФ-115. Фокус на окупаемость", "Обоснование ценности для B2B"],
    ["4", "Смещение к крупной фасовке", "B2B-клиенту выгоднее 8,5 л / 10 кг, чем 0,75 л. Дистрибьютор зарабатывает больше на крупной фасовке \u2014 мотивируем продвигать именно её", "Максимальная маржа дистрибьютора"],
    ["5", "Региональный B2B-инфлюенсер", "Выявление регионального лидера мнений в сфере B2B (производитель металлоконструкций, владелец цеха). Его рекомендация работает лучше рекламы", "Доверительная рекомендация"],
    ["6", "Калькулятор маржинальности", "Excel-калькулятор для дистрибьютора: вводит объём закупки и фасовку \u2192 получает расчёт прибыли по B2B-каналу", "Прозрачность доходности"],
    ["7", "Прямой контакт с B2B-клиентом", "Эмпилс подключается к переговорам дистрибьютора с B2B-клиентом. Техническая экспертиза, демо на объекте, расчёт ССВ", "Поддержка на сложных сделках"],
  ],
  [5, 20, 50, 25]
));

// 4. Фасовки и маржинальность
children.push(h1("4. Фасовки и маржинальность"));
children.push(body("Для B2B-сегмента приоритет \u2014 крупная фасовка. Это даёт максимальную маржу дистрибьютору и минимальную стоимость покраски для B2B-клиента."));

children.push(buildTable(
  ["Фасовка", "Цена Эмпилс", "Маржа дистрибьютора (25%)", "Цена B2B-клиента", "Маржа на 1 банку"],
  [
    ["0,75 л (3\u04321)", "494 \u20BD", "124 \u20BD", "617 \u20BD", "Низкая"],
    ["2 л (3\u04321)", "1 273 \u20BD", "318 \u20BD", "1 591 \u20BD", "Средняя"],
    ["0,8 кг (молотковая)", "701 \u20BD", "175 \u20BD", "876 \u20BD", "Средняя"],
    ["2,5 кг (молотковая)", "2 076 \u20BD", "519 \u20BD", "2 595 \u20BD", "Высокая"],
    ["8,5 л / 10 кг (B2B)", "По прайсу", "Максимальная", "По согласованию", "Максимальная"],
  ],
  [20, 15, 20, 20, 25]
));

children.push(bulletBold("Рекомендация", "Для B2B-канала \u2014 фокус на фасовки 2 л, 2,5 кг и 8,5 л. Это даёт максимальную абсолютную прибыль на единицу и снижает количество транзакций."));

// 5. Портфолио B2B-объектов
children.push(h1("5. Портфолио B2B-объектов"));
children.push(body("Эмпилс предоставляет портфолио реальных объектов, где применялся TITAN, с расчётом экономических параметров."));

children.push(buildTable(
  ["Объект", "Продукт TITAN", "Площадь", "Экономия vs ПФ-115"],
  [
    ["Ворота (заборные)", "Грунт-эмаль 3\u04321", "~6 м\u00B2", "2,6\u00D7 за 10 лет"],
    ["Дверь металлическая", "Быстросохнущая эмаль", "~2 м\u00B2", "2,1\u00D7 за 10 лет"],
    ["Гараж металлический", "Грунт-эмаль 3\u04321", "~12 м\u00B2", "2,6\u00D7 за 10 лет"],
    ["Трубы / резервуары", "Грунт-эмаль 3\u04321 (пром.)", "По проекту", "По расчёту"],
    ["Беседки / навесы", "Грунт-эмаль 3\u04321", "~4 м\u00B2", "2,6\u00D7 за 10 лет"],
  ],
  [25, 25, 20, 30]
));

// 6. Этапы работы с B2B-партнёром
children.push(h1("6. Этапы работы с B2B-партнёром"));

children.push(buildTable(
  ["Этап", "Цель", "Что делает Эмпилс", "Что делает партнёр", "Срок"],
  [
    ["1 \u2014 Знакомство", "Презентовать TITAN и расчёт ССВ", "Визит к B2B-клиенту, передача пробника 0,4 л, листовки B2B, расчёт маржинальности", "Тест продукта на своём объекте, оценка качества", "1-2 недели"],
    ["2 \u2014 Пробная закупка", "Первый заказ TITAN", "Согласование фасовки и цены, техническая поддержка, расчёт ССВ под конкретный объект", "Закупка пробной партии (крупная фасовка), нанесение на реальный объект", "2-4 недели"],
    ["3 \u2014 Регулярные поставки", "Постоянное сотрудничество", "Закрепление менеджера, техподдержка, регулярные перерасчёты ССВ, включение в портфолио", "Регулярные закупки TITAN, рекомендация партнёрам, кейсы", "Постоянно"],
    ["4 \u2014 Инфлюенсер", "B2B-клиент рекомендует TITAN", "Съёмка видеоотзыва, включение в портфолио, PR-поддержка", "Публичная рекомендация TITAN, участие в кейсах Эмпилс", "6-12 месяцев"],
  ],
  [12, 15, 30, 30, 13]
));

// 7. Обязательства
children.push(h1("7. Обязательства сторон"));

children.push(h2("Эмпилс обязуется:"));
children.push(bullet("Поставлять продукцию TITAN в согласованные сроки"));
children.push(bullet("Предоставлять пробники 0,4 л и веерные выкрасы для тестирования \u2014 бесплатно"));
children.push(bullet("Предоставлять листовки B2B с расчётом ССВ \u2014 бесплатно"));
children.push(bullet("Обеспечивать техническую поддержку (расчёт ССВ под объект, консультации по нанесению)"));
children.push(bullet("Подключаться к переговорам с B2B-клиентами (техническая экспертиза, демо на объекте)"));
children.push(bullet("Назначить персонального менеджера для B2B-канала"));
children.push(bullet("Предоставлять калькулятор маржинальности (Excel)"));

children.push(h2("B2B-партнёр / дистрибьютор обязуется:"));
children.push(bullet("Обеспечить приоритетное продвижение TITAN в B2B-сегменте"));
children.push(bullet("Использовать крупную фасовку (2 л, 2,5 кг, 8,5 л) как основную для B2B-канала"));
children.push(bullet("Соблюдать рекомендованные цены (допустимое отклонение \u00B1__%)"));
children.push(bullet("Предоставлять обратную связь по качеству и применению TITAN на объектах"));
children.push(bullet("При достижении этапа 4 \u2014 согласовать съёмку видеоотзыва и включение в портфолио Эмпилс"));

children.push(divider());

children.push(new Paragraph({ spacing: { before: 400, after: 100 }, alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "\u00ABДороже банка \u2014 дешевле покраска\u00BB", size: 24, bold: true, color: c(P.accent), font: FONT })] }));
children.push(new Paragraph({ spacing: { after: 200 }, alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "B2B-портфолио TITAN \u2014 самый маржинальный канал дистрибьютора", size: 18, italics: true, color: c(P.secondary), font: FONT })] }));
children.push(new Paragraph({ spacing: { after: 100 }, alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "ЗАО \u00ABЭмпилс\u00BB \u00B7 г. Ростов-на-Дону \u00B7 2026", size: 16, color: c(P.secondary), font: FONT })] }));

const doc = new Document({
  creator: "ЗАО Эмпилс",
  title: "Франшиза TITAN — Программа работы с B2B-партнёрами",
  styles: { default: { document: { run: { font: FONT, size: 22, color: c(P.body) }, paragraph: { spacing: { line: 280 } } },
    heading1: { run: { font: FONT, size: 28, bold: true, color: c(P.primary) } },
    heading2: { run: { font: FONT, size: 22, bold: true, color: c(P.primary) } },
  }},
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1200, bottom: 1200, left: 1417, right: 1134 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
      children: [new TextRun({ text: "Эмпилс \u00B7 TITAN \u00B7 Франшиза \u2014 B2B-партнёрам", size: 16, color: "808080", font: FONT })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: [PageNumber.CURRENT], size: 16, color: "808080", font: FONT })] })] }) },
    children,
  }],
});

Packer.toBuffer(doc).then(buf => {
  const OUT = "/home/z/my-project/download/TITAN_Франшиза_B2B_партнёрам.docx";
  fs.writeFileSync(OUT, buf);
  console.log("\u2705 Документ 3 сохранён:", OUT);
  console.log("   Размер:", buf.length.toLocaleString(), "байт");
}).catch(err => { console.error("\u274C", err); process.exit(1); });
