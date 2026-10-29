/**
 * Документ 1: «Франшиза TITAN — Программа работы с новыми партнёрами (дистрибьюторами)»
 *
 * Экстраполирует механизмы из презентации TITAN + базовые условия из дилерского предложения.
 * Официальный стиль. Формат Word.
 *
 * Структура:
 * 1. Базовые условия партнёрства (из дилерского предложения)
 * 2. Механизмы влияния на сбыт — программа TITAN
 *    2.1. POS-материалы на полке (5 инструментов)
 *    2.2. Контент и обучение покупателя (4 инструмента)
 *    2.3. Продвижение и digital (4 инструмента)
 * 3. Этапы работы с партнёром (0→1→2→3)
 * 4. KPI и контроль
 * 5. План «Б» — управление рисками
 * 6. Обязательства сторон
 */

const {
  Document, Packer, Paragraph, TextRun, Header, Footer,
  AlignmentType, HeadingLevel, PageNumber,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  TableLayoutType, SectionType,
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
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 180, line: 280 },
    children: [new TextRun({ text, bold: true, size: 28, color: c(P.primary), font: FONT })],
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 120, line: 260 },
    children: [new TextRun({ text, bold: true, size: 22, color: c(P.primary), font: FONT })],
  });
}

function body(text) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    indent: { firstLine: 420 },
    spacing: { line: 280, after: 100 },
    children: [new TextRun({ text, size: 22, color: c(P.body), font: FONT })],
  });
}

function bullet(text) {
  return new Paragraph({
    indent: { left: 360, hanging: 200 },
    spacing: { line: 260, after: 60 },
    children: [
      new TextRun({ text: "\u25B8 ", size: 22, color: c(P.accent), bold: true, font: FONT }),
      new TextRun({ text, size: 22, color: c(P.body), font: FONT }),
    ],
  });
}

function bulletBold(lead, rest) {
  return new Paragraph({
    indent: { left: 360, hanging: 200 },
    spacing: { line: 260, after: 60 },
    children: [
      new TextRun({ text: "\u25B8 ", size: 22, color: c(P.accent), bold: true, font: FONT }),
      new TextRun({ text: lead, size: 22, bold: true, color: c(P.primary), font: FONT }),
      new TextRun({ text: " \u2014 " + rest, size: 22, color: c(P.body), font: FONT }),
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

// === Универсальная таблица ===
function buildTable(headerCells, rows, colWidths) {
  const headerRow = new TableRow({
    tableHeader: true, cantSplit: true,
    children: headerCells.map((text, ci) => new TableCell({
      shading: { type: ShadingType.CLEAR, fill: c(P.primary) },
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      width: colWidths ? { size: colWidths[ci], type: WidthType.PERCENTAGE } : undefined,
      borders: { top: { style: BorderStyle.SINGLE, size: 4, color: c(P.primary) }, bottom: { style: BorderStyle.SINGLE, size: 4, color: c(P.primary) }, left: NB, right: NB },
      children: [new Paragraph({
        spacing: { line: 240 },
        children: [new TextRun({ text, bold: true, size: 18, color: c(P.white), font: FONT })],
      })],
    })),
  });

  const dataRows = rows.map((row, idx) => new TableRow({
    cantSplit: true,
    children: row.map((text, ci) => new TableCell({
      shading: { type: ShadingType.CLEAR, fill: idx % 2 === 0 ? "FAFAFA" : c(P.white) },
      margins: { top: 50, bottom: 50, left: 100, right: 100 },
      width: colWidths ? { size: colWidths[ci], type: WidthType.PERCENTAGE } : undefined,
      borders: { top: NB, bottom: { style: BorderStyle.SINGLE, size: 2, color: "E0E0E0" }, left: NB, right: NB },
      children: [new Paragraph({
        spacing: { line: 240 },
        children: [new TextRun({ text, size: 20, bold: ci === 0, color: c(P.body), font: FONT })],
      })],
    })),
  }));

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

// === Содержимое ===
const children = [];

// === Обложка-заголовок ===
children.push(new Paragraph({
  spacing: { before: 600, after: 200 },
  alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "ЗАО \u00ABЭМПИЛС\u00BB", size: 24, bold: true, color: c(P.accent), font: FONT, characterSpacing: 40 })],
}));

children.push(new Paragraph({
  spacing: { after: 100 },
  alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "ФРАНШИЗА TITAN", size: 44, bold: true, color: c(P.primary), font: FONT })],
}));

children.push(new Paragraph({
  spacing: { after: 400 },
  alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "Программа работы с новыми дистрибьюторами", size: 26, color: c(P.secondary), font: FONT, italics: true })],
}));

children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 100 },
  border: { top: { style: BorderStyle.SINGLE, size: 8, color: c(P.accent), space: 8 } },
  children: [],
}));

children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 300 },
  children: [new TextRun({ text: "\u00ABКрасишь 1 раз \u2014 служит 10 лет\u00BB", size: 22, color: c(P.accent), font: FONT, italics: true })],
}));

children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 400 },
  children: [new TextRun({ text: "Ростов-на-Дону \u00B7 2026", size: 18, color: c(P.secondary), font: FONT })],
}));

// === Раздел 1: Базовые условия ===
children.push(h1("1. Базовые условия партнёрства"));
children.push(body("ЗАО \u00ABЭмпилс\u00BB предлагает дистрибьюторское сотрудничество по продукции TITAN на следующих базовых условиях. Конкретные параметры (отсрочка, план, ассортимент) согласовываются индивидуально в зависимости от региона, объёма и канала сбыта партнёра."));

children.push(buildTable(
  ["Параметр", "Условие"],
  [
    ["Тип договора", "Дистрибьюторский"],
    ["Тип цен", "Дистрибьюторский прайс"],
    ["География продаж", "Согласованная территория (Эмпилс контролирует географию)"],
    ["Отгрузка товара", "В течение ___ дней с момента заказа"],
    ["Транспортные расходы", "Самовывоз / доставка по согласованию"],
    ["Отсрочка платежа", "___ календарных дней"],
    ["Ассортиментная матрица", "По согласованию сторон (рекомендуемый MML TITAN \u2014 от 12 SKU)"],
    ["Наценка дистрибьютора", "25% (рекомендуемая)"],
    ["Каталоги и материалы", "Каталоги, выкрасы, POS-материалы \u2014 бесплатно"],
    ["Маркетинговая активность", "Готовы обсудить мотивацию ТП и акции для торговых точек"],
    ["Система отчётности", "Установка системы отчётности CIS-Link"],
    ["Бонус за выполнение плана", "Бонус за выполнение месячного плана (согласуется)"],
  ],
  [30, 70]
));

// === Раздел 2: Механизмы влияния на сбыт ===
children.push(h1("2. Механизмы влияния на сбыт \u2014 программа TITAN"));
children.push(body("Программа TITAN \u2014 это не просто поставка продукта, а комплексная система инструментов, которые Эмпилс предоставляет дистрибьютору для увеличения продаж. Все механизмы разработаны на основе реального опыта работы с 109 действующими партнёрами и протестированы в пилотных регионах."));

// 2.1 POS-материалы
children.push(h2("2.1. POS-материалы на полке (5 инструментов)"));
children.push(body("Инструменты, которые работают непосредственно в точке продажи. Эмпилс предоставляет их бесплатно."));

children.push(buildTable(
  ["\u2116", "Инструмент", "Описание", "Эффект"],
  [
    ["1", "Шелфтокер", "На полку с выкладкой TITAN. Слоган \u00ABКрасишь 1 раз \u2014 служит 10 лет!\u00BB + QR-код на titanpaints.ru", "Привлечение внимания + переход на сайт"],
    ["2", "Воблер", "На полку с выкладкой TITAN. Подвижный элемент, привлекающий взгляд покупателя", "Увеличение конверсии в покупку"],
    ["3", "Брендированный ценник", "Переводит цену банки в \u00ABцену за год\u00BB (5720 \u20BD \u2192 532 \u20BD/год при сроке 10 лет)", "Снятие ценового возражения"],
    ["4", "Листовка / портфолио объектов", "Наиболее распространённые объекты с экономическими параметрами совокупной стоимости владения (ворота, дверь, гараж)", "Обоснование ценности для покупателя"],
    ["5", "MML + планограмма", "Унифицированная выкладка TITAN с учётом класса торговых точек (A, B, C)", "Стандартизация выкладки, контроль"],
  ],
  [5, 20, 50, 25]
));

// 2.2 Контент и обучение
children.push(h2("2.2. Контент и обучение покупателя (4 инструмента)"));
children.push(body("Инструменты, которые обучают покупателя и торгового представителя, обеспечивая осознанную покупку."));

children.push(buildTable(
  ["\u2116", "Инструмент", "Описание", "Эффект"],
  [
    ["6", "Инструкция по нанесению (видео)", "Пошаговая быстрая видеоинструкция для конечного потребителя (5 шагов + видео). Доступ по QR-коду", "Покупатель сам учится, без продавца"],
    ["7", "Видеоролики для ТП", "1. Видеопортфолио объектов. 2. Ролик для ТП: УТП, свойства, выгода по TITAN. Скрипт идеальной продажи", "ТП умеет продавать с первого дня"],
    ["8", "Фото/видео POSm на облаке", "Материалы для розничных точек на облаке: исходники, листовка, видеоролики, примеры выкладок", "Быстрый доступ к материалам"],
    ["9", "Контент для соцсетей", "Папка на облаке с библиотекой готовых постов про TITAN (картинка + текст) для SMM дистрибьютора", "Бесплатный контент для соцсетей"],
  ],
  [5, 20, 50, 25]
));

// 2.3 Продвижение и digital
children.push(h2("2.3. Продвижение и digital-маркетинг (4 инструмента)"));
children.push(body("Инструменты, работающие за пределами полки \u2014 привлекают покупателей извне и направляют в точки продаж TITAN."));

children.push(buildTable(
  ["\u2116", "Инструмент", "Описание", "Эффект"],
  [
    ["10", "Экстремальный тест 365", "Объект красится TITAN 3\u04321 и выставляется на 365 дней под открытым небом. Онлайн-камера 24/7 на titanpaints.ru", "Доказательство \u00AB10 лет\u00BB не на словах, а публично"],
    ["11", "Яндекс.Директ / VK", "Контекстная реклама: \u00ABкраска по ржавчине\u00BB, \u00ABгрунт-эмаль 3\u04321\u00BB \u2192 реклама TITAN с указанием магазина партнёра", "Ловля горячего спроса, направление в ТТ"],
    ["12", "Видеореклама на RuTube", "Пре-роллы 15-30 сек перед видео о ремонте, строительстве, даче, авто. Таргет: мужчины 25-55", "Массовый охват целевой аудитории"],
    ["13", "Работа с блогером / инфлюенсером", "Сотрудничество с блогером в сфере ЛКМ для продвижения TITAN", "Доверительная рекомендация"],
  ],
  [5, 20, 50, 25]
));

// === Раздел 3: Этапы работы ===
children.push(h1("3. Этапы работы с новым партнёром"));
children.push(body("Программа TITAN реализуется поэтапно. Каждый этап имеет конкретные цели, инструменты и сроки. Переход к следующему этапу \u2014 после достижения KPI предыдущего."));

children.push(buildTable(
  ["Этап", "Цель", "Что делает Эмпилс", "Что делает партнёр", "Срок"],
  [
    ["0 \u2014 Поиск", "Выйти на контакт и презентовать программу", "Персональный визит РМ, звонок, email, бумажное письмо, приглашение на завод", "Знакомство с программой, посещение завода", "1-2 месяца"],
    ["1 \u2014 Первичные продажи", "Завод \u2192 Дистрибьютор", "Экскурсия на завод, бонус за первичную закупку 5-10%, возврат неликвида 90 дней, отсрочка 70-90 дней, выездное обучение ТП, бесплатные POSm + стартовый набор", "Первая закупка, обучение ТП, выкладка в точках A и Б по планограмме", "0-3 месяца"],
    ["2 \u2014 Вторичные продажи", "Дистрибьютор \u2192 Розница", "Бонус за заполнение планограммы, отсрочка 30 дней для розницы, выездное обучение продавцов РТТ (45 мин), дополнительные POSm, консигнация для топ-5 точек", "Обучение 100% ТП, открытие новых точек (+25%), выкладка в точках В и С, фотоотчёты", "3-6 месяцев"],
    ["3 \u2014 Третичные продажи", "Розница \u2192 Покупатель", "Полный комплект POSm, QR на банке, контент для соцсетей, видеоролики, цифровое продвижение (Я.Директ, RuTube, блогер)", "Наличие POSm в 80% точек, сертификация продавцов, сбор обратной связи", "6-12 месяцев"],
  ],
  [10, 15, 30, 30, 15]
));

// === Раздел 4: KPI ===
children.push(h1("4. KPI и контроль"));
children.push(body("Успех программы измеряется 10 метриками по 3 этапам. Контроль \u2014 ежемесячный через систему отчётности CIS-Link."));

children.push(buildTable(
  ["Этап", "Метрика", "Целевое значение", "Периодичность"],
  [
    ["1", "Объём закупок TITAN", "+___% к базовому периоду", "Ежемесячно"],
    ["1", "Доля TITAN в ассортименте", "Не менее ___ цветов RAL в ___ категориях", "Ежеквартально"],
    ["2", "Количество активных РТТ с TITAN", "+___% к базовому периоду", "Ежемесячно"],
    ["2", "Охват ТП обучением", "___% ТП дистрибьютора", "Ко 2-му месяцу"],
    ["3", "Точки с POS-материалами", "___% точек с TITAN на полке", "К 3-му месяцу"],
    ["3", "Сканирования QR-кода", "___ сканов/мес", "Ежемесячно"],
  ],
  [10, 35, 35, 20]
));

// === Раздел 5: План «Б» ===
children.push(h1("5. План \u00ABБ\u00BB \u2014 управление рисками"));
children.push(body("Эмпилс не оставляет партнёра на полпути. Если через 6 месяцев уход с полки ниже заявленного, автоматически включается план \u00ABБ\u00BB. Это официальная часть программы, а не дополнительная услуга."));

children.push(h2("Когда включается план \u00ABБ\u00BB"));
children.push(buildTable(
  ["Уровень", "Условие", "Действия"],
  [
    ["\uD83D\uDD34 КРАСНЫЙ", "Уход с полки < 50% от плана", "Включение всех 5 инструментов + личный куратор"],
    ["\uD83D\uDFE1 ЖЁЛТЫЙ", "Уход с полки 50-70% от плана", "Обучение + корректировка ассортимента + промо"],
    ["\uD83D\uDFE2 ЗЕЛЁНЫЙ", "Уход с полки 70-85% от плана", "Дополнительные промо-акции"],
  ],
  [15, 40, 45]
));

children.push(h2("5 инструментов плана \u00ABБ\u00BB"));
children.push(bulletBold("Дополнительное обучение", "Очный выезд тренера Эмпилс к ТП дистрибьютора. Анализ ошибок, отработка возражений \u00ABпо цене\u00BB, разбор реальных кейсов."));
children.push(bulletBold("Корректировка ассортимента", "Анализ продаж по SKU за 6 мес. Замена медленных цветов на ходовые. Подбор оптимальной матрицы под конкретный регион."));
children.push(bulletBold("Дополнительные промо-акции", "Мотивация РТТ за банку. Совместное промо 50/50 с Эмпилс."));
children.push(bulletBold("Мастер-классы в топовых точках", "Выезд РМ в 5-7 самых крупных РТТ. Демонстрация: TITAN vs ПФ-115 (выкрас, адгезия, скорость). Обучение продавцов, 45 мин."));
children.push(bulletBold("Личный куратор от Эмпилс", "Закрепление старшего менеджера. Еженедельные созвоны. Личный визит 1 раз в месяц до выхода на KPI."));

children.push(divider());

// === Раздел 6: Обязательства сторон ===
children.push(h1("6. Обязательства сторон"));

children.push(h2("Эмпилс обязуется:"));
children.push(bullet("Поставлять продукцию TITAN в согласованные сроки и в согласованном объёме"));
children.push(bullet("Предоставлять POS-материалы, каталоги, выкрасы \u2014 бесплатно"));
children.push(bullet("Проводить обучение торговых представителей (выездное и онлайн)"));
children.push(bullet("Обеспечивать цифровое продвижение TITAN (Я.Директ, RuTube, блогер)"));
children.push(bullet("Предоставить систему отчётности CIS-Link"));
children.push(bullet("Включить план \u00ABБ\u00BB при отклонении от KPI"));
children.push(bullet("Назначить персонального регионального менеджера"));

children.push(h2("Дистрибьютор обязуется:"));
children.push(bullet("Соблюдать согласованный MML (не менее ___ SKU TITAN)"));
children.push(bullet("Обеспечить выкладку по планограмме в точках класса A и B"));
children.push(bullet("Обучить 100% торговых представителей по программе TITAN"));
children.push(bullet("Обеспечить наличие POS-материалов в 80% точек с TITAN"));
children.push(bullet("Предоставлять ежемесячную отчётность через CIS-Link"));
children.push(bullet("Не продавать TITAN за пределами согласованной географии"));
children.push(bullet("Соблюдать рекомендованную розничную цену (допустимое отклонение \u00B1__%)"));

children.push(divider());

// === Финал ===
children.push(new Paragraph({
  spacing: { before: 400, after: 100 },
  alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "\u00ABКрасишь 1 раз \u2014 служит 10 лет\u00BB", size: 24, bold: true, color: c(P.accent), font: FONT })],
}));

children.push(new Paragraph({
  spacing: { after: 200 },
  alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "Партнёрство, которое окупается за 4 месяца", size: 18, italics: true, color: c(P.secondary), font: FONT })],
}));

children.push(new Paragraph({
  spacing: { after: 100 },
  alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "ЗАО \u00ABЭмпилс\u00BB \u00B7 г. Ростов-на-Дону \u00B7 2026", size: 16, color: c(P.secondary), font: FONT })],
}));

// === Документ ===
const doc = new Document({
  creator: "ЗАО Эмпилс",
  title: "Франшиза TITAN — Программа работы с дистрибьюторами",
  styles: {
    default: {
      document: {
        run: { font: FONT, size: 22, color: c(P.body) },
        paragraph: { spacing: { line: 280 } },
      },
      heading1: {
        run: { font: FONT, size: 28, bold: true, color: c(P.primary) },
      },
      heading2: {
        run: { font: FONT, size: 22, bold: true, color: c(P.primary) },
      },
    },
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1200, bottom: 1200, left: 1417, right: 1134 },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: "Эмпилс \u00B7 TITAN \u00B7 Франшиза \u2014 Дистрибьюторам", size: 16, color: "808080", font: FONT })],
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ children: [PageNumber.CURRENT], size: 16, color: "808080", font: FONT })],
        })],
      }),
    },
    children,
  }],
});

const OUT = "/home/z/my-project/download/TITAN_Франшиза_Дистрибьюторам.docx";

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log("\u2705 Документ 1 сохранён:", OUT);
  console.log("   Размер:", buf.length.toLocaleString(), "байт");
}).catch(err => {
  console.error("\u274C", err);
  process.exit(1);
});
