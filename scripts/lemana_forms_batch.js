// Lemana PRO survey - all 17 questions
// Run in browser console on the form edit page

const FORM_ID = '6a4377dc49af47002b451dbe';
const PAGE_ID = 9322063;

let counter = Math.floor(Date.now() / 1000);
function genId() {
  counter++;
  return parseInt(String(counter) + String(Math.floor(Math.random() * 9000 + 1000)));
}

function getCsrf() {
  return document.querySelector('meta[name=csrf-token]').content;
}

const questions = [
  // Block 1: Role & Experience
  {
    text: '1. Ваша должность в «Лемана ПРО»',
    type: 'choices', view: 'radio', required: true,
    options: [
      'Руководитель отдела (закупки / категорийный менеджмент)',
      'Менеджер по работе с поставщиками',
      'Консультант / продавец торгового зала',
      'Мерчандайзер / супервайзер',
      'Другое: ________________'
    ]
  },
  {
    text: '2. Ваш опыт работы в категории ЛКМ',
    type: 'choices', view: 'radio', required: true,
    options: ['Менее 1 года', '1–3 года', '3–5 лет', 'Более 5 лет']
  },
  {
    text: '3. В каком формате вы чаще всего общаетесь с покупателями ЛКМ?',
    type: 'choices', view: 'radio', required: true,
    options: [
      'Ежедневно помогаю выбрать краску в зале',
      'Периодически консультирую по запросу',
      'Редко сталкиваюсь с ЛКМ — другая категория',
      'Не работаю с покупателями напрямую'
    ]
  },
  // Block 2: Brand awareness
  {
    text: '4. Какие торговые марки компании «Эмпилс» вам известны? (перечислите все, которые припомните)',
    type: 'text', view: 'textarea', required: false
  },
  {
    text: '5. Какая ассоциация возникает у вас в первую очередь при слове «Эмпилс»? (выберите одну)',
    type: 'choices', view: 'radio', required: true,
    options: [
      'Надёжность, проверенное качество',
      'Советское наследие / старый завод',
      'Устаревшие продукты',
      'Современные технологии и инновации',
      'Доступная цена / эконом-сегмент',
      'Широкий ассортимент',
      'Не могу сказать — мало знаком(а)'
    ]
  },
  // Block 3: Product perception
  {
    text: '6. Как вы оцениваете продукцию «Эмпилс» по сравнению с другими брендами в вашей категории ЛКМ? (по 5-балльной шкале, где 1 — значительно хуже, 5 — значительно лучше)',
    type: 'scale5', required: true
  },
  {
    text: '7. Как вы считаете, какого сегмента в основном продукция «Эмпилс»?',
    type: 'choices', view: 'radio', required: true,
    options: [
      'Только эконом',
      'В основном эконом, есть средний',
      'Эконом + средний + премиум — широкий охват',
      'В основном средний сегмент',
      'Затрудняюсь оценить'
    ]
  },
  // Block 4: TITAN 3в1
  {
    text: '8. Знаете ли вы о грунт-эмали по ржавчине TITAN 3в1 в ассортименте «Лемана ПРО»?',
    type: 'choices', view: 'radio', required: true,
    options: [
      'Да, хорошо знаю и могу объяснить покупателю',
      'Слышал(а), но подробностей не знаю',
      'Вижу на полке, но никогда не рекомендовал(а)',
      'Нет, впервые слышу'
    ]
  },
  {
    text: '9. Если бы покупатель спросил: «Какую краску взять для ржавых ворот на даче?» — что бы вы порекомендовали?',
    type: 'choices', view: 'radio', required: true,
    options: [
      'Грунт-эмаль 3в1 (TITAN или аналог)',
      'Сначала грунт ГФ-021, потом эмаль ПФ-115',
      'СТМ сети (лучшая цена)',
      'Уточнил(а) бы бюджет и тогда порекомендовал(а)',
      'Другое: ________________'
    ]
  },
  {
    text: '10. Что, на ваш взгляд, может стать главным препятствием для продаж TITAN 3в1 в «Лемана ПРО»? (выберите 1–2 варианта)',
    type: 'choices', view: 'checks', required: false,
    options: [
      'Цена выше, чем у обычной эмали ПФ-115',
      'Покупатель не знает бренд «Титан»',
      'Недоверие к формату «3в1» — сомнения, что грунт + краска в одной банке работают',
      'Конкуренция с СТМ «Лемана ПРО»',
      'Сложность объяснить выгоду покупателю',
      'Недостаточная рекламная поддержка на полке',
      'Другое: ________________'
    ]
  },
  // Block 5: TCO methodology
  {
    text: '11. Знакомы ли вы с методикой ТСО (Совокупная стоимость владения) в контексте ЛКМ? (ТСО — это учёт не только цены банки, а всех затрат на покраску: материалы, грунт, время, труд, повторные покрытия)',
    type: 'choices', view: 'radio', required: true,
    options: [
      'Да, хорошо понимаю и использую в работе с покупателями',
      'Слышал(а) термин, но не применял(а) на практике',
      'Не знаком(а) с термином, но интуитивно считаю полную стоимость при консультации',
      'Впервые слышу — всегда ориентируюсь только на цену банки'
    ]
  },
  {
    text: '12. Как вы сейчас объясняете покупателю разницу между дешёвой эмалью ПФ-115 и грунт-эмалью 3в1 по более высокой цене?',
    type: 'choices', view: 'radio', required: true,
    options: [
      'Говорю, что 3в1 — это качественнее и долговечнее',
      'Считаю вместе с покупателем полную стоимость (грунт + краска + работа)',
      'Рекомендую то, что дешевле — покупатель сам решит',
      'Обычно не сравниваю — покупатель сам выбирает по цене',
      'Другое: ________________'
    ]
  },
  {
    text: '13. Представьте: покупатель выбирает между двумя вариантами покраски гаража. Какой аргумент считаете наиболее убедительным для продажи 3в1?',
    type: 'choices', view: 'radio', required: true,
    options: [
      '«3в1 экономит время — не нужно отдельно грунтовать, один проход и готово»',
      '«Полная стоимость покраски 3в1 дешевле: одна банка вместо двух + экономия времени»',
      '«3в1 держится дольше, не придётся перекрашивать через год»',
      '«Это профессиональный продукт, таким пользуются мастера»',
      '«Смотрите, у 3в1 больше укрывистость — меньше слоёв нужно»',
      'Затрудняюсь — нужен конкретный расчёт'
    ]
  },
  {
    text: '14. Считаете ли вы, что расчёт ТСО (совокупной стоимости) может повлиять на решение покупателя выбрать более дорогой, но выгодный в итоге продукт? (по шкале от 0 до 10, где 0 — точно не повлияет, 10 — очень сильно повлияет)',
    type: 'scale10', required: true
  },
  {
    text: '15. Если бы у вас был наглядный инструмент (калькулятор / таблица / шпаргалка), показывающий полную стоимость покраски — стали бы вы им пользоваться при консультации покупателя?',
    type: 'choices', view: 'radio', required: true,
    options: [
      'Да, обязательно — это сильный аргумент',
      'Скорее да, если инструмент простой и понятный',
      'Скорее нет — покупателю достаточно цены банки',
      'Нет — покупатель не станет слушать про расчёты',
      'Зависит от покупателя — кому-то покажу, кому-то нет'
    ]
  },
  // Block 6: Support needs
  {
    text: '16. Что помогло бы вам увереннее продавать TITAN 3в1? (выберите до 3-х)',
    type: 'choices', view: 'checks', required: false,
    options: [
      'Краткая шпаргалка на полке (какие вопросы задать покупателю)',
      'Примеры образцов покрытия (до/после)',
      'Обучение: как за 2 минуты объяснить выгоду формата 3в1',
      'Калькулятор ТСО (полная стоимость покраски с учётом грунта, времени, слоёв)',
      'Промо-материалы для покупателя (листовки, видео на экране)',
      'Сравнительная таблица TITAN 3в1 vs конкуренты',
      'Мотивационная программа для продавцов',
      'Другое: ________________'
    ]
  },
  {
    text: '17. Использовали ли вы лично продукцию «Эмпилс» (любую)?',
    type: 'choices', view: 'radio', required: true,
    options: [
      'Да, регулярно',
      'Да, иногда',
      'Пробовал(а) один раз',
      'Нет, не доводилось'
    ]
  }
];

function buildQuestionBody(q, position) {
  const id = genId();

  if (q.type === 'choices') {
    return {
      surveyId: FORM_ID,
      page: PAGE_ID,
      position: position,
      question: {
        id: id,
        key: `answer_choices_${id}`,
        type: 'choices',
        title: q.text,
        view: q.view,
        widget: q.view === 'checks' ? 'checkbox' : 'radio',
        options: q.options.map(text => ({
          id: genId(),
          text: text,
          key: String(genId()),
          hidden: false
        })),
        sort: 'natural',
        hidden: false,
        externalValidator: false,
        required: !!q.required
      }
    };
  } else if (q.type === 'text') {
    return {
      surveyId: FORM_ID,
      page: PAGE_ID,
      position: position,
      question: {
        id: id,
        key: `answer_text_${id}`,
        type: 'text',
        title: q.text,
        view: q.view || 'textarea',
        widget: q.view || 'textarea',
        hidden: false,
        externalValidator: false,
        required: !!q.required
      }
    };
  } else if (q.type === 'scale5' || q.type === 'scale10') {
    const max = q.type === 'scale5' ? 5 : 10;
    const columns = [];
    for (let i = 1; i <= max; i++) {
      columns.push({id: genId(), label: String(i), slug: `matrix_column_${genId()}`});
    }
    const rowId = genId();
    return {
      surveyId: FORM_ID,
      page: PAGE_ID,
      position: position,
      question: {
        id: id,
        key: `answer_matrix_${id}`,
        type: 'matrix',
        title: q.text,
        rows: [{id: rowId, label: 'Оценка', slug: `matrix_row_${rowId}`}],
        columns: columns,
        required: !!q.required,
        hidden: false,
        externalValidator: false
      }
    };
  }
}

async function addAllQuestions() {
  let position = 1; // start after intro
  const results = [];

  for (let i = 0; i < questions.length; i++) {
    const q = questions[i];
    const body = buildQuestionBody(q, position);

    try {
      const resp = await fetch('/admin/gateway/root/form/addSurveyQuestion', {
        method: 'POST',
        credentials: 'include',
        headers: {
          'Content-Type': 'application/json',
          'X-CSRF-Token': getCsrf()
        },
        body: JSON.stringify(body)
      });
      const data = await resp.json();
      results.push({q: i + 1, success: resp.ok, id: data.id, error: data.message});
      console.log(`Q${i + 1}: ${resp.ok ? 'OK' : 'FAIL'} - ${data.id || data.message}`);
      position++;
      await new Promise(r => setTimeout(r, 400));
    } catch (e) {
      results.push({q: i + 1, success: false, error: String(e)});
      console.error(`Q${i + 1}: ERROR - ${e}`);
    }
  }

  return results;
}

addAllQuestions().then(r => window.__lemanaResults = r);
