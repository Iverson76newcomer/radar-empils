// Script to batch-create all TCO survey questions via Yandex Forms internal API
// Run in browser console on the form edit page

const FORM_ID = '6a4363f4d04688b527b81707';
const PAGE_ID = 9008982056195760; // current page ID

// Helper to generate unique IDs (Yandex-style: 16-digit numbers)
let counter = Math.floor(Date.now() / 1000);
function genId() {
  counter++;
  return parseInt(String(counter) + String(Math.floor(Math.random() * 9000 + 1000)));
}

function getCsrf() {
  const meta = document.querySelector('meta[name=csrf-token]');
  return meta ? meta.content : '';
}

// Question definitions - all 14 remaining questions
const questions = [
  // Q4: TCO familiarity - single choice
  {
    text: '4. Знакомы ли вы с методикой ТСО (Совокупная стоимость владения) в контексте продукции ЛКМ? (ТСО — это учёт не только цены банки, а всех затрат на покраску: материалы, грунт, время, труд, повторные покрытия)',
    type: 'choices', view: 'radio', required: true,
    options: [
      'Да, хорошо понимаю и использую при продажах',
      'Слышал(а) термин, но не применял(а) на практике',
      'Не знаком(а) с термином, но интуитивно считаю полную стоимость',
      'Впервые слышу — всегда ориентируюсь только на цену банки'
    ]
  },
  // Q5: TCO components - multi choice (checkboxes)
  {
    text: '5. Какие составляющие, на ваш взгляд, входят в совокупную стоимость покраски? (отметьте все, что считаете релевантными)',
    type: 'choices', view: 'checks', required: false,
    options: [
      'Цена банки краски',
      'Стоимость грунтовки / подготовки поверхности',
      'Количество слоёв (расход на каждый слой)',
      'Время высыхания (простой между слоями)',
      'Стоимость рабочей силы (своей или наемной)',
      'Срок службы покрытия (частота перекраски)',
      'Расход на инструмент (кисти, валики, растворители)',
      'Затраты на утилизацию / экологию'
    ]
  },
  // Q6: Practice example - long text
  {
    text: '6. Приведите пример из вашей практики, когда учёт полной стоимости помог продать более дорогий продукт (если такого примера нет — напишите «нет»)',
    type: 'text', view: 'textarea', required: false
  },
  // Q7: Recommendation method - single
  {
    text: '7. Как вы обычно рекомендуете ЛКМ покупателю?',
    type: 'choices', view: 'radio', required: true,
    options: [
      'Называю цену и основные характеристики',
      'Сравниваю несколько вариантов по цене и качеству',
      'Считаю вместе с покупателем полную стоимость (грунт + краска + работа)',
      'Предлагаю несколько вариантов и оставляю выбор покупателю',
      'Другое: _______________'
    ]
  },
  // Q8: Best argument for premium - single
  {
    text: '8. Какой аргумент для покупателя считаете наиболее эффективным при продаже продукта выше среднего ценового сегмента?',
    type: 'choices', view: 'radio', required: true,
    options: [
      '«Это качественнее / долговечнее»',
      '«Экономит время — меньше слоёв, быстрее работа»',
      '«Полная стоимость покраски в итоге дешевле»',
      '«Этим пользуются профессионалы»',
      '«Не придётся перекрашивать через год»',
      '«Укрывистость лучше — меньше расход»',
      'Другое: _______________'
    ]
  },
  // Q9: Customer phrases - multi
  {
    text: '9. Какие вопросы покупателя чаще всего демонстрируют, что он выбирает по цене, а не по выгоде? (выберите до 3-х)',
    type: 'choices', view: 'checks', required: false,
    options: [
      '«А чем это дороже обычной краски?»',
      '«А нет чего подешевле?»',
      '«Мне нужно покрасить недорого»',
      '«Я покрашу сам — не хочу платить за дорогую краску»',
      '«Давайте то, что подешевле — всё равно одинаково»',
      'Другое: _______________'
    ]
  },
  // Q10: TCO influence scale - scale 0-10
  {
    text: '10. Считаете ли вы, что расчёт ТСО может повлиять на решение покупателя выбрать более дорогой, но выгодный продукт? (0 — точно не повлияет, 10 — очень сильно повлияет)',
    type: 'scale', required: true
  },
  // Q11: Target audience - single
  {
    text: '11. Какой тип покупателя, на ваш взгляд, наиболее восприимчив к аргументу о совокупной стоимости?',
    type: 'choices', view: 'radio', required: true,
    options: [
      'Дачник / хозяин дома — красит сам, ценит своё время',
      'Профессиональный маляр / строитель — считает рабочую силу',
      'Покупатель эконом-сегмента — всегда выбирает по цене',
      'Корпоративный заказчик — считает бюджет проекта',
      'Все одинаково — неважно, кто покупает'
    ]
  },
  // Q12: Best products for TCO - multi
  {
    text: '12. Какие продукты ЛКМ, на ваш взгляд, лучше всего продаются через аргумент ТСО? (выберите до 3-х)',
    type: 'choices', view: 'checks', required: false,
    options: [
      'Грунт-эмали 3в1 (замена грунт + эмаль)',
      'Премиум-эмали (замена стандартных ПФ-115)',
      'Водно-дисперсионные краски повышенной стойкости',
      'Фасадные краски с повышенной укрывистостью',
      'Лаки / пропитки премиум-сегмента',
      'Другое: _______________'
    ]
  },
  // Q13: Barriers - multi
  {
    text: '13. Что мешает вам сейчас продавать через аргумент ТСО? (выберите до 3-х)',
    type: 'choices', view: 'checks', required: false,
    options: [
      'Нет наглядного расчёта / калькулятора',
      'Не хватает знаний для расчёта ТСО',
      'Покупателю не интересны расчёты — хочет цену банки',
      'Нет времени на долгие объяснения',
      'Не верю, что покупатель захочет платить больше',
      'Конкуренты продают дешевле, и покупатель уходит',
      'Другое: _______________'
    ]
  },
  // Q14: Needed tools - multi
  {
    text: '14. Какие инструменты помогли бы вам начать продавать через ТСО? (выберите до 3-х)',
    type: 'choices', view: 'checks', required: false,
    options: [
      'Калькулятор ТСО (онлайн / в телефоне)',
      'Печатная шпаргалка с готовыми сравнениями',
      'Образцы покрытий (до / после, через год)',
      'Обучение: как за 2 минуты объяснить ТСО',
      'Видео / презентация для покупателя',
      'Сравнительная таблица ТСО: наш продукт vs конкуренты',
      'Мотивационная программа за продажу через ТСО',
      'Другое: _______________'
    ]
  },
  // Q15: Implementation timing - single
  {
    text: '15. Если бы производитель предоставил вам готовый инструмент (калькулятор + шпаргалку + обучение) — как скоро вы бы его внедрили?',
    type: 'choices', view: 'radio', required: true,
    options: [
      'Сразу же — это то, чего нам не хватало',
      'В течение 1–2 месяцев после обучения',
      'Через полгода — нужно время на привыкание',
      'Скорее не внедрил бы — не уверен в эффективности'
    ]
  },
  // Q16: Expectations - multi
  {
    text: '16. Что вы ожидаете от производителя ЛКМ в части поддержки продаж через ТСО? (выберите до 3-х)',
    type: 'choices', view: 'checks', required: false,
    options: [
      'Готовые материалы для продажи (листовки, таблицы, калькулятор)',
      'Обучение персонала методике ТСО',
      'Политика ценообразования, подтверждающая выгоду ТСО',
      'Маркетинговая поддержка (промо, реклама, дегустации)',
      'Сопровождение на месте (мерчандайзинг, тренинги)',
      'Другое: _______________'
    ]
  },
  // Q17: TCO prospects - scale 0-10
  {
    text: '17. Оцените перспективность методики ТСО для вашего бизнеса (0 — бесполезна, 10 — очень перспективно)',
    type: 'scale', required: true
  }
];

function buildQuestionBody(q, position) {
  const id = genId();
  const key = `answer_${q.type}_${id}`;
  const base = {
    id: id,
    key: key,
    type: q.type,
    title: q.text,
    hidden: false,
    externalValidator: false
  };

  if (q.type === 'choices') {
    base.view = q.view;
    base.widget = q.view;
    base.options = q.options.map((text, idx) => ({
      id: genId(),
      text: text,
      key: String(genId()),
      hidden: false
    }));
    base.sort = 'natural';
    if (q.required) base.required = true;
  } else if (q.type === 'text') {
    base.view = q.view || 'textarea';
    base.widget = q.view || 'textarea';
    if (q.required) base.required = true;
  } else if (q.type === 'scale') {
    base.type = 'choices';
    base.view = 'scale';
    base.widget = 'scale';
    base.range = { min: 0, max: 10, step: 1 };
    if (q.required) base.required = true;
  }

  return {
    surveyId: FORM_ID,
    page: PAGE_ID,
    position: position,
    question: base
  };
}

async function addAllQuestions() {
  let position = 4; // start after the 4 existing questions (intro + Q1-Q3)
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
      results.push({ idx: i + 4, success: resp.ok, id: data.id, error: data.message });
      console.log(`Q${i + 4}: ${resp.ok ? 'OK' : 'FAIL'} - ${data.id || data.message}`);
      position++;
      // Small delay to avoid rate limit
      await new Promise(r => setTimeout(r, 500));
    } catch (e) {
      results.push({ idx: i + 4, success: false, error: String(e) });
      console.error(`Q${i + 4}: ERROR - ${e}`);
    }
  }

  return results;
}

// Run it and return results
addAllQuestions().then(r => window.__tsoResults = r);
