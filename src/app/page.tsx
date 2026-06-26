'use client'

import { useState, useEffect, useCallback } from 'react'
import { motion, AnimatePresence } from 'framer-motion'
import {
  Card, CardContent, CardDescription, CardHeader, CardTitle,
} from '@/components/ui/card'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Textarea } from '@/components/ui/textarea'
import { Label } from '@/components/ui/label'
import { RadioGroup, RadioGroupItem } from '@/components/ui/radio-group'
import { Checkbox } from '@/components/ui/checkbox'
import { Progress } from '@/components/ui/progress'

import { Badge } from '@/components/ui/badge'
import { Separator } from '@/components/ui/separator'
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip as RTooltip,
  ResponsiveContainer, PieChart, Pie, Cell, Legend,
} from 'recharts'
import {
  ClipboardList, BarChart3, ChevronRight, ChevronLeft, Send,
  CheckCircle2, Clock, Users, TrendingUp,
} from 'lucide-react'
import { useToast } from '@/hooks/use-toast'

/* ------------------------------------------------------------------ */
/*  QUESTION DEFINITIONS                                                */
/* ------------------------------------------------------------------ */

type QuestionType = 'single' | 'multi' | 'scale' | 'text'

interface Option {
  value: string
  label: string
  hasOther?: boolean
}

interface Question {
  id: string
  block: number
  blockTitle: string
  text: string
  type: QuestionType
  options?: Option[]
  maxSelect?: number
  required?: boolean
}

const BLOCKS = [
  { num: 1, title: 'Ваша роль и опыт', icon: '👤' },
  { num: 2, title: 'Понимание методики ТСО', icon: '💡' },
  { num: 3, title: 'Текущий подход к продаже ЛКМ', icon: '🛒' },
  { num: 4, title: 'ТСО как инструмент продажи', icon: '🎯' },
  { num: 5, title: 'Барьеры и потребности', icon: '🚧' },
  { num: 6, title: 'Ожидания от производителя', icon: '🤝' },
]

const QUESTIONS: Question[] = [
  // Block 1
  {
    id: 'q1', block: 1, blockTitle: 'Ваша роль и опыт',
    text: 'Ваша должность', type: 'single',
    options: [
      { value: 'head', label: 'Руководитель / владелец компании' },
      { value: 'purchase_head', label: 'Руководитель отдела закупок / категорийный менеджер' },
      { value: 'sales_mgr', label: 'Менеджер по продажам / торговый представитель' },
      { value: 'consultant', label: 'Продавец / консультант в зале' },
      { value: 'other', label: 'Другое', hasOther: true },
    ],
    required: true,
  },
  {
    id: 'q2', block: 1, blockTitle: 'Ваша роль и опыт',
    text: 'Ваш опыт работы в сфере ЛКМ', type: 'single',
    options: [
      { value: 'lt1', label: 'Менее 1 года' },
      { value: '1-3', label: '1–3 года' },
      { value: '3-5', label: '3–5 лет' },
      { value: 'gt5', label: 'Более 5 лет' },
    ],
    required: true,
  },
  {
    id: 'q3', block: 1, blockTitle: 'Ваша роль и опыт',
    text: 'Какой формат продаж ЛКМ доминирует в вашей компании?', type: 'single',
    options: [
      { value: 'retail', label: 'Преимущественно розничные продажи (сеть магазинов / DIY)' },
      { value: 'wholesale', label: 'Преимущественно оптовые продажи (дистрибуция)' },
      { value: 'mixed', label: 'Смешанный формат (опт + розница)' },
      { value: 'project', label: 'Продажи проектным / корпоративным клиентам' },
    ],
    required: true,
  },
  // Block 2
  {
    id: 'q4', block: 2, blockTitle: 'Понимание методики ТСО',
    text: 'Знакомы ли вы с методикой ТСО (Совокупная стоимость владения) в контексте продукции ЛКМ? (ТСО — учёт не только цены банки, а всех затрат на покраску: материалы, грунт, время, труд, повторные покрытия)',
    type: 'single',
    options: [
      { value: 'know_use', label: 'Да, хорошо понимаю и использую при продажах' },
      { value: 'heard', label: 'Слышал(а) термин, но не применял(а) на практике' },
      { value: 'intuitive', label: 'Не знаком(а) с термином, но интуитивно считаю полную стоимость' },
      { value: 'first_time', label: 'Впервые слышу — всегда ориентируюсь только на цену банки' },
    ],
    required: true,
  },
  {
    id: 'q5', block: 2, blockTitle: 'Понимание методики ТСО',
    text: 'Какие составляющие, на ваш взгляд, входят в совокупную стоимость покраски? (отметьте все, что считаете релевантными)',
    type: 'multi', maxSelect: 8,
    options: [
      { value: 'paint_price', label: 'Цена банки краски' },
      { value: 'primer', label: 'Стоимость грунтовки / подготовки поверхности' },
      { value: 'layers', label: 'Количество слоёв (расход на каждый слой)' },
      { value: 'drying', label: 'Время высыхания (простой между слоями)' },
      { value: 'labor', label: 'Стоимость рабочей силы (своей или наемной)' },
      { value: 'lifespan', label: 'Срок службы покрытия (частота перекраски)' },
      { value: 'tools', label: 'Расход на инструмент (кисти, валики, растворители)' },
      { value: 'disposal', label: 'Затраты на утилизацию / экологию' },
    ],
  },
  {
    id: 'q6', block: 2, blockTitle: 'Понимание методики ТСО',
    text: 'Приведите пример из вашей практики, когда учёт полной стоимости помог продать более дорогий продукт (если такого примера нет — напишите «нет»)',
    type: 'text',
  },
  // Block 3
  {
    id: 'q7', block: 3, blockTitle: 'Текущий подход к продаже ЛКМ',
    text: 'Как вы обычно рекомендуете ЛКМ покупателю?', type: 'single',
    options: [
      { value: 'price_chars', label: 'Называю цену и основные характеристики' },
      { value: 'compare', label: 'Сравниваю несколько вариантов по цене и качеству' },
      { value: 'full_cost', label: 'Считаю вместе с покупателем полную стоимость (грунт + краска + работа)' },
      { value: 'options', label: 'Предлагаю несколько вариантов и оставляю выбор покупателю' },
      { value: 'other', label: 'Другое', hasOther: true },
    ],
    required: true,
  },
  {
    id: 'q8', block: 3, blockTitle: 'Текущий подход к продаже ЛКМ',
    text: 'Какой аргумент для покупателя считаете наиболее эффективным при продаже продукта выше среднего ценового сегмента?', type: 'single',
    options: [
      { value: 'quality', label: '«Это качественнее / долговечнее»' },
      { value: 'time', label: '«Экономит время — меньше слоёв, быстрее работа»' },
      { value: 'tco', label: '«Полная стоимость покраски в итоге дешевле»' },
      { value: 'pros', label: '«Этим пользуются профессионалы»' },
      { value: 'no_repaint', label: '«Не придётся перекрашивать через год»' },
      { value: 'coverage', label: '«Укрывистость лучше — меньше расход»' },
      { value: 'other', label: 'Другое', hasOther: true },
    ],
    required: true,
  },
  {
    id: 'q9', block: 3, blockTitle: 'Текущий подход к продаже ЛКМ',
    text: 'Какие вопросы покупателя чаще всего демонстрируют, что он выбирает по цене, а не по выгоде? (выберите до 3-х)',
    type: 'multi', maxSelect: 3,
    options: [
      { value: 'why_expensive', label: '«А чем это дороже обычной краски?»' },
      { value: 'cheaper', label: '«А нет чего подешевле?»' },
      { value: 'inexpensive', label: '«Мне нужно покрасить недорого»' },
      { value: 'diy', label: '«Я покрашу сам — не хочу платить за дорогую краску»' },
      { value: 'same', label: '«Давайте то, что подешевле — всё равно одинаково»' },
      { value: 'other', label: 'Другое', hasOther: true },
    ],
  },
  // Block 4
  {
    id: 'q10', block: 4, blockTitle: 'ТСО как инструмент продажи',
    text: 'Считаете ли вы, что расчёт ТСО может повлиять на решение покупателя выбрать более дорогой, но выгодный продукт? (0 — точно не повлияет, 10 — очень сильно повлияет)',
    type: 'scale',
    required: true,
  },
  {
    id: 'q11', block: 4, blockTitle: 'ТСО как инструмент продажи',
    text: 'Какой тип покупателя, на ваш взгляд, наиболее восприимчив к аргументу о совокупной стоимости?', type: 'single',
    options: [
      { value: 'dacha', label: 'Дачник / хозяин дома — красит сам, ценит своё время' },
      { value: 'pro', label: 'Профессиональный маляр / строитель — считает рабочую силу' },
      { value: 'economy', label: 'Покупатель эконом-сегмента — всегда выбирает по цене' },
      { value: 'corp', label: 'Корпоративный заказчик — считает бюджет проекта' },
      { value: 'all_same', label: 'Все одинаково — неважно, кто покупает' },
    ],
    required: true,
  },
  {
    id: 'q12', block: 4, blockTitle: 'ТСО как инструмент продажи',
    text: 'Какие продукты ЛКМ, на ваш взгляд, лучше всего продаются через аргумент ТСО? (выберите до 3-х)',
    type: 'multi', maxSelect: 3,
    options: [
      { value: '3in1', label: 'Грунт-эмали 3в1 (замена грунт + эмаль)' },
      { value: 'premium_enamel', label: 'Премиум-эмали (замена стандартных ПФ-115)' },
      { value: 'water_dispersion', label: 'Водно-дисперсионные краски повышенной стойкости' },
      { value: 'facade', label: 'Фасадные краски с повышенной укрывистостью' },
      { value: 'lacquer', label: 'Лаки / пропитки премиум-сегмента' },
      { value: 'other', label: 'Другое', hasOther: true },
    ],
  },
  // Block 5
  {
    id: 'q13', block: 5, blockTitle: 'Барьеры и потребности',
    text: 'Что мешает вам сейчас продавать через аргумент ТСО? (выберите до 3-х)',
    type: 'multi', maxSelect: 3,
    options: [
      { value: 'no_calc', label: 'Нет наглядного расчёта / калькулятора' },
      { value: 'no_knowledge', label: 'Не хватает знаний для расчёта ТСО' },
      { value: 'no_interest', label: 'Покупателю не интересны расчёты — хочет цену банки' },
      { value: 'no_time', label: 'Нет времени на долгие объяснения' },
      { value: 'no_belief', label: 'Не верю, что покупатель захочет платить больше' },
      { value: 'competitors', label: 'Конкуренты продают дешевле, и покупатель уходит' },
      { value: 'other', label: 'Другое', hasOther: true },
    ],
  },
  {
    id: 'q14', block: 5, blockTitle: 'Барьеры и потребности',
    text: 'Какие инструменты помогли бы вам начать продавать через ТСО? (выберите до 3-х)',
    type: 'multi', maxSelect: 3,
    options: [
      { value: 'calculator', label: 'Калькулятор ТСО (онлайн / в телефоне)' },
      { value: 'cheat_sheet', label: 'Печатная шпаргалка с готовыми сравнениями' },
      { value: 'samples', label: 'Образцы покрытий (до / после, через год)' },
      { value: 'training', label: 'Обучение: как за 2 минуты объяснить ТСО' },
      { value: 'video', label: 'Видео / презентация для покупателя' },
      { value: 'comparison', label: 'Сравнительная таблица ТСО: наш продукт vs конкуренты' },
      { value: 'motivation', label: 'Мотивационная программа за продажу через ТСО' },
      { value: 'other', label: 'Другое', hasOther: true },
    ],
  },
  {
    id: 'q15', block: 5, blockTitle: 'Барьеры и потребности',
    text: 'Если бы производитель предоставил вам готовый инструмент (калькулятор + шпаргалку + обучение) — как скоро вы бы его внедрили?', type: 'single',
    options: [
      { value: 'immediately', label: 'Сразу же — это то, чего нам не хватало' },
      { value: '1-2_months', label: 'В течение 1–2 месяцев после обучения' },
      { value: 'half_year', label: 'Через полгода — нужно время на привыкание' },
      { value: 'unlikely', label: 'Скорее не внедрил бы — не уверен в эффективности' },
    ],
    required: true,
  },
  // Block 6
  {
    id: 'q16', block: 6, blockTitle: 'Ожидания от производителя',
    text: 'Что вы ожидаете от производителя ЛКМ в части поддержки продаж через ТСО? (выберите до 3-х)',
    type: 'multi', maxSelect: 3,
    options: [
      { value: 'materials', label: 'Готовые материалы для продажи (листовки, таблицы, калькулятор)' },
      { value: 'education', label: 'Обучение персонала методике ТСО' },
      { value: 'pricing', label: 'Политика ценообразования, подтверждающая выгоду ТСО' },
      { value: 'marketing', label: 'Маркетинговая поддержка (промо, реклама, дегустации)' },
      { value: 'on_site', label: 'Сопровождение на месте (мерчандайзинг, тренинги)' },
      { value: 'other', label: 'Другое', hasOther: true },
    ],
  },
  {
    id: 'q17', block: 6, blockTitle: 'Ожидания от производителя',
    text: 'Оцените перспективность методики ТСО для вашего бизнеса (0 — бесполезна, 10 — очень перспективно)',
    type: 'scale',
    required: true,
  },
]

/* ------------------------------------------------------------------ */
/*  LABEL MAPS                                                          */
/* ------------------------------------------------------------------ */

const LABEL_MAPS: Record<string, Record<string, string>> = {
  q1: { head: 'Руководитель / владелец', purchase_head: 'Руководитель закупок', sales_mgr: 'Менеджер по продажам', consultant: 'Продавец / консультант', other: 'Другое' },
  q2: { lt1: 'Менее 1 года', '1-3': '1–3 года', '3-5': '3–5 лет', gt5: 'Более 5 лет' },
  q3: { retail: 'Розница', wholesale: 'Опт', mixed: 'Смешанный', project: 'Проекты' },
  q4: { know_use: 'Знаю и использую', heard: 'Слышал, не применял', intuitive: 'Интуитивно считаю', first_time: 'Впервые слышу' },
  q5: { paint_price: 'Цена банки', primer: 'Грунтовка', layers: 'Кол-во слоёв', drying: 'Время сушки', labor: 'Рабочая сила', lifespan: 'Срок службы', tools: 'Инструмент', disposal: 'Утилизация' },
  q7: { price_chars: 'Цена + характеристики', compare: 'Сравнение вариантов', full_cost: 'Полная стоимость', options: 'Варианты покупателю', other: 'Другое' },
  q8: { quality: 'Качественнее', time: 'Экономит время', tco: 'Полная стоимость дешевле', pros: 'Профессионалы', no_repaint: 'Без перекраски', coverage: 'Укрывистость', other: 'Другое' },
  q9: { why_expensive: 'Чем дороже?', cheaper: 'Подешевле?', inexpensive: 'Недорого', diy: 'Покрашу сам', same: 'Всё равно одинаково', other: 'Другое' },
  q11: { dacha: 'Дачник', pro: 'Маляр / строитель', economy: 'Эконом-сегмент', corp: 'Корп. заказчик', all_same: 'Все одинаково' },
  q12: { '3in1': 'Грунт-эмали 3в1', premium_enamel: 'Премиум-эмали', water_dispersion: 'ВД повыш. стойкости', facade: 'Фасадные краски', lacquer: 'Лаки / пропитки', other: 'Другое' },
  q13: { no_calc: 'Нет калькулятора', no_knowledge: 'Нет знаний', no_interest: 'Покупателю не интересно', no_time: 'Нет времени', no_belief: 'Не верю в ТСО', competitors: 'Конкуренты дешевле', other: 'Другое' },
  q14: { calculator: 'Калькулятор ТСО', cheat_sheet: 'Шпаргалка', samples: 'Образцы покрытий', training: 'Обучение', video: 'Видео / презентация', comparison: 'Сравнит. таблица', motivation: 'Мотивация', other: 'Другое' },
  q15: { immediately: 'Сразу же', '1-2_months': '1–2 месяца', half_year: 'Через полгода', unlikely: 'Не внедрил бы' },
  q16: { materials: 'Готовые материалы', education: 'Обучение', pricing: 'Ценообразование', marketing: 'Маркетинг', on_site: 'Сопровождение', other: 'Другое' },
}

const COLORS = [
  '#2563eb', '#7c3aed', '#db2777', '#ea580c', '#16a34a',
  '#0891b2', '#4f46e5', '#c026d3', '#dc2626', '#ca8a04',
  '#059669', '#0284c7',
]

/* ------------------------------------------------------------------ */
/*  SURVEY FORM COMPONENT                                               */
/* ------------------------------------------------------------------ */

function SurveyForm({ onSubmit }: { onSubmit: () => void }) {
  const { toast } = useToast()
  const [currentBlock, setCurrentBlock] = useState(0) // 0-indexed
  const [answers, setAnswers] = useState<Record<string, string | string[] | number | null>>({})
  const [otherTexts, setOtherTexts] = useState<Record<string, string>>({})
  const [submitting, setSubmitting] = useState(false)
  const [submitted, setSubmitted] = useState(false)

  const blockQuestions = QUESTIONS.filter((q) => q.block === currentBlock + 1)
  const totalBlocks = BLOCKS.length
  const progress = ((currentBlock + 1) / totalBlocks) * 100

  const handleSingleChange = (qId: string, value: string) => {
    setAnswers((prev) => ({ ...prev, [qId]: value }))
  }

  const handleMultiChange = (qId: string, value: string, checked: boolean, maxSelect?: number) => {
    const current = ((answers[qId] as string[]) || [])
    let updated: string[]
    if (checked) {
      if (maxSelect && current.length >= maxSelect) {
        toast({
          title: `Максимум ${maxSelect} вариант(а)`,
          description: `Можно выбрать не более ${maxSelect} вариантов`,
          variant: 'destructive',
        })
        return
      }
      updated = [...current, value]
    } else {
      updated = current.filter((v) => v !== value)
    }
    setAnswers((prev) => ({ ...prev, [qId]: updated }))
  }

  const handleScaleChange = (qId: string, value: number) => {
    setAnswers((prev) => ({ ...prev, [qId]: value }))
  }

  const handleTextChange = (qId: string, value: string) => {
    setAnswers((prev) => ({ ...prev, [qId]: value }))
  }

  const handleOtherText = (qId: string, text: string) => {
    setOtherTexts((prev) => ({ ...prev, [qId]: text }))
  }

  const canProceed = () => {
    const requiredQuestions = blockQuestions.filter((q) => q.required)
    return requiredQuestions.every((q) => {
      const a = answers[q.id]
      if (q.type === 'scale') return a !== null && a !== undefined
      if (q.type === 'single') return a !== null && a !== undefined && a !== ''
      return true
    })
  }

  const handleSubmit = async () => {
    setSubmitting(true)
    try {
      // Merge other texts into answers
      const finalAnswers = { ...answers }
      for (const [key, text] of Object.entries(otherTexts)) {
        if (text.trim()) {
          const qId = key.replace('_other', '')
          const q = QUESTIONS.find((qq) => qq.id === qId)
          if (q?.type === 'single') {
            finalAnswers[qId] = `Другое: ${text}`
          } else if (q?.type === 'multi') {
            const arr = (finalAnswers[qId] as string[]) || []
            const idx = arr.indexOf('other')
            if (idx >= 0) arr[idx] = `Другое: ${text}`
            finalAnswers[qId] = arr
          }
        }
      }

      const res = await fetch('/api/tso-survey', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ answers: finalAnswers }),
      })

      if (!res.ok) throw new Error('Failed to submit')

      setSubmitted(true)
      onSubmit()
      toast({
        title: 'Спасибо за ваши ответы!',
        description: 'Ваши ответы успешно сохранены.',
      })
    } catch {
      toast({
        title: 'Ошибка отправки',
        description: 'Попробуйте ещё раз.',
        variant: 'destructive',
      })
    } finally {
      setSubmitting(false)
    }
  }

  if (submitted) {
    return (
      <motion.div
        initial={{ opacity: 0, scale: 0.9 }}
        animate={{ opacity: 1, scale: 1 }}
        className="flex flex-col items-center justify-center py-16 text-center"
      >
        <CheckCircle2 className="h-20 w-20 text-green-500 mb-6" />
        <h2 className="text-3xl font-bold mb-3">Спасибо за ваше время!</h2>
        <p className="text-muted-foreground text-lg max-w-md">
          Ваши ответы помогут нам создать эффективные инструменты для вашего бизнеса.
        </p>
        <Button
          variant="outline"
          className="mt-8"
          onClick={() => {
            setSubmitted(false)
            setAnswers({})
            setOtherTexts({})
            setCurrentBlock(0)
          }}
        >
          Заполнить ещё раз
        </Button>
      </motion.div>
    )
  }

  return (
    <div className="max-w-3xl mx-auto">
      {/* Progress */}
      <div className="mb-8">
        <div className="flex items-center justify-between mb-2">
          <span className="text-sm text-muted-foreground">
            Блок {currentBlock + 1} из {totalBlocks}
          </span>
          <span className="text-sm font-medium">{Math.round(progress)}%</span>
        </div>
        <Progress value={progress} className="h-2" />
        <div className="flex gap-1 mt-3">
          {BLOCKS.map((b, i) => (
            <button
              key={b.num}
              onClick={() => setCurrentBlock(i)}
              className={`flex-1 h-2 rounded-full transition-all ${
                i <= currentBlock ? 'bg-primary' : 'bg-muted'
              }`}
              title={b.title}
            />
          ))}
        </div>
      </div>

      {/* Block header */}
      <motion.div
        key={currentBlock}
        initial={{ opacity: 0, x: 20 }}
        animate={{ opacity: 1, x: 0 }}
        transition={{ duration: 0.3 }}
      >
        <div className="flex items-center gap-3 mb-6">
          <span className="text-3xl">{BLOCKS[currentBlock].icon}</span>
          <h2 className="text-xl font-bold">Блок {currentBlock + 1}. {BLOCKS[currentBlock].title}</h2>
        </div>

        {/* Questions */}
        <div className="space-y-8">
          {blockQuestions.map((q, qIdx) => (
            <Card key={q.id} className="border-l-4 border-l-primary/30">
              <CardHeader className="pb-3">
                <CardTitle className="text-base flex items-start gap-2">
                  <span className="text-primary font-bold shrink-0">{qIdx + 1}.</span>
                  <span>{q.text}</span>
                  {q.required && <span className="text-red-500 ml-1">*</span>}
                </CardTitle>
                {q.type === 'multi' && q.maxSelect && (
                  <CardDescription>Выберите до {q.maxSelect} вариантов</CardDescription>
                )}
              </CardHeader>
              <CardContent>
                {q.type === 'single' && q.options && (
                  <RadioGroup
                    value={(answers[q.id] as string) || ''}
                    onValueChange={(val) => handleSingleChange(q.id, val)}
                    className="space-y-2"
                  >
                    {q.options.map((opt) => (
                      <div key={opt.value} className="flex items-start gap-3">
                        <RadioGroupItem value={opt.value} id={`${q.id}_${opt.value}`} className="mt-0.5" />
                        <Label htmlFor={`${q.id}_${opt.value}`} className="font-normal cursor-pointer leading-snug">
                          {opt.label}
                        </Label>
                      </div>
                    ))}
                    {q.options.some((o) => o.hasOther) && answers[q.id] === 'other' && (
                      <Input
                        placeholder="Ваш вариант..."
                        value={otherTexts[`${q.id}_other`] || ''}
                        onChange={(e) => handleOtherText(`${q.id}_other`, e.target.value)}
                        className="mt-2 ml-7"
                      />
                    )}
                  </RadioGroup>
                )}

                {q.type === 'multi' && q.options && (
                  <div className="space-y-2">
                    {q.options.map((opt) => (
                      <div key={opt.value} className="flex items-start gap-3">
                        <Checkbox
                          id={`${q.id}_${opt.value}`}
                          checked={((answers[q.id] as string[]) || []).includes(opt.value)}
                          onCheckedChange={(checked) =>
                            handleMultiChange(q.id, opt.value, !!checked, q.maxSelect)
                          }
                          className="mt-0.5"
                        />
                        <Label htmlFor={`${q.id}_${opt.value}`} className="font-normal cursor-pointer leading-snug">
                          {opt.label}
                        </Label>
                      </div>
                    ))}
                    {q.options.some((o) => o.hasOther) &&
                      ((answers[q.id] as string[]) || []).includes('other') && (
                        <Input
                          placeholder="Ваш вариант..."
                          value={otherTexts[`${q.id}_other`] || ''}
                          onChange={(e) => handleOtherText(`${q.id}_other`, e.target.value)}
                          className="mt-2 ml-7"
                        />
                      )}
                  </div>
                )}

                {q.type === 'scale' && (
                  <div className="space-y-3">
                    <div className="flex items-center justify-between text-xs text-muted-foreground px-1">
                      <span>0 — Точно нет</span>
                      <span>5 — Нейтрально</span>
                      <span>10 — Определённо да</span>
                    </div>
                    <div className="flex gap-1.5 justify-center flex-wrap">
                      {Array.from({ length: 11 }, (_, i) => (
                        <button
                          key={i}
                          type="button"
                          onClick={() => handleScaleChange(q.id, i)}
                          className={`w-10 h-10 rounded-lg text-sm font-semibold transition-all ${
                            answers[q.id] === i
                              ? 'bg-primary text-primary-foreground shadow-md scale-110'
                              : 'bg-muted hover:bg-muted/80 text-muted-foreground'
                          }`}
                        >
                          {i}
                        </button>
                      ))}
                    </div>
                  </div>
                )}

                {q.type === 'text' && (
                  <Textarea
                    placeholder="Ваш ответ..."
                    value={(answers[q.id] as string) || ''}
                    onChange={(e) => handleTextChange(q.id, e.target.value)}
                    rows={3}
                  />
                )}
              </CardContent>
            </Card>
          ))}
        </div>

        {/* Navigation */}
        <div className="flex justify-between mt-8 mb-4">
          <Button
            variant="outline"
            onClick={() => setCurrentBlock((prev) => Math.max(0, prev - 1))}
            disabled={currentBlock === 0}
          >
            <ChevronLeft className="mr-1 h-4 w-4" /> Назад
          </Button>

          {currentBlock < totalBlocks - 1 ? (
            <Button
              onClick={() => setCurrentBlock((prev) => prev + 1)}
              disabled={!canProceed()}
            >
              Далее <ChevronRight className="ml-1 h-4 w-4" />
            </Button>
          ) : (
            <Button
              onClick={handleSubmit}
              disabled={!canProceed() || submitting}
              className="bg-green-600 hover:bg-green-700"
            >
              {submitting ? (
                <>Отправка...</>
              ) : (
                <>
                  <Send className="mr-2 h-4 w-4" /> Отправить ответы
                </>
              )}
            </Button>
          )}
        </div>
      </motion.div>
    </div>
  )
}

/* ------------------------------------------------------------------ */
/*  RESULTS DASHBOARD COMPONENT                                         */
/* ------------------------------------------------------------------ */

interface ResultsData {
  totalResponses: number
  singleChoice: Record<string, Record<string, number>>
  multiChoice: Record<string, Record<string, number>>
  scale: {
    q10: { average: number; distribution: Record<string, number> }
    q17: { average: number; distribution: Record<string, number> }
  }
  freeText: { q6: string[] }
}

function ResultsDashboard() {
  const [data, setData] = useState<ResultsData | null>(null)
  const [loading, setLoading] = useState(true)

  const fetchResults = useCallback(async () => {
    try {
      const res = await fetch('/api/tso-results')
      if (res.ok) {
        const json = await res.json()
        setData(json)
      }
    } catch (e) {
      console.error('Failed to fetch results', e)
    } finally {
      setLoading(false)
    }
  }, [])

  useEffect(() => {
    fetchResults()
  }, [fetchResults])

  if (loading) {
    return (
      <div className="flex items-center justify-center py-20">
        <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-primary" />
      </div>
    )
  }

  if (!data || data.totalResponses === 0) {
    return (
      <div className="flex flex-col items-center justify-center py-20 text-center">
        <ClipboardList className="h-16 w-16 text-muted-foreground mb-4" />
        <h3 className="text-xl font-semibold mb-2">Пока нет ответов</h3>
        <p className="text-muted-foreground">
          Когда респонденты заполнят анкету, результаты появятся здесь.
        </p>
      </div>
    )
  }

  const { totalResponses, singleChoice, multiChoice, scale, freeText } = data

  // Helper: convert frequency to chart data
  const freqToBarData = (freq: Record<string, number>, labelMap?: Record<string, string>) =>
    Object.entries(freq)
      .map(([key, count]) => ({
        name: labelMap?.[key] || key,
        value: count,
        pct: Math.round((count / totalResponses) * 100),
      }))
      .sort((a, b) => b.value - a.value)

  const distToBarData = (dist: Record<string, number>) =>
    Object.entries(dist).map(([key, count]) => ({
      name: key,
      value: count,
    }))

  const SingleChoiceChart = ({ qId, title }: { qId: string; title: string }) => {
    const freq = singleChoice[qId]
    if (!freq || Object.keys(freq).length === 0) return null
    const chartData = freqToBarData(freq, LABEL_MAPS[qId])
    return (
      <Card>
        <CardHeader className="pb-2">
          <CardTitle className="text-sm font-semibold">{title}</CardTitle>
        </CardHeader>
        <CardContent>
          <ResponsiveContainer width="100%" height={chartData.length * 36 + 20}>
            <BarChart data={chartData} layout="vertical" margin={{ left: 10, right: 20, top: 5, bottom: 5 }}>
              <CartesianGrid strokeDasharray="3 3" horizontal={false} />
              <XAxis type="number" allowDecimals={false} />
              <YAxis dataKey="name" type="category" width={160} tick={{ fontSize: 12 }} />
              <RTooltip formatter={(val: number) => [`${val} (${Math.round(val / totalResponses * 100)}%)`, 'Ответов']} />
              <Bar dataKey="value" radius={[0, 4, 4, 0]}>
                {chartData.map((_, i) => (
                  <Cell key={i} fill={COLORS[i % COLORS.length]} />
                ))}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </CardContent>
      </Card>
    )
  }

  const MultiChoiceChart = ({ qId, title }: { qId: string; title: string }) => {
    const freq = multiChoice[qId]
    if (!freq || Object.keys(freq).length === 0) return null
    const chartData = freqToBarData(freq, LABEL_MAPS[qId])
    return (
      <Card>
        <CardHeader className="pb-2">
          <CardTitle className="text-sm font-semibold">{title}</CardTitle>
        </CardHeader>
        <CardContent>
          <ResponsiveContainer width="100%" height={chartData.length * 36 + 20}>
            <BarChart data={chartData} layout="vertical" margin={{ left: 10, right: 20, top: 5, bottom: 5 }}>
              <CartesianGrid strokeDasharray="3 3" horizontal={false} />
              <XAxis type="number" allowDecimals={false} />
              <YAxis dataKey="name" type="category" width={160} tick={{ fontSize: 12 }} />
              <RTooltip formatter={(val: number) => [`${val} ответов`, 'Выбрали']} />
              <Bar dataKey="value" radius={[0, 4, 4, 0]}>
                {chartData.map((_, i) => (
                  <Cell key={i} fill={COLORS[i % COLORS.length]} />
                ))}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </CardContent>
      </Card>
    )
  }

  const ScaleChart = ({
    qId,
    title,
    avg,
    dist,
  }: {
    qId: string
    title: string
    avg: number
    dist: Record<string, number>
  }) => {
    const chartData = distToBarData(dist)
    const pieData = [
      { name: '0–3 (Низкая)', value: Object.entries(dist).filter(([k]) => Number(k) <= 3).reduce((s, [, v]) => s + v, 0) },
      { name: '4–7 (Средняя)', value: Object.entries(dist).filter(([k]) => Number(k) >= 4 && Number(k) <= 7).reduce((s, [, v]) => s + v, 0) },
      { name: '8–10 (Высокая)', value: Object.entries(dist).filter(([k]) => Number(k) >= 8).reduce((s, [, v]) => s + v, 0) },
    ].filter((d) => d.value > 0)

    return (
      <Card>
        <CardHeader className="pb-2">
          <CardTitle className="text-sm font-semibold">{title}</CardTitle>
          <div className="flex items-center gap-2 mt-1">
            <Badge variant="secondary" className="text-base">
              Среднее: {avg}
            </Badge>
          </div>
        </CardHeader>
        <CardContent>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <ResponsiveContainer width="100%" height={200}>
              <BarChart data={chartData} margin={{ left: 0, right: 10, top: 5, bottom: 5 }}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis dataKey="name" tick={{ fontSize: 11 }} />
                <YAxis allowDecimals={false} />
                <RTooltip />
                <Bar dataKey="value" radius={[4, 4, 0, 0]}>
                  {chartData.map((entry) => {
                    const v = Number(entry.name)
                    const color = v <= 3 ? '#ef4444' : v <= 7 ? '#f59e0b' : '#22c55e'
                    return <Cell key={entry.name} fill={color} />
                  })}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
            <ResponsiveContainer width="100%" height={200}>
              <PieChart>
                <Pie data={pieData} dataKey="value" cx="50%" cy="50%" outerRadius={70} label={({ name, percent }) => `${name}: ${(percent * 100).toFixed(0)}%`}>
                  {pieData.map((_, i) => (
                    <Cell key={i} fill={['#ef4444', '#f59e0b', '#22c55e'][i]} />
                  ))}
                </Pie>
                <Legend />
                <RTooltip />
              </PieChart>
            </ResponsiveContainer>
          </div>
        </CardContent>
      </Card>
    )
  }

  return (
    <div className="max-w-5xl mx-auto space-y-8">
      {/* Summary cards */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
        <Card>
          <CardContent className="pt-6 text-center">
            <Users className="h-8 w-8 mx-auto mb-2 text-primary" />
            <div className="text-3xl font-bold">{totalResponses}</div>
            <div className="text-sm text-muted-foreground">Ответов</div>
          </CardContent>
        </Card>
        <Card>
          <CardContent className="pt-6 text-center">
            <TrendingUp className="h-8 w-8 mx-auto mb-2 text-green-500" />
            <div className="text-3xl font-bold">{scale.q10.average}</div>
            <div className="text-sm text-muted-foreground">Влияние ТСО (ср.)</div>
          </CardContent>
        </Card>
        <Card>
          <CardContent className="pt-6 text-center">
            <CheckCircle2 className="h-8 w-8 mx-auto mb-2 text-blue-500" />
            <div className="text-3xl font-bold">{scale.q17.average}</div>
            <div className="text-sm text-muted-foreground">Перспективность ТСО (ср.)</div>
          </CardContent>
        </Card>
        <Card>
          <CardContent className="pt-6 text-center">
            <Clock className="h-8 w-8 mx-auto mb-2 text-amber-500" />
            <div className="text-3xl font-bold">
              {singleChoice.q15?.immediately || 0}
            </div>
            <div className="text-sm text-muted-foreground">Готовы внедрить сразу</div>
          </CardContent>
        </Card>
      </div>

      <Separator />

      {/* Block 1: Role & Experience */}
      <div>
        <h3 className="text-lg font-bold mb-4 flex items-center gap-2">
          <span className="text-2xl">👤</span> Блок 1. Роль и опыт
        </h3>
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          <SingleChoiceChart qId="q1" title="Должность" />
          <SingleChoiceChart qId="q2" title="Опыт работы в ЛКМ" />
          <SingleChoiceChart qId="q3" title="Формат продаж" />
        </div>
      </div>

      {/* Block 2: Understanding TCO */}
      <div>
        <h3 className="text-lg font-bold mb-4 flex items-center gap-2">
          <span className="text-2xl">💡</span> Блок 2. Понимание ТСО
        </h3>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <SingleChoiceChart qId="q4" title="Знакомство с ТСО" />
          <MultiChoiceChart qId="q5" title="Составляющие стоимости покраски" />
        </div>
        {freeText.q6.length > 0 && (
          <Card className="mt-4">
            <CardHeader className="pb-2">
              <CardTitle className="text-sm font-semibold">Примеры из практики (вопрос 6)</CardTitle>
            </CardHeader>
            <CardContent>
              <ul className="space-y-2 text-sm text-muted-foreground">
                {freeText.q6.map((t, i) => (
                  <li key={i} className="border-l-2 border-primary/30 pl-3">{t}</li>
                ))}
              </ul>
            </CardContent>
          </Card>
        )}
      </div>

      {/* Block 3: Current approach */}
      <div>
        <h3 className="text-lg font-bold mb-4 flex items-center gap-2">
          <span className="text-2xl">🛒</span> Блок 3. Текущий подход к продаже ЛКМ
        </h3>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <SingleChoiceChart qId="q7" title="Метод рекомендации ЛКМ" />
          <SingleChoiceChart qId="q8" title="Лучший аргумент для премиума" />
        </div>
        <div className="mt-4">
          <MultiChoiceChart qId="q9" title='Фразы покупателей «по цене»' />
        </div>
      </div>

      {/* Block 4: TCO as sales tool */}
      <div>
        <h3 className="text-lg font-bold mb-4 flex items-center gap-2">
          <span className="text-2xl">🎯</span> Блок 4. ТСО как инструмент продажи
        </h3>
        <ScaleChart qId="q10" title="Влияние расчёта ТСО на решение покупателя" avg={scale.q10.average} dist={scale.q10.distribution} />
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mt-4">
          <SingleChoiceChart qId="q11" title="Целевая аудитория для ТСО" />
          <MultiChoiceChart qId="q12" title="Продукты для продажи через ТСО" />
        </div>
      </div>

      {/* Block 5: Barriers & Needs */}
      <div>
        <h3 className="text-lg font-bold mb-4 flex items-center gap-2">
          <span className="text-2xl">🚧</span> Блок 5. Барьеры и потребности
        </h3>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <MultiChoiceChart qId="q13" title="Барьеры для ТСО" />
          <MultiChoiceChart qId="q14" title="Нужные инструменты" />
        </div>
        <div className="mt-4">
          <SingleChoiceChart qId="q15" title="Готовность внедрить инструменты ТСО" />
        </div>
      </div>

      {/* Block 6: Expectations */}
      <div>
        <h3 className="text-lg font-bold mb-4 flex items-center gap-2">
          <span className="text-2xl">🤝</span> Блок 6. Ожидания от производителя
        </h3>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <MultiChoiceChart qId="q16" title="Ожидания от производителя" />
          <ScaleChart qId="q17" title="Перспективность ТСО для бизнеса" avg={scale.q17.average} dist={scale.q17.distribution} />
        </div>
      </div>
    </div>
  )
}

/* ------------------------------------------------------------------ */
/*  MAIN PAGE                                                           */
/* ------------------------------------------------------------------ */

export default function Home() {
  const [activeTab, setActiveTab] = useState('survey')
  const [refreshKey, setRefreshKey] = useState(0)

  const handleSurveySubmitted = () => {
    setRefreshKey((k) => k + 1)
  }

  return (
    <div className="min-h-screen flex flex-col bg-gradient-to-b from-slate-50 to-white">
      {/* Header */}
      <header className="border-b bg-white/80 backdrop-blur-sm sticky top-0 z-50">
        <div className="max-w-6xl mx-auto px-4 py-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 bg-red-600 rounded-lg flex items-center justify-center">
                <span className="text-white font-black text-lg">T</span>
              </div>
              <div>
                <h1 className="text-xl font-bold tracking-tight">TITAN | ТСО</h1>
                <p className="text-xs text-muted-foreground">
                  Совокупная стоимость владения как инструмент продажи ЛКМ
                </p>
              </div>
            </div>
            <Badge variant="outline" className="text-xs">
              Empils
            </Badge>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="flex-1 px-4 py-6">
        <div className="flex justify-center mb-6">
          <div className="inline-flex h-10 items-center justify-center rounded-md bg-muted p-1 text-muted-foreground grid w-full max-w-md grid-cols-2">
            <button
              onClick={() => setActiveTab('survey')}
              className={`inline-flex items-center justify-center gap-2 whitespace-nowrap rounded-sm px-3 py-1.5 text-sm font-medium ring-offset-background transition-all focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2 ${
                activeTab === 'survey'
                  ? 'bg-background text-foreground shadow-sm'
                  : 'hover:bg-background/50'
              }`}
            >
              <ClipboardList className="h-4 w-4" />
              Анкета
            </button>
            <button
              onClick={() => setActiveTab('results')}
              className={`inline-flex items-center justify-center gap-2 whitespace-nowrap rounded-sm px-3 py-1.5 text-sm font-medium ring-offset-background transition-all focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2 ${
                activeTab === 'results'
                  ? 'bg-background text-foreground shadow-sm'
                  : 'hover:bg-background/50'
              }`}
            >
              <BarChart3 className="h-4 w-4" />
              Результаты
            </button>
          </div>
        </div>

        <div style={{ display: activeTab === 'survey' ? 'block' : 'none' }}>
          <SurveyForm onSubmit={handleSurveySubmitted} />
        </div>
        <div style={{ display: activeTab === 'results' ? 'block' : 'none' }}>
          <ResultsDashboard key={refreshKey} />
        </div>
      </main>

      {/* Footer */}
      <footer className="border-t py-4 text-center text-xs text-muted-foreground bg-white">
        Анкета ТСО • Анонимно • Результаты помогут подготовить программу сотрудничества
      </footer>
    </div>
  )
}
