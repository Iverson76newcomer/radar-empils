'use client'

import { useState, useEffect } from 'react'
import { Card } from '@/components/ui/card'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Textarea } from '@/components/ui/textarea'
import { Label } from '@/components/ui/label'
import { Badge } from '@/components/ui/badge'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs'
import { ScrollArea } from '@/components/ui/scroll-area'
import { toast } from 'sonner'
import {
  Eye, EyeOff, AlertTriangle, Lightbulb, Package,
  Swords, Send, Camera, BarChart3, Clock, MapPin, User, CheckCircle2,
  XCircle, Loader2, Filter, FileSpreadsheet, FileText, Download,
  Tag, Building2
} from 'lucide-react'
import {
  createMessage,
  getMessages,
  updateMessageStatus,
  getDashboardStats,
  type Message,
} from '@/lib/supabase'
import { exportToExcel, exportToPDF } from '@/lib/export'

// Используем импортированный тип Message

// Сжатие изображения на клиенте — возвращает JPEG data URL
// Агрессивное сжатие для Netlify Function timeout (10 сек на бесплатном тарифе)
async function compressImage(file: File, maxSize: number, quality: number): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = (e) => {
      const img = new Image()
      img.onload = () => {
        // Сохраняем пропорции, ограничиваем максимальный размер
        let { width, height } = img
        if (width > height && width > maxSize) {
          height = (height * maxSize) / width
          width = maxSize
        } else if (height > maxSize) {
          width = (width * maxSize) / height
          height = maxSize
        }

        const canvas = document.createElement('canvas')
        canvas.width = width
        canvas.height = height
        const ctx = canvas.getContext('2d')
        if (!ctx) {
          reject(new Error('Canvas context не доступен'))
          return
        }
        // Белый фон нужен для PNG с прозрачностью (чтобы не было чёрного)
        ctx.fillStyle = '#ffffff'
        ctx.fillRect(0, 0, width, height)
        ctx.drawImage(img, 0, 0, width, height)
        let dataUrl = canvas.toDataURL('image/jpeg', quality)

        // Если всё ещё больше 500 КБ — сжимаем ещё сильнее
        let currentQuality = quality
        while (dataUrl.length > 500 * 1024 && currentQuality > 0.3) {
          currentQuality -= 0.15
          dataUrl = canvas.toDataURL('image/jpeg', currentQuality)
        }

        // Если всё ещё слишком большой — уменьшаем размер
        let currentMaxSize = maxSize
        while (dataUrl.length > 500 * 1024 && currentMaxSize > 400) {
          currentMaxSize = Math.floor(currentMaxSize * 0.7)
          const newW = currentMaxSize
          const newH = Math.floor((height * currentMaxSize) / width)
          canvas.width = newW
          canvas.height = newH
          ctx.fillStyle = '#ffffff'
          ctx.fillRect(0, 0, newW, newH)
          ctx.drawImage(img, 0, 0, newW, newH)
          dataUrl = canvas.toDataURL('image/jpeg', 0.5)
        }

        resolve(dataUrl)
      }
      img.onerror = () => reject(new Error('Не удалось загрузить изображение'))
      img.src = e.target?.result as string
    }
    reader.onerror = () => reject(new Error('Не удалось прочитать файл'))
    reader.readAsDataURL(file)
  })
}

// Категории наблюдений (порядок: верхний ряд 3 + нижний ряд 3)
const CATEGORIES = [
  // Верхний ряд — про риски/внешнюю среду
  { id: 'competitors', label: 'Конкуренты', icon: Swords, color: 'bg-red-500', textColor: 'text-red-500', bgLight: 'bg-red-50' },
  { id: 'prices',      label: 'Цены и акции', icon: Tag, color: 'bg-orange-500', textColor: 'text-orange-500', bgLight: 'bg-orange-50' },
  { id: 'problem',     label: 'Проблема', icon: AlertTriangle, color: 'bg-yellow-500', textColor: 'text-yellow-500', bgLight: 'bg-yellow-50' },
  // Нижний ряд — про возможности/улучшения
  { id: 'product',     label: 'Запрос продукта', icon: Package, color: 'bg-blue-500', textColor: 'text-blue-500', bgLight: 'bg-blue-50' },
  { id: 'idea',        label: 'Идея оформления', icon: Lightbulb, color: 'bg-green-500', textColor: 'text-green-500', bgLight: 'bg-green-50' },
  { id: 'portfolio',   label: 'Объекты для портфолио', icon: Building2, color: 'bg-purple-500', textColor: 'text-purple-500', bgLight: 'bg-purple-50' },
]

const REGIONS = [
  'Москва и МО', 'Санкт-Петербург и ЛО', 'Красноярский край', 'Республика Хакасия',
  'Республика Бурятия', 'Новосибирская обл.', 'Томская обл.', 'Кемеровская обл.',
  'Алтайский край', 'Республика Алтай', 'Республика Саха (Якутия)',
  'Забайкальский край', 'Воронежская обл.', 'Липецкая обл.', 'Вологодская обл.',
  'Другой регион',
]

const ROLES = [
  { id: 'RM', label: 'Региональный менеджер' },
  { id: 'ETP', label: 'Эксклюзивный ТП' },
  { id: 'merchandiser', label: 'Мерчендайзер' },
]

const STATUSES = [
  { id: 'new', label: 'Новое', color: 'bg-blue-100 text-blue-700' },
  { id: 'accepted', label: 'Принято', color: 'bg-green-100 text-green-700' },
  { id: 'in_progress', label: 'В работе', color: 'bg-yellow-100 text-yellow-700' },
  { id: 'closed', label: 'Закрыто', color: 'bg-gray-100 text-gray-700' },
  { id: 'rejected', label: 'Отклонено', color: 'bg-red-100 text-red-700' },
]

export default function RadarPage() {
  const [tab, setTab] = useState<'report' | 'dashboard'>('report')
  const [selectedCategory, setSelectedCategory] = useState<string | null>(null)
  const [text, setText] = useState('')
  const [authorName, setAuthorName] = useState('')
  const [authorRole, setAuthorRole] = useState('')
  const [region, setRegion] = useState('')
  const [photoUrl, setPhotoUrl] = useState<string | null>(null)
  const [submitting, setSubmitting] = useState(false)

  // Dashboard state
  const [messages, setMessages] = useState<Message[]>([])
  const [loading, setLoading] = useState(false)
  const [filterCategory, setFilterCategory] = useState<string>('all')
  const [filterStatus, setFilterStatus] = useState<string>('all')
  const [stats, setStats] = useState<any>(null)

  // PIN-авторизация для дашборда
  const [isAuthed, setIsAuthed] = useState(false)
  const [showPinModal, setShowPinModal] = useState(false)
  const [pinInput, setPinInput] = useState('')
  const [pinError, setPinError] = useState(false)
  const DASHBOARD_PIN = process.env.NEXT_PUBLIC_DASHBOARD_PIN || ''

  // Состояние экспорта
  const [exporting, setExporting] = useState<'excel' | 'pdf' | null>(null)

  // Экспорт в Excel
  const handleExportExcel = async () => {
    if (messages.length === 0) {
      toast.error('Нет данных для экспорта')
      return
    }
    setExporting('excel')
    try {
      await exportToExcel(messages)
      toast.success('Excel-отчёт сформирован')
    } catch (err: any) {
      console.error('Excel export error:', err)
      toast.error(`Ошибка экспорта: ${err?.message || err}`)
    } finally {
      setExporting(null)
    }
  }

  // Экспорт в PDF
  const handleExportPDF = async () => {
    if (messages.length === 0) {
      toast.error('Нет данных для экспорта')
      return
    }
    setExporting('pdf')
    try {
      await exportToPDF(messages)
      toast.success('PDF-отчёт сформирован')
    } catch (err: any) {
      console.error('PDF export error:', err)
      toast.error(`Ошибка экспорта: ${err?.message || err}`)
    } finally {
      setExporting(null)
    }
  }

  // Проверка авторизации из sessionStorage при загрузке
  useEffect(() => {
    if (sessionStorage.getItem('radar_authed') === 'true' && DASHBOARD_PIN) {
      setIsAuthed(true)
    }
  }, [])

  // Обработчик переключения вкладок
  const handleTabChange = (v: string) => {
    if (v === 'dashboard' && !isAuthed && DASHBOARD_PIN) {
      // Запросить PIN-код перед открытием дашборда
      setShowPinModal(true)
      setPinInput('')
      setPinError(false)
      return
    }
    setTab(v as 'report' | 'dashboard')
  }

  // Проверка PIN-кода
  const handlePinSubmit = () => {
    if (pinInput === DASHBOARD_PIN) {
      setIsAuthed(true)
      setShowPinModal(false)
      setPinInput('')
      setPinError(false)
      sessionStorage.setItem('radar_authed', 'true')
      setTab('dashboard')
      toast.success('Доступ к дашборду открыт')
    } else {
      setPinError(true)
      setPinInput('')
    }
  }

  // Выход из дашборда
  const handleLogout = () => {
    setIsAuthed(false)
    sessionStorage.removeItem('radar_authed')
    setTab('report')
    toast.info('Вы вышли из дашборда')
  }

  // Загрузить сохранённые данные сотрудника
  useEffect(() => {
    const saved = localStorage.getItem('radar_user')
    if (saved) {
      const user = JSON.parse(saved)
      setAuthorName(user.name || '')
      setAuthorRole(user.role || '')
      setRegion(user.region || '')
    }
  }, [])

  // Сохранить данные сотрудника
  const saveUser = () => {
    localStorage.setItem('radar_user', JSON.stringify({ name: authorName, role: authorRole, region }))
  }

  // Обработка фото — со сжатием на клиенте (чтобы влезть в лимит Netlify 6 МБ)
  const handlePhoto = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (!file) return

    // Если файл слишком большой (больше 10 МБ) — не даём загрузить
    if (file.size > 10 * 1024 * 1024) {
      toast.error('Фото слишком большое (макс. 10 МБ)')
      return
    }

    try {
      // Сжимаем изображение через canvas — агрессивные настройки для Netlify
      const compressed = await compressImage(file, 800, 0.6)
      setPhotoUrl(compressed)
      const sizeKb = (compressed.length / 1024).toFixed(0)
      toast.success(`Фото обработано: ${sizeKb} КБ`)
      if (compressed.length > 500 * 1024) {
        toast.warning('Фото большое — отправка может занять до 10 секунд')
      }
    } catch (err) {
      console.error('Photo compression error:', err)
      toast.error('Не удалось обработать фото. Попробуйте другое.')
    }
  }

  // Отправка сообщения
  const handleSubmit = async () => {
    if (!selectedCategory) { toast.error('Выберите категорию'); return }
    if (!text.trim()) { toast.error('Введите текст сообщения'); return }
    if (!authorName.trim()) { toast.error('Введите ваше имя'); return }

    setSubmitting(true)
    saveUser()
    try {
      // Прямой запрос к Supabase из браузера — НЕ тратит кредиты Netlify!
      await createMessage({
        category: selectedCategory,
        text: text.trim(),
        region: region || 'Не указан',
        authorName: authorName.trim(),
        authorRole: authorRole || 'Не указана',
        photoUrl,
      })
      toast.success('Сообщение отправлено в Радар!')
      setText('')
      setPhotoUrl(null)
      setSelectedCategory(null)
    } catch (err: any) {
      console.error('Submit error:', err)
      const msg = err?.message || String(err)
      if (msg.includes('row-level security') || msg.includes('RLS') || msg.includes('policy')) {
        toast.error('Нет прав на запись. Проверьте RLS-политики в Supabase.')
      } else if (msg.includes('relation') && msg.includes('does not exist')) {
        toast.error('Таблица messages не найдена. Создайте её через SQL Editor в Supabase.')
      } else if (msg.includes('Failed to fetch') || msg.includes('NetworkError')) {
        toast.error('Ошибка сети. Проверьте интернет-соединение.')
      } else {
        toast.error(`Ошибка при отправке: ${msg}`)
      }
    } finally {
      setSubmitting(false)
    }
  }

  // Загрузка данных дашборда — напрямую из Supabase
  const loadDashboard = async () => {
    setLoading(true)
    try {
      const [msgs, statsData] = await Promise.all([
        getMessages({ category: filterCategory, status: filterStatus }),
        getDashboardStats(),
      ])
      setMessages(msgs)
      setStats(statsData)
    } catch (err: any) {
      console.error('Dashboard load error:', err)
      toast.error(`Ошибка загрузки дашборда: ${err?.message || err}`)
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => {
    if (tab === 'dashboard') loadDashboard()
  }, [tab, filterCategory, filterStatus])

  // Обновить статус сообщения — напрямую через Supabase
  const updateStatus = async (id: string, status: string) => {
    try {
      await updateMessageStatus(id, status)
      toast.success('Статус обновлён')
      loadDashboard()
    } catch (err: any) {
      toast.error(`Ошибка обновления: ${err?.message || err}`)
    }
  }

  const getCategoryInfo = (catId: string) => CATEGORIES.find(c => c.id === catId)
  const getStatusInfo = (statusId: string) => STATUSES.find(s => s.id === statusId)
  const formatDate = (d: string) => new Date(d).toLocaleString('ru-RU', {
    day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit'
  })

  return (
    <div className="min-h-screen bg-gradient-to-b from-slate-50 to-slate-100">
      {/* Header */}
      <header className="sticky top-0 z-50 bg-white border-b border-slate-200 shadow-sm">
        <div className="max-w-5xl mx-auto px-4 py-3 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="w-8 h-8 rounded-lg bg-gradient-to-br from-slate-800 to-slate-600 flex items-center justify-center">
              <svg viewBox="0 0 24 24" className="w-5 h-5 text-white" fill="currentColor">
                <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 18c-4.41 0-8-3.59-8-8s3.59-8 8-8 8 3.59 8 8-3.59 8-8 8zm-1-13h2v6h-2zm0 8h2v2h-2z"/>
              </svg>
            </div>
            <div>
              <h1 className="text-lg font-bold text-slate-800">Радар Эмпилс</h1>
              <p className="text-xs text-slate-500">Полевой информационный портал</p>
            </div>
          </div>
          <Tabs value={tab} onValueChange={handleTabChange}>
            <TabsList>
              <TabsTrigger value="report" className="text-sm">Сообщить</TabsTrigger>
              <TabsTrigger value="dashboard" className="text-sm">
                <BarChart3 className="w-4 h-4 mr-1" />
                Дашборд
              </TabsTrigger>
            </TabsList>
          </Tabs>
        </div>
      </header>

      {/* Модальное окно ввода PIN-кода */}
      {showPinModal && (
        <div className="fixed inset-0 z-[100] bg-black/50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-2xl p-6 w-full max-w-sm">
            <div className="flex items-center justify-between mb-4">
              <h2 className="text-lg font-bold text-slate-800">🔒 Доступ к дашборду</h2>
              <button
                onClick={() => { setShowPinModal(false); setPinInput(''); setPinError(false) }}
                className="text-slate-400 hover:text-slate-600 text-2xl leading-none"
                aria-label="Закрыть"
              >×</button>
            </div>
            <p className="text-sm text-slate-600 mb-4">
              Введите PIN-код для доступа к дашборду руководителя
            </p>
            <input
              type="password"
              inputMode="numeric"
              pattern="[0-9]*"
              maxLength={8}
              value={pinInput}
              onChange={(e) => {
                setPinInput(e.target.value.replace(/\D/g, ''))
                setPinError(false)
              }}
              onKeyDown={(e) => {
                if (e.key === 'Enter' && pinInput.length >= 4) handlePinSubmit()
              }}
              autoFocus
              placeholder="••••"
              className={`w-full text-center text-3xl tracking-[0.5em] py-3 px-4 border-2 rounded-xl outline-none transition-colors ${
                pinError
                  ? 'border-red-500 bg-red-50 text-red-700'
                  : 'border-slate-200 focus:border-blue-500 bg-slate-50 text-slate-800'
              }`}
            />
            {pinError && (
              <p className="text-sm text-red-600 mt-2 text-center">
                Неверный PIN-код. Попробуйте снова.
              </p>
            )}
            <button
              onClick={handlePinSubmit}
              disabled={pinInput.length < 4}
              className="w-full mt-4 py-3 bg-slate-800 hover:bg-slate-900 text-white font-semibold rounded-xl transition-colors disabled:opacity-40 disabled:cursor-not-allowed"
            >
              Войти
            </button>
            <p className="text-xs text-slate-400 mt-3 text-center">
              PIN-код доступен руководителям. Обратитесь к администратору.
            </p>
          </div>
        </div>
      )}

      <main className="max-w-5xl mx-auto px-4 py-6 pb-20">
        {/* === ВКЛАДКА: СООБЩИТЬ === */}
        {tab === 'report' && (
          <div className="space-y-4">
            {/* Шаг 1: Выбор категории */}
            <div>
              <Label className="text-sm font-semibold text-slate-700 mb-2 block">
                1. Что вы заметили?
              </Label>
              <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
                {CATEGORIES.map((cat) => {
                  const Icon = cat.icon
                  const isSelected = selectedCategory === cat.id
                  return (
                    <button
                      key={cat.id}
                      onClick={() => setSelectedCategory(cat.id)}
                      className={`
                        flex flex-col items-center gap-1.5 p-3 rounded-xl border-2 transition-all
                        ${isSelected
                          ? `${cat.bgLight} border-current ${cat.textColor}`
                          : 'bg-white border-slate-200 text-slate-500 hover:border-slate-300'
                        }
                      `}
                    >
                      <Icon className="w-6 h-6" />
                      <span className="text-xs font-medium text-center">{cat.label}</span>
                    </button>
                  )
                })}
              </div>
            </div>

            {/* Шаг 2: Данные сотрудника — ВСЕГДА ВИДНЫ, автозаполнение из localStorage */}
            <Card className="p-4">
              <Label className="text-sm font-semibold text-slate-700 mb-3 block">
                Ваши данные (сохранятся для следующих сообщений)
              </Label>
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                <div>
                  <Label className="text-xs text-slate-500">Ваше имя</Label>
                  <Input
                    value={authorName}
                    onChange={(e) => setAuthorName(e.target.value)}
                    placeholder="Иван Иванов"
                    className="mt-1"
                  />
                </div>
                <div>
                  <Label className="text-xs text-slate-500">Роль</Label>
                  <Select value={authorRole} onValueChange={setAuthorRole}>
                    <SelectTrigger className="mt-1"><SelectValue placeholder="Выберите" /></SelectTrigger>
                    <SelectContent>
                      {ROLES.map(r => <SelectItem key={r.id} value={r.id}>{r.label}</SelectItem>)}
                    </SelectContent>
                  </Select>
                </div>
                <div>
                  <Label className="text-xs text-slate-500">Регион</Label>
                  <Select value={region} onValueChange={setRegion}>
                    <SelectTrigger className="mt-1"><SelectValue placeholder="Выберите" /></SelectTrigger>
                    <SelectContent>
                      {REGIONS.map(r => <SelectItem key={r} value={r}>{r}</SelectItem>)}
                    </SelectContent>
                  </Select>
                </div>
              </div>
            </Card>

            {/* Шаг 3: Текст сообщения */}
            <div>
              <Label className="text-sm font-semibold text-slate-700 mb-2 block">
                2. Опишите наблюдение
              </Label>
              <Textarea
                value={text}
                onChange={(e) => setText(e.target.value)}
                placeholder="Например: В Леруа на Ленина увидел новинку от Текса — краска для обоев, цена 450 руб за 0.9 кг. Стоит на отдельном дисплее."
                className="min-h-[120px] resize-none"
              />
              <p className="text-xs text-slate-400 mt-1">
                {text.length} символов
              </p>
            </div>

            {/* Шаг 4: Фото */}
            <div>
              <Label className="text-sm font-semibold text-slate-700 mb-2 block">
                3. Фото (необязательно)
              </Label>
              {photoUrl ? (
                <div className="relative inline-block">
                  <img src={photoUrl} alt="Прикреплённое фото" className="max-h-48 rounded-lg border border-slate-200" />
                  <button
                    onClick={() => setPhotoUrl(null)}
                    className="absolute -top-2 -right-2 bg-red-500 text-white rounded-full w-6 h-6 flex items-center justify-center text-xs hover:bg-red-600"
                  >
                    ✕
                  </button>
                </div>
              ) : (
                <label className="flex flex-col items-center justify-center w-full h-32 border-2 border-dashed border-slate-300 rounded-xl cursor-pointer hover:border-slate-400 transition-colors bg-slate-50">
                  <Camera className="w-6 h-6 text-slate-400" />
                  <span className="text-xs text-slate-400 mt-1">Нажмите, чтобы прикрепить фото</span>
                  <input type="file" accept="image/*" capture="environment" onChange={handlePhoto} className="hidden" />
                </label>
              )}
            </div>

            {/* Отправка */}
            <Button
              onClick={handleSubmit}
              disabled={submitting || !selectedCategory || !text.trim() || !authorName.trim()}
              className="w-full h-12 text-base"
              size="lg"
            >
              {submitting ? (
                <><Loader2 className="w-5 h-5 mr-2 animate-spin" /> Отправка...</>
              ) : (
                <><Send className="w-5 h-5 mr-2" /> Отправить в Радар</>
              )}
            </Button>

            {/* Информация о том, что данные сохранены */}
            {authorName && (
              <div className="text-center text-xs text-slate-400">
                ✓ Данные сохранены в этом браузере
              </div>
            )}
          </div>
        )}

        {/* === ВКЛАДКА: ДАШБОРД === */}
        {tab === 'dashboard' && (
          <div className="space-y-4">
            {/* Статистика */}
            {stats && (
              <div className="grid grid-cols-2 sm:grid-cols-5 gap-2">
                {CATEGORIES.map(cat => (
                  <Card key={cat.id} className="p-3 text-center">
                    <div className={`text-2xl font-bold ${cat.textColor}`}>
                      {stats.byCategory[cat.id] || 0}
                    </div>
                    <div className="text-xs text-slate-500 mt-1">{cat.label}</div>
                  </Card>
                ))}
              </div>
            )}

            {/* Фильтры */}
            <div className="flex flex-wrap gap-2 items-center">
              <Filter className="w-4 h-4 text-slate-400" />
              <Select value={filterCategory} onValueChange={setFilterCategory}>
                <SelectTrigger className="w-40 h-8 text-xs"><SelectValue /></SelectTrigger>
                <SelectContent>
                  <SelectItem value="all">Все категории</SelectItem>
                  {CATEGORIES.map(c => <SelectItem key={c.id} value={c.id}>{c.label}</SelectItem>)}
                </SelectContent>
              </Select>
              <Select value={filterStatus} onValueChange={setFilterStatus}>
                <SelectTrigger className="w-32 h-8 text-xs"><SelectValue /></SelectTrigger>
                <SelectContent>
                  <SelectItem value="all">Все статусы</SelectItem>
                  {STATUSES.map(s => <SelectItem key={s.id} value={s.id}>{s.label}</SelectItem>)}
                </SelectContent>
              </Select>
              <Button variant="outline" size="sm" onClick={loadDashboard} className="h-8">
                <Filter className="w-3 h-3 mr-1" /> Обновить
              </Button>

              {/* Кнопки экспорта */}
              <Button
                variant="outline"
                size="sm"
                onClick={handleExportExcel}
                disabled={exporting !== null || messages.length === 0}
                className="h-8 border-green-600 text-green-700 hover:bg-green-50 hover:text-green-800"
                title="Скачать полный отчёт в Excel (6 листов)"
              >
                {exporting === 'excel' ? (
                  <><Loader2 className="w-3 h-3 mr-1 animate-spin" /> Формирование...</>
                ) : (
                  <><FileSpreadsheet className="w-3 h-3 mr-1" /> Excel</>
                )}
              </Button>
              <Button
                variant="outline"
                size="sm"
                onClick={handleExportPDF}
                disabled={exporting !== null || messages.length === 0}
                className="h-8 border-red-600 text-red-700 hover:bg-red-50 hover:text-red-800"
                title="Скачать PDF-резюме"
              >
                {exporting === 'pdf' ? (
                  <><Loader2 className="w-3 h-3 mr-1 animate-spin" /> Формирование...</>
                ) : (
                  <><FileText className="w-3 h-3 mr-1" /> PDF</>
                )}
              </Button>

              <div className="flex-1" />
              <Button
                variant="ghost"
                size="sm"
                onClick={handleLogout}
                className="h-8 text-red-600 hover:text-red-700 hover:bg-red-50"
                title="Выйти из дашборда"
              >
                <XCircle className="w-3 h-3 mr-1" /> Выйти
              </Button>
            </div>

            {/* Сообщения */}
            {loading ? (
              <div className="flex items-center justify-center py-12">
                <Loader2 className="w-6 h-6 text-slate-400 animate-spin" />
              </div>
            ) : messages.length === 0 ? (
              <Card className="p-8 text-center text-slate-400">
                Сообщений пока нет. Они появятся здесь, когда полевые сотрудники начнут отправлять наблюдения.
              </Card>
            ) : (
              <div className="space-y-3">
                {messages.map((msg) => {
                  const cat = getCategoryInfo(msg.category)
                  const stat = getStatusInfo(msg.status)
                  const CatIcon = cat?.icon || Package
                  return (
                    <Card key={msg.id} className="p-4">
                      <div className="flex items-start gap-3">
                        <div className={`flex-shrink-0 w-10 h-10 rounded-lg ${cat?.bgLight} flex items-center justify-center`}>
                          <CatIcon className={`w-5 h-5 ${cat?.textColor}`} />
                        </div>
                        <div className="flex-1 min-w-0">
                          <div className="flex items-center gap-2 flex-wrap mb-1">
                            <span className={`text-xs font-medium ${cat?.textColor}`}>{cat?.label}</span>
                            <Badge className={`text-xs ${stat?.color}`}>{stat?.label}</Badge>
                          </div>
                          <p className="text-sm text-slate-700 whitespace-pre-wrap break-words">{msg.text}</p>
                          {msg.photoUrl && (
                            <img src={msg.photoUrl} alt="Фото" className="mt-2 max-h-40 rounded-lg border border-slate-200" />
                          )}
                          <div className="flex items-center gap-3 mt-2 text-xs text-slate-400">
                            <span className="flex items-center gap-1"><User className="w-3 h-3" />{msg.authorName}</span>
                            <span className="flex items-center gap-1"><MapPin className="w-3 h-3" />{msg.region}</span>
                            <span className="flex items-center gap-1"><Clock className="w-3 h-3" />{formatDate(msg.createdAt)}</span>
                          </div>
                          {/* Кнопки смены статуса */}
                          <div className="flex gap-1 mt-2">
                            <Button size="sm" variant="outline" className="h-7 text-xs"
                              onClick={() => updateStatus(msg.id, 'accepted')}>
                              <CheckCircle2 className="w-3 h-3 mr-1" /> Принято
                            </Button>
                            <Button size="sm" variant="outline" className="h-7 text-xs"
                              onClick={() => updateStatus(msg.id, 'in_progress')}>
                              <Loader2 className="w-3 h-3 mr-1" /> В работе
                            </Button>
                            <Button size="sm" variant="outline" className="h-7 text-xs"
                              onClick={() => updateStatus(msg.id, 'closed')}>
                              Закрыто
                            </Button>
                            <Button size="sm" variant="outline" className="h-7 text-xs"
                              onClick={() => updateStatus(msg.id, 'rejected')}>
                              <XCircle className="w-3 h-3 mr-1" /> Отклонить
                            </Button>
                          </div>
                          {msg.managerNote && (
                            <div className="mt-2 p-2 bg-blue-50 rounded text-xs text-blue-700">
                              <strong>Комментарий РД:</strong> {msg.managerNote}
                            </div>
                          )}
                        </div>
                      </div>
                    </Card>
                  )
                })}
              </div>
            )}
          </div>
        )}
      </main>

      {/* Footer */}
      <footer className="mt-auto border-t border-slate-200 bg-white py-3">
        <div className="max-w-5xl mx-auto px-4 text-center text-xs text-slate-400">
          ЗАО «Эмпилс» · Радар полевых наблюдений · 8 (800) 200-49-11 · www.empils.ru
        </div>
      </footer>
    </div>
  )
}
