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
  Eye, EyeOff, TrendingUp, AlertTriangle, Lightbulb, Package,
  Swords, Send, Camera, BarChart3, Clock, MapPin, User, CheckCircle2,
  XCircle, Loader2, Filter
} from 'lucide-react'

// Категории
const CATEGORIES = [
  { id: 'competitors', label: 'Конкуренты', icon: Swords, color: 'bg-red-500', textColor: 'text-red-500', bgLight: 'bg-red-50' },
  { id: 'product', label: 'Запрос продукта', icon: Package, color: 'bg-blue-500', textColor: 'text-blue-500', bgLight: 'bg-blue-50' },
  { id: 'idea', label: 'Идея оформления', icon: Lightbulb, color: 'bg-green-500', textColor: 'text-green-500', bgLight: 'bg-green-50' },
  { id: 'problem', label: 'Проблема', icon: AlertTriangle, color: 'bg-yellow-500', textColor: 'text-yellow-500', bgLight: 'bg-yellow-50' },
  { id: 'trend', label: 'Тренд рынка', icon: TrendingUp, color: 'bg-purple-500', textColor: 'text-purple-500', bgLight: 'bg-purple-50' },
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

interface Message {
  id: string
  category: string
  text: string
  region: string
  authorName: string
  authorRole: string
  photoUrl: string | null
  status: string
  managerNote: string | null
  createdAt: string
}

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

  // Обработка фото
  const handlePhoto = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (!file) return
    if (file.size > 3 * 1024 * 1024) {
      toast.error('Фото слишком большое (макс. 3 МБ)')
      return
    }
    const reader = new FileReader()
    reader.onload = () => setPhotoUrl(reader.result as string)
    reader.readAsDataURL(file)
  }

  // Отправка сообщения
  const handleSubmit = async () => {
    if (!selectedCategory) { toast.error('Выберите категорию'); return }
    if (!text.trim()) { toast.error('Введите текст сообщения'); return }
    if (!authorName.trim()) { toast.error('Введите ваше имя'); return }

    setSubmitting(true)
    saveUser()
    try {
      const res = await fetch('/api/radar', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          category: selectedCategory,
          text: text.trim(),
          region: region || 'Не указан',
          authorName: authorName.trim(),
          authorRole: authorRole || 'Не указана',
          photoUrl,
        }),
      })
      if (res.ok) {
        toast.success('Сообщение отправлено в Радар!')
        setText('')
        setPhotoUrl(null)
        setSelectedCategory(null)
      } else {
        toast.error('Ошибка при отправке')
      }
    } catch {
      toast.error('Ошибка сети')
    } finally {
      setSubmitting(false)
    }
  }

  // Загрузка данных дашборда
  const loadDashboard = async () => {
    setLoading(true)
    try {
      const params = new URLSearchParams()
      if (filterCategory !== 'all') params.set('category', filterCategory)
      if (filterStatus !== 'all') params.set('status', filterStatus)

      const [msgRes, statRes] = await Promise.all([
        fetch(`/api/radar?${params}`),
        fetch('/api/dashboard'),
      ])
      const msgData = await msgRes.json()
      const statData = await statRes.json()
      setMessages(msgData.messages || [])
      setStats(statData)
    } catch {
      toast.error('Ошибка загрузки дашборда')
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => {
    if (tab === 'dashboard') loadDashboard()
  }, [tab, filterCategory, filterStatus])

  // Обновить статус сообщения
  const updateStatus = async (id: string, status: string) => {
    try {
      await fetch('/api/radar', {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id, status }),
      })
      toast.success('Статус обновлён')
      loadDashboard()
    } catch {
      toast.error('Ошибка обновления')
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
          <Tabs value={tab} onValueChange={(v) => setTab(v as 'report' | 'dashboard')}>
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

            {/* Шаг 2: Данные сотрудника (если не сохранены) */}
            {!authorName && (
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
            )}

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

            {/* Информация о пользователе */}
            {authorName && (
              <div className="text-center text-xs text-slate-400">
                Отправляет: {authorName} · {ROLES.find(r => r.id === authorRole)?.label || authorRole} · {region}
                <button
                  onClick={() => { setAuthorName(''); setAuthorRole(''); setRegion('') }}
                  className="ml-2 text-blue-500 hover:underline"
                >
                  изменить
                </button>
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
