// Supabase через Yandex Cloud Function (прокси)
// Используется, потому что Supabase заблокирован в России,
// а functions.yandexcloud.net — нет

const PROXY_URL = process.env.NEXT_PUBLIC_YANDEX_FUNCTION_URL || ''
const SUPABASE_ANON_KEY = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY || ''

export interface Message {
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
  updatedAt: string
}

// Базовый fetch-обёртка с автоматическим добавлением API-ключей
async function supabaseFetch(path: string, options: RequestInit = {}): Promise<Response> {
  const url = `${PROXY_URL}?path=${encodeURIComponent(path)}`
  const headers: Record<string, string> = {
    'Content-Type': 'application/json',
    'apikey': SUPABASE_ANON_KEY,
    'Authorization': `Bearer ${SUPABASE_ANON_KEY}`,
    ...(options.headers as Record<string, string> || {}),
  }
  return fetch(url, { ...options, headers })
}

// Создать новое сообщение
export async function createMessage(data: {
  category: string
  text: string
  region: string
  authorName: string
  authorRole: string
  photoUrl: string | null
}): Promise<Message> {
  const id = `msg_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`
  const now = new Date().toISOString()

  const res = await supabaseFetch('/rest/v1/messages', {
    method: 'POST',
    headers: {
      'Prefer': 'return=representation',
    },
    body: JSON.stringify({
      id,
      category: data.category,
      text: data.text,
      region: data.region,
      authorName: data.authorName,
      authorRole: data.authorRole,
      photoUrl: data.photoUrl,
      status: 'new',
      managerNote: null,
      createdAt: now,
      updatedAt: now,
    }),
  })

  if (!res.ok) {
    const errText = await res.text()
    throw new Error(`Supabase insert error: ${res.status} ${errText}`)
  }

  const arr = await res.json()
  return arr[0] as Message
}

// Получить все сообщения с фильтрами
export async function getMessages(filters?: {
  category?: string
  status?: string
  region?: string
  authorName?: string
}): Promise<Message[]> {
  const supabaseQuery = new URLSearchParams()
  supabaseQuery.set('select', '*')
  supabaseQuery.set('order', 'createdAt.desc')
  supabaseQuery.set('limit', '200')

  if (filters?.category && filters.category !== 'all') {
    supabaseQuery.set('category', `eq.${filters.category}`)
  }
  if (filters?.status && filters.status !== 'all') {
    supabaseQuery.set('status', `eq.${filters.status}`)
  }

  const res = await supabaseFetch(`/rest/v1/messages?${supabaseQuery.toString()}`)
  if (!res.ok) {
    throw new Error(`Supabase select error: ${res.status} ${await res.text()}`)
  }
  return await res.json()
}

// Обновить статус сообщения
export async function updateMessageStatus(
  id: string,
  status: string,
  managerNote?: string
): Promise<Message | null> {
  const updates: any = { status, updatedAt: new Date().toISOString() }
  if (managerNote !== undefined) updates.managerNote = managerNote

  const supabaseQuery = new URLSearchParams()
  supabaseQuery.set('id', `eq.${id}`)

  const res = await supabaseFetch(`/rest/v1/messages?${supabaseQuery.toString()}`, {
    method: 'PATCH',
    headers: {
      'Prefer': 'return=representation',
    },
    body: JSON.stringify(updates),
  })

  if (!res.ok) {
    throw new Error(`Supabase update error: ${res.status} ${await res.text()}`)
  }

  const arr = await res.json()
  return arr[0] || null
}

// Удалить сообщение
export async function deleteMessage(id: string): Promise<boolean> {
  const supabaseQuery = new URLSearchParams()
  supabaseQuery.set('id', `eq.${id}`)

  const res = await supabaseFetch(`/rest/v1/messages?${supabaseQuery.toString()}`, {
    method: 'DELETE',
  })

  if (!res.ok) {
    throw new Error(`Supabase delete error: ${res.status} ${await res.text()}`)
  }
  return true
}

// Статистика для дашборда
export async function getDashboardStats() {
  const res = await supabaseFetch('/rest/v1/messages?select=*')
  if (!res.ok) {
    throw new Error(`Supabase stats error: ${res.status} ${await res.text()}`)
  }

  const all: Message[] = await res.json()

  const byCategory = {
    competitors: all.filter(m => m.category === 'competitors').length,
    product: all.filter(m => m.category === 'product').length,
    idea: all.filter(m => m.category === 'idea').length,
    problem: all.filter(m => m.category === 'problem').length,
    trend: all.filter(m => m.category === 'trend').length,
  }

  const byStatus = {
    new: all.filter(m => m.status === 'new').length,
    accepted: all.filter(m => m.status === 'accepted').length,
    in_progress: all.filter(m => m.status === 'in_progress').length,
    closed: all.filter(m => m.status === 'closed').length,
    rejected: all.filter(m => m.status === 'rejected').length,
  }

  const authorMap = new Map<string, number>()
  all.forEach(m => {
    authorMap.set(m.authorName, (authorMap.get(m.authorName) || 0) + 1)
  })
  const topAuthors = Array.from(authorMap.entries())
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5)
    .map(([name, count]) => ({ name, count }))

  const regionMap = new Map<string, number>()
  all.forEach(m => {
    regionMap.set(m.region, (regionMap.get(m.region) || 0) + 1)
  })
  const byRegion = Array.from(regionMap.entries())
    .sort((a, b) => b[1] - a[1])
    .map(([region, count]) => ({ region, count }))

  const last7Days: { date: string; count: number }[] = []
  for (let i = 6; i >= 0; i--) {
    const d = new Date()
    d.setDate(d.getDate() - i)
    const dateStr = d.toISOString().split('T')[0]
    const count = all.filter(m => {
      const mDate = new Date(m.createdAt).toISOString().split('T')[0]
      return mDate === dateStr
    }).length
    last7Days.push({ date: dateStr, count })
  }

  return {
    total: all.length,
    byCategory,
    byStatus,
    topAuthors,
    byRegion,
    last7Days,
  }
}
