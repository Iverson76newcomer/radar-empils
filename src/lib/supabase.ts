import { createClient } from '@supabase/supabase-js'

// Используем NEXT_PUBLIC_ префикс, чтобы ключ был доступен в браузере
// anon key безопасно публиковать — он защищён RLS политиками
const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL || ''
const supabaseAnonKey = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY || ''

export const supabase = createClient(supabaseUrl, supabaseAnonKey, {
  auth: {
    persistSession: false,
  },
})

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

  const { data: row, error } = await supabase
    .from('messages')
    .insert({
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
    })
    .select()
    .single()

  if (error) throw new Error(`Supabase insert error: ${error.message}`)
  return row as unknown as Message
}

// Получить все сообщения с фильтрами
export async function getMessages(filters?: {
  category?: string
  status?: string
  region?: string
  authorName?: string
}): Promise<Message[]> {
  let query = supabase
    .from('messages')
    .select('*')
    .order('createdAt', { ascending: false })
    .limit(200)

  if (filters?.category && filters.category !== 'all') {
    query = query.eq('category', filters.category)
  }
  if (filters?.status && filters.status !== 'all') {
    query = query.eq('status', filters.status)
  }
  if (filters?.region) {
    query = query.eq('region', filters.region)
  }
  if (filters?.authorName) {
    query = query.eq('authorName', filters.authorName)
  }

  const { data, error } = await query
  if (error) throw new Error(`Supabase select error: ${error.message}`)
  return (data || []) as unknown as Message[]
}

// Обновить статус сообщения
export async function updateMessageStatus(
  id: string,
  status: string,
  managerNote?: string
): Promise<Message | null> {
  const updates: any = { status, updatedAt: new Date().toISOString() }
  if (managerNote !== undefined) updates.managerNote = managerNote

  const { data, error } = await supabase
    .from('messages')
    .update(updates)
    .eq('id', id)
    .select()
    .single()

  if (error) throw new Error(`Supabase update error: ${error.message}`)
  return data as unknown as Message
}

// Удалить сообщение
export async function deleteMessage(id: string): Promise<boolean> {
  const { error } = await supabase.from('messages').delete().eq('id', id)
  if (error) throw new Error(`Supabase delete error: ${error.message}`)
  return true
}

// Статистика для дашборда
export async function getDashboardStats() {
  const { data: messages, error } = await supabase.from('messages').select('*')
  if (error) throw new Error(`Supabase stats error: ${error.message}`)

  const all = messages || []

  const byCategory = {
    competitors: all.filter((m: any) => m.category === 'competitors').length,
    product: all.filter((m: any) => m.category === 'product').length,
    idea: all.filter((m: any) => m.category === 'idea').length,
    problem: all.filter((m: any) => m.category === 'problem').length,
    trend: all.filter((m: any) => m.category === 'trend').length,
  }

  const byStatus = {
    new: all.filter((m: any) => m.status === 'new').length,
    accepted: all.filter((m: any) => m.status === 'accepted').length,
    in_progress: all.filter((m: any) => m.status === 'in_progress').length,
    closed: all.filter((m: any) => m.status === 'closed').length,
    rejected: all.filter((m: any) => m.status === 'rejected').length,
  }

  const authorMap = new Map<string, number>()
  all.forEach((m: any) => {
    authorMap.set(m.authorName, (authorMap.get(m.authorName) || 0) + 1)
  })
  const topAuthors = Array.from(authorMap.entries())
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5)
    .map(([name, count]) => ({ name, count }))

  const regionMap = new Map<string, number>()
  all.forEach((m: any) => {
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
    const count = all.filter((m: any) => {
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
