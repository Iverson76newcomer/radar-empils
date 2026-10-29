import { NextRequest, NextResponse } from 'next/server'
import { githubStore } from '@/lib/github-store'

// GET — статистика для дашборда
export async function GET(request: NextRequest) {
  try {
    const messages = await githubStore.getAll()

    // Статистика по категориям
    const byCategory = {
      competitors: messages.filter(m => m.category === 'competitors').length,
      product: messages.filter(m => m.category === 'product').length,
      idea: messages.filter(m => m.category === 'idea').length,
      problem: messages.filter(m => m.category === 'problem').length,
      trend: messages.filter(m => m.category === 'trend').length,
    }

    // Статистика по статусам
    const byStatus = {
      new: messages.filter(m => m.status === 'new').length,
      accepted: messages.filter(m => m.status === 'accepted').length,
      in_progress: messages.filter(m => m.status === 'in_progress').length,
      closed: messages.filter(m => m.status === 'closed').length,
      rejected: messages.filter(m => m.status === 'rejected').length,
    }

    // Топ активных сотрудников
    const authorMap = new Map<string, number>()
    messages.forEach(m => {
      authorMap.set(m.authorName, (authorMap.get(m.authorName) || 0) + 1)
    })
    const topAuthors = Array.from(authorMap.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5)
      .map(([name, count]) => ({ name, count }))

    // По регионам
    const regionMap = new Map<string, number>()
    messages.forEach(m => {
      regionMap.set(m.region, (regionMap.get(m.region) || 0) + 1)
    })
    const byRegion = Array.from(regionMap.entries())
      .sort((a, b) => b[1] - a[1])
      .map(([region, count]) => ({ region, count }))

    // Последние 7 дней — динамика по дням
    const last7Days: { date: string; count: number }[] = []
    for (let i = 6; i >= 0; i--) {
      const d = new Date()
      d.setDate(d.getDate() - i)
      const dateStr = d.toISOString().split('T')[0]
      const count = messages.filter(m => m.createdAt.split('T')[0] === dateStr).length
      last7Days.push({ date: dateStr, count })
    }

    return NextResponse.json({
      total: messages.length,
      byCategory,
      byStatus,
      topAuthors,
      byRegion,
      last7Days,
    }, { status: 200 })
  } catch (error) {
    console.error('Error fetching dashboard stats:', error)
    return NextResponse.json(
      { error: 'Ошибка при получении статистики: ' + (error as Error).message },
      { status: 500 }
    )
  }
}
