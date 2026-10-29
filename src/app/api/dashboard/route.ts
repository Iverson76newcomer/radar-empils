import { NextRequest, NextResponse } from 'next/server'
import { db } from '@/lib/db'

// GET — статистика для дашборда
export async function GET(request: NextRequest) {
  try {
    const messages = await db.radarMessage.findMany({
      orderBy: { createdAt: 'desc' },
    })

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

    return NextResponse.json({
      total: messages.length,
      byCategory,
      byStatus,
      topAuthors,
      byRegion,
    }, { status: 200 })
  } catch (error) {
    console.error('Error fetching dashboard stats:', error)
    return NextResponse.json(
      { error: 'Ошибка при получении статистики' },
      { status: 500 }
    )
  }
}
