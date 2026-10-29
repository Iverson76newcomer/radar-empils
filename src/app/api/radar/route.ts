import { NextRequest, NextResponse } from 'next/server'
import { db } from '@/lib/db'

// POST — создание нового сообщения от полевого сотрудника
export async function POST(request: NextRequest) {
  try {
    const body = await request.json()
    const { category, text, region, authorName, authorRole, photoUrl } = body

    if (!category || !text || !authorName) {
      return NextResponse.json(
        { error: 'Не заполнены обязательные поля: category, text, authorName' },
        { status: 400 }
      )
    }

    const message = await db.radarMessage.create({
      data: {
        category,
        text,
        region: region || 'Не указан',
        authorName,
        authorRole: authorRole || 'Не указана',
        photoUrl: photoUrl || null,
        status: 'new',
      },
    })

    return NextResponse.json({ success: true, message }, { status: 201 })
  } catch (error) {
    console.error('Error creating radar message:', error)
    return NextResponse.json(
      { error: 'Ошибка при создании сообщения' },
      { status: 500 }
    )
  }
}

// GET — получить все сообщения (для дашборда)
export async function GET(request: NextRequest) {
  try {
    const { searchParams } = new URL(request.url)
    const category = searchParams.get('category')
    const status = searchParams.get('status')
    const region = searchParams.get('region')
    const authorName = searchParams.get('authorName')

    const where: Record<string, string> = {}
    if (category) where.category = category
    if (status) where.status = status
    if (region) where.region = region
    if (authorName) where.authorName = authorName

    const messages = await db.radarMessage.findMany({
      where,
      orderBy: { createdAt: 'desc' },
      take: 200,
    })

    return NextResponse.json({ messages }, { status: 200 })
  } catch (error) {
    console.error('Error fetching radar messages:', error)
    return NextResponse.json(
      { error: 'Ошибка при получении сообщений' },
      { status: 500 }
    )
  }
}

// PATCH — обновить статус сообщения (для дашборда руководителя)
export async function PATCH(request: NextRequest) {
  try {
    const body = await request.json()
    const { id, status, managerNote } = body

    if (!id || !status) {
      return NextResponse.json(
        { error: 'Не указан id или status' },
        { status: 400 }
      )
    }

    const message = await db.radarMessage.update({
      where: { id },
      data: {
        status,
        ...(managerNote !== undefined && { managerNote }),
      },
    })

    return NextResponse.json({ success: true, message }, { status: 200 })
  } catch (error) {
    console.error('Error updating radar message:', error)
    return NextResponse.json(
      { error: 'Ошибка при обновлении сообщения' },
      { status: 500 }
    )
  }
}
