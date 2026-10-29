import { NextRequest, NextResponse } from 'next/server'
import { githubStore } from '@/lib/github-store'

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

    // Валидация категории
    const validCategories = ['competitors', 'product', 'idea', 'problem', 'trend']
    if (!validCategories.includes(category)) {
      return NextResponse.json(
        { error: `Неверная категория: ${category}. Допустимые: ${validCategories.join(', ')}` },
        { status: 400 }
      )
    }

    // Генерируем уникальный ID
    const id = `msg_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`

    const message = {
      id,
      category,
      text,
      region: region || 'Не указан',
      authorName,
      authorRole: authorRole || 'Не указана',
      photoUrl: photoUrl || null,
      status: 'new',
      managerNote: null,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    }

    // Сохраняем через GitHub API
    await githubStore.add(message)

    return NextResponse.json({ success: true, message }, { status: 201 })
  } catch (error) {
    console.error('Error creating radar message:', error)
    return NextResponse.json(
      { error: 'Ошибка при создании сообщения: ' + (error as Error).message },
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

    const allMessages = await githubStore.getAll()

    // Сортируем от новых к старым
    const sorted = [...allMessages].sort(
      (a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime()
    )

    // Применяем фильтры
    const filtered = sorted.filter(m => {
      if (category && m.category !== category) return false
      if (status && m.status !== status) return false
      if (region && m.region !== region) return false
      if (authorName && m.authorName !== authorName) return false
      return true
    })

    // Возвращаем последние 200
    const result = filtered.slice(0, 200)

    return NextResponse.json({ messages: result }, { status: 200 })
  } catch (error) {
    console.error('Error fetching radar messages:', error)
    return NextResponse.json(
      { error: 'Ошибка при получении сообщений: ' + (error as Error).message },
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

    const updated = await githubStore.update(id, {
      status,
      ...(managerNote !== undefined && { managerNote }),
    })

    if (!updated) {
      return NextResponse.json(
        { error: 'Сообщение не найдено' },
        { status: 404 }
      )
    }

    return NextResponse.json({ success: true, message: updated }, { status: 200 })
  } catch (error) {
    console.error('Error updating radar message:', error)
    return NextResponse.json(
      { error: 'Ошибка при обновлении сообщения: ' + (error as Error).message },
      { status: 500 }
    )
  }
}

// DELETE — удалить сообщение
export async function DELETE(request: NextRequest) {
  try {
    const { searchParams } = new URL(request.url)
    const id = searchParams.get('id')

    if (!id) {
      return NextResponse.json(
        { error: 'Не указан id' },
        { status: 400 }
      )
    }

    const deleted = await githubStore.delete(id)
    if (!deleted) {
      return NextResponse.json(
        { error: 'Сообщение не найдено' },
        { status: 404 }
      )
    }

    return NextResponse.json({ success: true }, { status: 200 })
  } catch (error) {
    console.error('Error deleting radar message:', error)
    return NextResponse.json(
      { error: 'Ошибка при удалении сообщения: ' + (error as Error).message },
      { status: 500 }
    )
  }
}
