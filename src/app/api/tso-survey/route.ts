import { db } from '@/lib/db'
import { NextRequest, NextResponse } from 'next/server'

export async function POST(request: NextRequest) {
  try {
    const body = await request.json()
    const { answers } = body

    if (!answers || typeof answers !== 'object') {
      return NextResponse.json(
        { error: 'Answers object is required' },
        { status: 400 }
      )
    }

    const response = await db.tsoResponse.create({
      data: {
        answers: JSON.stringify(answers),
      },
    })

    return NextResponse.json({ id: response.id, success: true }, { status: 201 })
  } catch (error) {
    console.error('Error saving survey response:', error)
    return NextResponse.json(
      { error: 'Failed to save response' },
      { status: 500 }
    )
  }
}

export async function GET() {
  try {
    const responses = await db.tsoResponse.findMany({
      orderBy: { createdAt: 'desc' },
    })

    const parsed = responses.map((r) => ({
      id: r.id,
      answers: JSON.parse(r.answers),
      createdAt: r.createdAt,
    }))

    return NextResponse.json(parsed)
  } catch (error) {
    console.error('Error fetching survey responses:', error)
    return NextResponse.json(
      { error: 'Failed to fetch responses' },
      { status: 500 }
    )
  }
}
