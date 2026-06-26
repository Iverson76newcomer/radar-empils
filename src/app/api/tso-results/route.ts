import { db } from '@/lib/db'
import { NextResponse } from 'next/server'

interface AnswerMap {
  [key: string]: string | string[] | number | null
}

interface ParsedResponse {
  id: string
  answers: AnswerMap
  createdAt: Date
}

function countFrequency(responses: ParsedResponse[], key: string): Record<string, number> {
  const freq: Record<string, number> = {}
  for (const r of responses) {
    const val = r.answers[key]
    if (Array.isArray(val)) {
      for (const v of val) {
        freq[v] = (freq[v] || 0) + 1
      }
    } else if (val !== null && val !== undefined) {
      freq[String(val)] = (freq[String(val)] || 0) + 1
    }
  }
  return freq
}

function calcAverage(responses: ParsedResponse[], key: string): number {
  const nums = responses
    .map((r) => Number(r.answers[key]))
    .filter((n) => !isNaN(n))
  if (nums.length === 0) return 0
  return nums.reduce((a, b) => a + b, 0) / nums.length
}

function calcDistribution(responses: ParsedResponse[], key: string): Record<string, number> {
  const dist: Record<string, number> = {}
  for (let i = 0; i <= 10; i++) {
    dist[String(i)] = 0
  }
  for (const r of responses) {
    const val = Number(r.answers[key])
    if (!isNaN(val) && val >= 0 && val <= 10) {
      dist[String(val)]++
    }
  }
  return dist
}

export async function GET() {
  try {
    const responses = await db.tsoResponse.findMany({
      orderBy: { createdAt: 'desc' },
    })

    const parsed: ParsedResponse[] = responses.map((r) => ({
      id: r.id,
      answers: JSON.parse(r.answers) as AnswerMap,
      createdAt: r.createdAt,
    }))

    const totalResponses = parsed.length

    // Single choice frequency counts
    const q1Freq = countFrequency(parsed, 'q1')
    const q2Freq = countFrequency(parsed, 'q2')
    const q3Freq = countFrequency(parsed, 'q3')
    const q4Freq = countFrequency(parsed, 'q4')
    const q7Freq = countFrequency(parsed, 'q7')
    const q8Freq = countFrequency(parsed, 'q8')
    const q11Freq = countFrequency(parsed, 'q11')
    const q15Freq = countFrequency(parsed, 'q15')

    // Multiple choice frequency counts
    const q5Freq = countFrequency(parsed, 'q5')
    const q9Freq = countFrequency(parsed, 'q9')
    const q12Freq = countFrequency(parsed, 'q12')
    const q13Freq = countFrequency(parsed, 'q13')
    const q14Freq = countFrequency(parsed, 'q14')
    const q16Freq = countFrequency(parsed, 'q16')

    // Scale questions
    const q10Avg = calcAverage(parsed, 'q10')
    const q10Dist = calcDistribution(parsed, 'q10')
    const q17Avg = calcAverage(parsed, 'q17')
    const q17Dist = calcDistribution(parsed, 'q17')

    // Free text
    const q6Texts = parsed
      .map((r) => r.answers['q6'] as string)
      .filter((t) => t && t.trim() !== '' && t.trim().toLowerCase() !== 'нет')

    return NextResponse.json({
      totalResponses,
      singleChoice: { q1: q1Freq, q2: q2Freq, q3: q3Freq, q4: q4Freq, q7: q7Freq, q8: q8Freq, q11: q11Freq, q15: q15Freq },
      multiChoice: { q5: q5Freq, q9: q9Freq, q12: q12Freq, q13: q13Freq, q14: q14Freq, q16: q16Freq },
      scale: {
        q10: { average: Math.round(q10Avg * 10) / 10, distribution: q10Dist },
        q17: { average: Math.round(q17Avg * 10) / 10, distribution: q17Dist },
      },
      freeText: { q6: q6Texts },
    })
  } catch (error) {
    console.error('Error aggregating results:', error)
    return NextResponse.json(
      { error: 'Failed to aggregate results' },
      { status: 500 }
    )
  }
}
