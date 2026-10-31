// Экспорт данных Радар Эмпилс в Excel (6 листов) и PDF (резюме)
// Браузерная генерация — не тратит кредиты Yandex Cloud

import ExcelJS from 'exceljs'
import { jsPDF } from 'jspdf'
import autoTable from 'jspdf-autotable'
import type { Message } from './supabase'

// Утилита: форматирование даты
const formatDate = (d: string) =>
  new Date(d).toLocaleString('ru-RU', {
    day: '2-digit', month: '2-digit', year: 'numeric',
    hour: '2-digit', minute: '2-digit',
  })

// Утилита: триггер скачивания файла
function triggerDownload(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  setTimeout(() => URL.revokeObjectURL(url), 1000)
}

// Загрузка TTF-шрифта Roboto с поддержкой кириллицы
// jsPDF по умолчанию использует Helvetica — она НЕ поддерживает кириллицу
let fontCache: { regular?: string; bold?: string } = {}

async function loadRobotoFont(variant: 'regular' | 'bold'): Promise<string> {
  if (fontCache[variant]) return fontCache[variant]!

  const url = variant === 'bold'
    ? '/fonts/Roboto-Bold.ttf'
    : '/fonts/Roboto-Regular.ttf'

  const res = await fetch(url)
  if (!res.ok) {
    throw new Error(`Не удалось загрузить шрифт ${variant}: HTTP ${res.status}`)
  }
  const buffer = await res.arrayBuffer()
  // Конвертируем в base64 для jsPDF VFS
  const bytes = new Uint8Array(buffer)
  let binary = ''
  const chunk = 0x8000
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + chunk)))
  }
  const base64 = btoa(binary)
  fontCache[variant] = base64
  return base64
}

// Создание jsPDF с встроенным шрифтом Roboto (с кириллицей)
async function createPdfWithRoboto(): Promise<jsPDF> {
  const [regularBase64, boldBase64] = await Promise.all([
    loadRobotoFont('regular'),
    loadRobotoFont('bold'),
  ])

  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' })

  // Регистрируем шрифты в VFS jsPDF
  doc.addFileToVFS('Roboto-Regular.ttf', regularBase64)
  doc.addFileToVFS('Roboto-Bold.ttf', boldBase64)
  doc.addFont('Roboto-Regular.ttf', 'Roboto', 'normal')
  doc.addFont('Roboto-Bold.ttf', 'Roboto', 'bold')

  // Устанавливаем Roboto как шрифт по умолчанию
  doc.setFont('Roboto', 'normal')

  return doc
}

// Категории (должны совпадать с page.tsx)
const CATEGORIES = [
  { id: 'competitors', label: 'Конкуренты', color: 'FFDC2626' },
  { id: 'product', label: 'Запрос продукта', color: 'FF2563EB' },
  { id: 'idea', label: 'Идея оформления', color: 'FF16A34A' },
  { id: 'problem', label: 'Проблема', color: 'FFCA8A04' },
  { id: 'trend', label: 'Тренд рынка', color: 'FF9333EA' },
]

const STATUSES = [
  { id: 'new', label: 'Новое', color: 'FFDBEAFE', textColor: 'FF1E40AF' },
  { id: 'accepted', label: 'Принято', color: 'FFD1FAE5', textColor: 'FF166534' },
  { id: 'in_progress', label: 'В работе', color: 'FFFEF9C3', textColor: 'FF854D0E' },
  { id: 'closed', label: 'Закрыто', color: 'FFF1F5F9', textColor: 'FF475569' },
  { id: 'rejected', label: 'Отклонено', color: 'FFFEE2E2', textColor: 'FF991B1B' },
]

const ROLES: Record<string, string> = {
  RM: 'Региональный менеджер',
  ETP: 'Эксклюзивный ТП',
  merchandiser: 'Мерчендайзер',
}

// ============================================================
// ЭКСПОРТ В EXCEL — 6 листов
// ============================================================
export async function exportToExcel(messages: Message[]) {
  const wb = new ExcelJS.Workbook()
  wb.creator = 'Эмпилс Радар'
  wb.created = new Date()

  // -------- ЛИСТ 1: СООБЩЕНИЯ (детальный список) --------
  const ws1 = wb.addWorksheet('Сообщения', {
    views: [{ state: 'frozen', ySplit: 4, xSplit: 0 }],
  })

  // Шапка-резюме
  ws1.mergeCells('A1:K1')
  const titleCell = ws1.getCell('A1')
  titleCell.value = 'ЭМПИЛС · Радар полевых наблюдений — Детальный отчёт'
  titleCell.font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } }
  titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B3D91' } }
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' }
  ws1.getRow(1).height = 28

  ws1.mergeCells('A2:K2')
  const periodCell = ws1.getCell('A2')
  const firstDate = messages.length > 0 ? formatDate(messages[messages.length - 1].createdAt) : '—'
  const lastDate = messages.length > 0 ? formatDate(messages[0].createdAt) : '—'
  periodCell.value = `Период: ${firstDate} — ${lastDate}   ·   Всего сообщений: ${messages.length}`
  periodCell.font = { size: 11, color: { argb: 'FF475569' } }
  periodCell.alignment = { horizontal: 'center', vertical: 'middle' }
  ws1.getRow(2).height = 18

  ws1.getRow(3).height = 8 // пустая строка-разделитель

  // Заголовки колонок
  const headers = [
    'ID', 'Дата создания', 'Категория', 'Текст сообщения',
    'Автор', 'Роль', 'Регион', 'Статус',
    'Комментарий РД', 'Дата обновления', 'Фото',
  ]
  headers.forEach((h, i) => {
    const cell = ws1.getCell(4, i + 1)
    cell.value = h
    cell.font = { bold: true, size: 11, color: { argb: 'FFFFFFFF' } }
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E5BC6' } }
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF0B3D91' } },
      bottom: { style: 'thin', color: { argb: 'FF0B3D91' } },
      left: { style: 'thin', color: { argb: 'FF0B3D91' } },
      right: { style: 'thin', color: { argb: 'FF0B3D91' } },
    }
  })
  ws1.getRow(4).height = 28

  // Данные
  messages.forEach((msg, idx) => {
    const row = ws1.getRow(5 + idx)
    const cat = CATEGORIES.find(c => c.id === msg.category)
    const stat = STATUSES.find(s => s.id === msg.status)

    row.getCell(1).value = msg.id
    row.getCell(2).value = formatDate(msg.createdAt)
    row.getCell(3).value = cat?.label || msg.category
    row.getCell(4).value = msg.text
    row.getCell(5).value = msg.authorName
    row.getCell(6).value = ROLES[msg.authorRole] || msg.authorRole
    row.getCell(7).value = msg.region
    row.getCell(8).value = stat?.label || msg.status
    row.getCell(9).value = msg.managerNote || ''
    row.getCell(10).value = formatDate(msg.updatedAt)
    row.getCell(11).value = msg.photoUrl ? '✓ есть' : '—'

    // Цвет статуса
    if (stat) {
      row.getCell(8).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: stat.color } }
      row.getCell(8).font = { bold: true, color: { argb: stat.textColor } }
      row.getCell(8).alignment = { horizontal: 'center' }
    }

    // Цвет категории
    if (cat) {
      row.getCell(3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: cat.color } }
      row.getCell(3).font = { bold: true, color: { argb: 'FFFFFFFF' } }
      row.getCell(3).alignment = { horizontal: 'center' }
    }

    // Чередование строк
    if (idx % 2 === 1) {
      for (let i = 1; i <= 11; i++) {
        if (i !== 3 && i !== 8) { // не перекрашиваем колонки с категорией/статусом
          const cell = row.getCell(i)
          if (!cell.fill || cell.fill.type !== 'pattern') {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } }
          }
        }
      }
    }

    // Wrap текста
    row.getCell(4).alignment = { wrapText: true, vertical: 'top' }
    row.getCell(9).alignment = { wrapText: true, vertical: 'top' }

    // Border
    for (let i = 1; i <= 11; i++) {
      const cell = row.getCell(i)
      cell.border = {
        top: { style: 'hair', color: { argb: 'FFE2E8F0' } },
        bottom: { style: 'hair', color: { argb: 'FFE2E8F0' } },
        left: { style: 'hair', color: { argb: 'FFE2E8F0' } },
        right: { style: 'hair', color: { argb: 'FFE2E8F0' } },
      }
    }
  })

  // Ширина колонок
  const widths = [22, 16, 16, 50, 18, 22, 22, 12, 30, 16, 10]
  widths.forEach((w, i) => ws1.getColumn(i + 1).width = w)

  // -------- ЛИСТ 2: СТАТИСТИКА ПО КАТЕГОРИЯМ --------
  const ws2 = wb.addWorksheet('По категориям')
  ws2.mergeCells('A1:D1')
  ws2.getCell('A1').value = 'Статистика по категориям'
  ws2.getCell('A1').font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } }
  ws2.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B3D91' } }
  ws2.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' }
  ws2.getRow(1).height = 24

  const catHeaders = ['Категория', 'Кол-во', '% от общего', 'График']
  catHeaders.forEach((h, i) => {
    const c = ws2.getCell(2, i + 1)
    c.value = h
    c.font = { bold: true, color: { argb: 'FFFFFFFF' } }
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E5BC6' } }
    c.alignment = { horizontal: 'center' }
  })

  const total = messages.length
  CATEGORIES.forEach((cat, idx) => {
    const count = messages.filter(m => m.category === cat.id).length
    const pct = total > 0 ? (count / total * 100) : 0
    const row = ws2.getRow(3 + idx)
    row.getCell(1).value = cat.label
    row.getCell(2).value = count
    row.getCell(3).value = `${pct.toFixed(1)}%`
    // Простой ASCII-график
    row.getCell(4).value = '█'.repeat(Math.round(pct / 2))
    // Цвета
    row.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: cat.color } }
    row.getCell(1).font = { bold: true, color: { argb: 'FFFFFFFF' } }
    row.getCell(4).font = { color: { argb: cat.color } }
  })

  // Итого
  const totalRow = ws2.getRow(3 + CATEGORIES.length)
  totalRow.getCell(1).value = 'ИТОГО'
  totalRow.getCell(2).value = total
  totalRow.getCell(3).value = '100%'
  totalRow.font = { bold: true }
  totalRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } }

  ws2.getColumn(1).width = 20
  ws2.getColumn(2).width = 12
  ws2.getColumn(3).width = 14
  ws2.getColumn(4).width = 50

  // -------- ЛИСТ 3: СТАТИСТИКА ПО СТАТУСАМ --------
  const ws3 = wb.addWorksheet('По статусам')
  ws3.mergeCells('A1:D1')
  ws3.getCell('A1').value = 'Статусы обработки (воронка)'
  ws3.getCell('A1').font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } }
  ws3.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B3D91' } }
  ws3.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' }
  ws3.getRow(1).height = 24

  const stHeaders = ['Статус', 'Кол-во', '% от общего', 'График']
  stHeaders.forEach((h, i) => {
    const c = ws3.getCell(2, i + 1)
    c.value = h
    c.font = { bold: true, color: { argb: 'FFFFFFFF' } }
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E5BC6' } }
    c.alignment = { horizontal: 'center' }
  })

  STATUSES.forEach((st, idx) => {
    const count = messages.filter(m => m.status === st.id).length
    const pct = total > 0 ? (count / total * 100) : 0
    const row = ws3.getRow(3 + idx)
    row.getCell(1).value = st.label
    row.getCell(2).value = count
    row.getCell(3).value = `${pct.toFixed(1)}%`
    row.getCell(4).value = '█'.repeat(Math.round(pct / 2))
    row.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: st.color } }
    row.getCell(1).font = { bold: true, color: { argb: st.textColor } }
    row.getCell(4).font = { color: { argb: st.textColor } }
  })

  const totalRow3 = ws3.getRow(3 + STATUSES.length)
  totalRow3.getCell(1).value = 'ИТОГО'
  totalRow3.getCell(2).value = total
  totalRow3.getCell(3).value = '100%'
  totalRow3.font = { bold: true }
  totalRow3.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } }

  ws3.getColumn(1).width = 20
  ws3.getColumn(2).width = 12
  ws3.getColumn(3).width = 14
  ws3.getColumn(4).width = 50

  // -------- ЛИСТ 4: ПО РЕГИОНАМ --------
  const ws4 = wb.addWorksheet('По регионам')
  ws4.mergeCells('A1:F1')
  ws4.getCell('A1').value = 'География сообщений'
  ws4.getCell('A1').font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } }
  ws4.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B3D91' } }
  ws4.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' }
  ws4.getRow(1).height = 24

  const regHeaders = ['Регион', 'Всего', '%', 'Новые', 'В работе', 'Закрыто']
  regHeaders.forEach((h, i) => {
    const c = ws4.getCell(2, i + 1)
    c.value = h
    c.font = { bold: true, color: { argb: 'FFFFFFFF' } }
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E5BC6' } }
    c.alignment = { horizontal: 'center' }
  })

  const regionMap = new Map<string, { total: number, new: number, in_progress: number, closed: number }>()
  messages.forEach(m => {
    const r = m.region
    if (!regionMap.has(r)) regionMap.set(r, { total: 0, new: 0, in_progress: 0, closed: 0 })
    const obj = regionMap.get(r)!
    obj.total++
    if (m.status === 'new') obj.new++
    if (m.status === 'in_progress') obj.in_progress++
    if (m.status === 'closed') obj.closed++
  })

  const sortedRegions = Array.from(regionMap.entries()).sort((a, b) => b[1].total - a[1].total)
  sortedRegions.forEach(([region, data], idx) => {
    const pct = total > 0 ? (data.total / total * 100) : 0
    const row = ws4.getRow(3 + idx)
    row.getCell(1).value = region
    row.getCell(2).value = data.total
    row.getCell(3).value = `${pct.toFixed(1)}%`
    row.getCell(4).value = data.new
    row.getCell(5).value = data.in_progress
    row.getCell(6).value = data.closed
    if (idx % 2 === 1) {
      for (let i = 1; i <= 6; i++) {
        row.getCell(i).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } }
      }
    }
  })

  ws4.getColumn(1).width = 28
  ws4.getColumn(2).width = 10
  ws4.getColumn(3).width = 10
  ws4.getColumn(4).width = 10
  ws4.getColumn(5).width = 12
  ws4.getColumn(6).width = 12

  // -------- ЛИСТ 5: ТОП СОТРУДНИКОВ --------
  const ws5 = wb.addWorksheet('ТОП сотрудников')
  ws5.mergeCells('A1:G1')
  ws5.getCell('A1').value = 'ТОП активных сотрудников'
  ws5.getCell('A1').font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } }
  ws5.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B3D91' } }
  ws5.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' }
  ws5.getRow(1).height = 24

  const topHeaders = ['Место', 'Сотрудник', 'Роль', 'Регион', 'Всего сообщений', 'Закрыто', '% закрытия']
  topHeaders.forEach((h, i) => {
    const c = ws5.getCell(2, i + 1)
    c.value = h
    c.font = { bold: true, color: { argb: 'FFFFFFFF' } }
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E5BC6' } }
    c.alignment = { horizontal: 'center' }
  })

  const authorMap = new Map<string, { role: string, region: string, total: number, closed: number }>()
  messages.forEach(m => {
    const key = `${m.authorName}|${m.region}`
    if (!authorMap.has(key)) {
      authorMap.set(key, { role: m.authorRole, region: m.region, total: 0, closed: 0 })
    }
    const obj = authorMap.get(key)!
    obj.total++
    if (m.status === 'closed') obj.closed++
  })

  const sortedAuthors = Array.from(authorMap.entries())
    .sort((a, b) => b[1].total - a[1].total)
    .slice(0, 20) // ТОП-20

  sortedAuthors.forEach(([key, data], idx) => {
    const [name] = key.split('|')
    const closeRate = data.total > 0 ? (data.closed / data.total * 100) : 0
    const row = ws5.getRow(3 + idx)
    row.getCell(1).value = idx + 1
    row.getCell(2).value = name
    row.getCell(3).value = ROLES[data.role] || data.role
    row.getCell(4).value = data.region
    row.getCell(5).value = data.total
    row.getCell(6).value = data.closed
    row.getCell(7).value = `${closeRate.toFixed(0)}%`

    // ТОП-3 — золото/серебро/бронза
    if (idx === 0) {
      row.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } }
      row.getCell(1).font = { bold: true, color: { argb: 'FF92400E' } }
    } else if (idx === 1) {
      row.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE5E7EB' } }
      row.getCell(1).font = { bold: true, color: { argb: 'FF475569' } }
    } else if (idx === 2) {
      row.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFED7AA' } }
      row.getCell(1).font = { bold: true, color: { argb: 'FF9A3412' } }
    }
  })

  ws5.getColumn(1).width = 8
  ws5.getColumn(2).width = 22
  ws5.getColumn(3).width = 22
  ws5.getColumn(4).width = 22
  ws5.getColumn(5).width = 16
  ws5.getColumn(6).width = 12
  ws5.getColumn(7).width = 14

  // -------- ЛИСТ 6: ДИНАМИКА ПО ДНЯМ --------
  const ws6 = wb.addWorksheet('Динамика по дням')
  ws6.mergeCells('A1:D1')
  ws6.getCell('A1').value = 'Динамика сообщений (последние 30 дней)'
  ws6.getCell('A1').font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } }
  ws6.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B3D91' } }
  ws6.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' }
  ws6.getRow(1).height = 24

  const dynHeaders = ['Дата', 'День недели', 'Кол-во сообщений', 'График']
  dynHeaders.forEach((h, i) => {
    const c = ws6.getCell(2, i + 1)
    c.value = h
    c.font = { bold: true, color: { argb: 'FFFFFFFF' } }
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E5BC6' } }
    c.alignment = { horizontal: 'center' }
  })

  // Группировка по дате
  const dateMap = new Map<string, number>()
  messages.forEach(m => {
    const dateStr = new Date(m.createdAt).toISOString().split('T')[0]
    dateMap.set(dateStr, (dateMap.get(dateStr) || 0) + 1)
  })

  // Последние 30 дней
  const days: { date: string, count: number }[] = []
  for (let i = 29; i >= 0; i--) {
    const d = new Date()
    d.setDate(d.getDate() - i)
    const dateStr = d.toISOString().split('T')[0]
    days.push({ date: dateStr, count: dateMap.get(dateStr) || 0 })
  }

  const maxCount = Math.max(...days.map(d => d.count), 1)
  days.forEach((day, idx) => {
    const date = new Date(day.date)
    const dayName = date.toLocaleDateString('ru-RU', { weekday: 'long' })
    const row = ws6.getRow(3 + idx)
    row.getCell(1).value = date.toLocaleDateString('ru-RU', { day: '2-digit', month: '2-digit', year: 'numeric' })
    row.getCell(2).value = dayName.charAt(0).toUpperCase() + dayName.slice(1)
    row.getCell(3).value = day.count
    // ASCII-график
    const barLen = Math.round((day.count / maxCount) * 30)
    row.getCell(4).value = '█'.repeat(barLen)
    row.getCell(4).font = { color: { argb: 'FF0B3D91' } }

    // Выходные — серым
    const wd = date.getDay()
    if (wd === 0 || wd === 6) {
      for (let i = 1; i <= 4; i++) {
        row.getCell(i).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } }
        row.getCell(i).font = { ...row.getCell(i).font, italic: true }
      }
    }
  })

  ws6.getColumn(1).width = 14
  ws6.getColumn(2).width = 14
  ws6.getColumn(3).width = 18
  ws6.getColumn(4).width = 35

  // -------- Сохранение --------
  const buffer = await wb.xlsx.writeBuffer()
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
  const dateStr = new Date().toISOString().split('T')[0]
  triggerDownload(blob, `Эмпилс_Радар_отчёт_${dateStr}.xlsx`)
}

// ============================================================
// ЭКСПОРТ В PDF — краткий отчёт-резюме (1-2 страницы)
// ============================================================
export async function exportToPDF(messages: Message[]) {
  const doc = await createPdfWithRoboto()
  const pageWidth = doc.internal.pageSize.getWidth()
  const pageHeight = doc.internal.pageSize.getHeight()
  const margin = 15
  let y = margin

  // -------- ШАПКА --------
  // Синяя полоса с заголовком
  doc.setFillColor(11, 61, 145) // #0B3D91
  doc.rect(0, 0, pageWidth, 30, 'F')
  doc.setTextColor(255, 255, 255)
  doc.setFont('Roboto', 'bold')
  doc.setFontSize(18)
  doc.text('EMPILS', margin, 14)
  doc.setFontSize(11)
  doc.setFont('Roboto', 'normal')
  doc.text('Радар полевых наблюдений — Отчёт', margin, 22)

  y = 40
  doc.setTextColor(15, 23, 42)

  // -------- KPI-БЛОК --------
  const total = messages.length
  const newCount = messages.filter(m => m.status === 'new').length
  const inProgress = messages.filter(m => m.status === 'in_progress').length
  const closed = messages.filter(m => m.status === 'closed').length
  const rejected = messages.filter(m => m.status === 'rejected').length

  doc.setFont('Roboto', 'bold')
  doc.setFontSize(13)
  doc.text('Сводка', margin, y)
  y += 5

  // Линия
  doc.setDrawColor(255, 107, 53) // #FF6B35
  doc.setLineWidth(0.8)
  doc.line(margin, y, margin + 30, y)
  y += 8

  // KPI карточки — 2x2
  const kpiData = [
    { label: 'ВСЕГО СООБЩЕНИЙ', value: total, color: [11, 61, 145] },
    { label: 'НОВЫЕ', value: newCount, color: [30, 64, 175] },
    { label: 'В РАБОТЕ', value: inProgress, color: [202, 138, 4] },
    { label: 'ЗАКРЫТО', value: closed, color: [22, 101, 52] },
  ]

  const cardW = (pageWidth - margin * 2 - 5) / 2
  const cardH = 22

  kpiData.forEach((kpi, idx) => {
    const col = idx % 2
    const row = Math.floor(idx / 2)
    const x = margin + col * (cardW + 5)
    const cy = y + row * (cardH + 4)

    // Фон карточки
    doc.setFillColor(248, 250, 252)
    doc.roundedRect(x, cy, cardW, cardH, 2, 2, 'F')

    // Левая цветная полоска
    doc.setFillColor(kpi.color[0], kpi.color[1], kpi.color[2])
    doc.roundedRect(x, cy, 2, cardH, 1, 1, 'F')

    // Лейбл
    doc.setFontSize(7)
    doc.setFont('Roboto', 'bold')
    doc.setTextColor(100, 116, 139)
    doc.text(kpi.label, x + 6, cy + 8)

    // Значение
    doc.setFontSize(20)
    doc.setTextColor(kpi.color[0], kpi.color[1], kpi.color[2])
    doc.text(String(kpi.value), x + 6, cy + 18)
  })

  y += cardH * 2 + 12

  // -------- РАСПРЕДЕЛЕНИЕ ПО КАТЕГОРИЯМ --------
  doc.setFontSize(13)
  doc.setFont('Roboto', 'bold')
  doc.setTextColor(15, 23, 42)
  doc.text('По категориям', margin, y)
  y += 4
  doc.setDrawColor(255, 107, 53)
  doc.setLineWidth(0.8)
  doc.line(margin, y, margin + 30, y)
  y += 4

  const catData = CATEGORIES.map(cat => ({
    label: cat.label,
    count: messages.filter(m => m.category === cat.id).length,
    color: cat.color,
  }))

  autoTable(doc, {
    startY: y,
    head: [['Категория', 'Кол-во', '%']],
    body: catData.map(c => [
      c.label,
      String(c.count),
      total > 0 ? `${(c.count / total * 100).toFixed(1)}%` : '0%',
    ]),
    styles: { font: 'Roboto', fontSize: 9 },
    headStyles: {
      fillColor: [30, 91, 198],
      textColor: 255,
      fontStyle: 'bold',
      fontSize: 9,
    },
    bodyStyles: { fontSize: 9, textColor: [15, 23, 42] },
    columnStyles: {
      0: { cellWidth: 70 },
      1: { cellWidth: 30, halign: 'center' },
      2: { cellWidth: 30, halign: 'center' },
    },
    margin: { left: margin, right: margin },
  })

  y = (doc as any).lastAutoTable.finalY + 8

  // -------- СТАТУСЫ --------
  doc.setFontSize(13)
  doc.setFont('Roboto', 'bold')
  doc.text('По статусам', margin, y)
  y += 4
  doc.setDrawColor(255, 107, 53)
  doc.setLineWidth(0.8)
  doc.line(margin, y, margin + 30, y)
  y += 4

  const statData = STATUSES.map(st => ({
    label: st.label,
    count: messages.filter(m => m.status === st.id).length,
  }))

  autoTable(doc, {
    startY: y,
    head: [['Статус', 'Кол-во', '%']],
    body: statData.map(s => [
      s.label,
      String(s.count),
      total > 0 ? `${(s.count / total * 100).toFixed(1)}%` : '0%',
    ]),
    styles: { font: 'Roboto', fontSize: 9 },
    headStyles: {
      fillColor: [30, 91, 198],
      textColor: 255,
      fontStyle: 'bold',
      fontSize: 9,
    },
    bodyStyles: { fontSize: 9, textColor: [15, 23, 42] },
    columnStyles: {
      0: { cellWidth: 70 },
      1: { cellWidth: 30, halign: 'center' },
      2: { cellWidth: 30, halign: 'center' },
    },
    margin: { left: margin, right: margin },
  })

  y = (doc as any).lastAutoTable.finalY + 8

  // -------- ТОП-5 РЕГИОНОВ --------
  if (y > pageHeight - 60) {
    doc.addPage()
    y = margin
  }

  doc.setFontSize(13)
  doc.setFont('Roboto', 'bold')
  doc.text('ТОП-5 регионов', margin, y)
  y += 4
  doc.setDrawColor(255, 107, 53)
  doc.setLineWidth(0.8)
  doc.line(margin, y, margin + 30, y)
  y += 4

  const regionMap = new Map<string, number>()
  messages.forEach(m => {
    regionMap.set(m.region, (regionMap.get(m.region) || 0) + 1)
  })
  const topRegions = Array.from(regionMap.entries())
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5)

  autoTable(doc, {
    startY: y,
    head: [['Регион', 'Кол-во', '%']],
    body: topRegions.map(([r, c]) => [
      r,
      String(c),
      total > 0 ? `${(c / total * 100).toFixed(1)}%` : '0%',
    ]),
    styles: { font: 'Roboto', fontSize: 9 },
    headStyles: {
      fillColor: [30, 91, 198],
      textColor: 255,
      fontStyle: 'bold',
      fontSize: 9,
    },
    bodyStyles: { fontSize: 9, textColor: [15, 23, 42] },
    columnStyles: {
      0: { cellWidth: 100 },
      1: { cellWidth: 20, halign: 'center' },
      2: { cellWidth: 20, halign: 'center' },
    },
    margin: { left: margin, right: margin },
  })

  y = (doc as any).lastAutoTable.finalY + 8

  // -------- ТОП-5 СОТРУДНИКОВ --------
  if (y > pageHeight - 60) {
    doc.addPage()
    y = margin
  }

  doc.setFontSize(13)
  doc.setFont('Roboto', 'bold')
  doc.text('ТОП-5 сотрудников', margin, y)
  y += 4
  doc.setDrawColor(255, 107, 53)
  doc.setLineWidth(0.8)
  doc.line(margin, y, margin + 30, y)
  y += 4

  const authorMap = new Map<string, { role: string, count: number, closed: number }>()
  messages.forEach(m => {
    if (!authorMap.has(m.authorName)) {
      authorMap.set(m.authorName, { role: m.authorRole, count: 0, closed: 0 })
    }
    const obj = authorMap.get(m.authorName)!
    obj.count++
    if (m.status === 'closed') obj.closed++
  })
  const topAuthors = Array.from(authorMap.entries())
    .sort((a, b) => b[1].count - a[1].count)
    .slice(0, 5)

  autoTable(doc, {
    startY: y,
    head: [['Сотрудник', 'Роль', 'Сообщений', 'Закрыто']],
    body: topAuthors.map(([name, d]) => [
      name,
      ROLES[d.role] || d.role,
      String(d.count),
      String(d.closed),
    ]),
    styles: { font: 'Roboto', fontSize: 9 },
    headStyles: {
      fillColor: [30, 91, 198],
      textColor: 255,
      fontStyle: 'bold',
      fontSize: 9,
    },
    bodyStyles: { fontSize: 9, textColor: [15, 23, 42] },
    columnStyles: {
      0: { cellWidth: 60 },
      1: { cellWidth: 55 },
      2: { cellWidth: 22, halign: 'center' },
      3: { cellWidth: 22, halign: 'center' },
    },
    margin: { left: margin, right: margin },
  })

  y = (doc as any).lastAutoTable.finalY + 10

  // -------- ФУТЕР --------
  if (y > pageHeight - 25) {
    doc.addPage()
    y = margin
  }

  doc.setDrawColor(226, 232, 240)
  doc.setLineWidth(0.3)
  doc.line(margin, y, pageWidth - margin, y)
  y += 5

  doc.setFontSize(8)
  doc.setFont('Roboto', 'normal')
  doc.setTextColor(100, 116, 139)
  const dateStr = new Date().toLocaleString('ru-RU', {
    day: '2-digit', month: '2-digit', year: 'numeric',
    hour: '2-digit', minute: '2-digit',
  })
  doc.text(`Отчёт сформирован: ${dateStr}`, margin, y)
  doc.text('ЗАО «Эмпилс» · 8 (800) 200-49-11 · www.empils.ru', pageWidth - margin, y, { align: 'right' })

  // Колонтитул на каждой странице
  const pageCount = doc.getNumberOfPages()
  for (let i = 1; i <= pageCount; i++) {
    doc.setPage(i)
    doc.setFontSize(7)
    doc.setTextColor(148, 163, 184)
    doc.text(
      `Страница ${i} из ${pageCount}`,
      pageWidth - margin,
      pageHeight - 8,
      { align: 'right' }
    )
  }

  const blob = doc.output('blob')
  const dStr = new Date().toISOString().split('T')[0]
  triggerDownload(blob, `Эмпилс_Радар_отчёт_${dStr}.pdf`)
}
