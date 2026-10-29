// Хранилище сообщений в JSON-файле GitHub репозитория через GitHub API.
// Это надёжнее Netlify Blobs: данные гарантированно сохраняются и видны в репо.

const GITHUB_TOKEN = process.env.GITHUB_TOKEN || process.env.GH_TOKEN
const REPO_OWNER = 'Iverson76newcomer'
const REPO_NAME = 'radar-empils'
const FILE_PATH = 'data/messages.json'
const BRANCH = 'main'

const GITHUB_API = `https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/contents/${FILE_PATH}`

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
  updatedAt: string
}

interface StorageData {
  messages: Message[]
}

// In-memory кэш для быстрого чтения в одном запросе
let cache: { data: StorageData; sha: string } | null = null

async function fetchFromGitHub(): Promise<{ data: StorageData; sha: string }> {
  const res = await fetch(`${GITHUB_API}?ref=${BRANCH}`, {
    headers: {
      Authorization: `Bearer ${GITHUB_TOKEN}`,
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
    },
    cache: 'no-store',
  })

  if (res.status === 404) {
    return { data: { messages: [] }, sha: '' }
  }
  if (!res.ok) {
    throw new Error(`GitHub API error: ${res.status} ${await res.text()}`)
  }

  const json = await res.json()
  const content = Buffer.from(json.content, 'base64').toString('utf-8')
  const data: StorageData = JSON.parse(content || '{"messages":[]}')

  // Защита от кривых данных
  if (!data.messages || !Array.isArray(data.messages)) {
    data.messages = []
  }

  return { data, sha: json.sha }
}

async function saveToGitHub(data: StorageData, currentSha?: string): Promise<string> {
  // Сначала получаем актуальный SHA (если не передан), чтобы избежать конфликта
  let sha = currentSha
  if (!sha) {
    const existing = await fetchFromGitHub()
    sha = existing.sha
  }

  const content = Buffer.from(JSON.stringify(data, null, 2), 'utf-8').toString('base64')

  const body: any = {
    message: `chore: update messages.json (${data.messages.length} total)`,
    content,
    branch: BRANCH,
  }
  if (sha) body.sha = sha

  const res = await fetch(GITHUB_API, {
    method: 'PUT',
    headers: {
      Authorization: `Bearer ${GITHUB_TOKEN}`,
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(body),
  })

  if (!res.ok) {
    const text = await res.text()
    throw new Error(`GitHub API PUT error: ${res.status} ${text}`)
  }

  const json = await res.json()
  return json.content.sha
}

// Retry-логика для разрешения конфликтов (когда 2 запроса пишут одновременно)
async function saveWithRetry(data: StorageData, maxRetries = 3): Promise<void> {
  let attempt = 0
  while (attempt < maxRetries) {
    try {
      const { data: current, sha } = await fetchFromGitHub()

      // Мержим: добавляем новые сообщения к существующим
      // (важно, т.к. между чтением и записью могло что-то добавиться)
      const existingIds = new Set(current.messages.map(m => m.id))
      const newMessages = data.messages.filter(m => !existingIds.has(m.id))
      const merged: StorageData = {
        messages: [...current.messages, ...newMessages],
      }

      await saveToGitHub(merged, sha)
      return
    } catch (err) {
      attempt++
      if (attempt >= maxRetries) throw err
      // Пауза 200ms перед повтором
      await new Promise(resolve => setTimeout(resolve, 200 * attempt))
    }
  }
}

export const githubStore = {
  // Получить все сообщения
  async getAll(): Promise<Message[]> {
    if (cache) return cache.data.messages
    const { data } = await fetchFromGitHub()
    cache = { data, sha: '' }
    return data.messages
  },

  // Добавить новое сообщение
  async add(message: Message): Promise<Message> {
    const { data: current, sha } = await fetchFromGitHub()
    const updated: StorageData = {
      messages: [...current.messages, message],
    }
    await saveToGitHub(updated, sha)
    cache = null
    return message
  },

  // Обновить статус сообщения
  async update(id: string, updates: Partial<Message>): Promise<Message | null> {
    const { data: current, sha } = await fetchFromGitHub()
    const idx = current.messages.findIndex(m => m.id === id)
    if (idx === -1) return null

    const updated: Message = {
      ...current.messages[idx],
      ...updates,
      updatedAt: new Date().toISOString(),
    }
    current.messages[idx] = updated

    await saveToGitHub(current, sha)
    cache = null
    return updated
  },

  // Удалить сообщение
  async delete(id: string): Promise<boolean> {
    const { data: current, sha } = await fetchFromGitHub()
    const filtered = current.messages.filter(m => m.id !== id)
    if (filtered.length === current.messages.length) return false

    const updated: StorageData = { messages: filtered }
    await saveToGitHub(updated, sha)
    cache = null
    return true
  },
}
