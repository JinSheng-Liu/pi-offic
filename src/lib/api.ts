const CRLF = String.fromCharCode(13, 10)
const LF = String.fromCharCode(10)
const DOUBLE_LF = LF + LF

export async function api<T>(url: string, options: RequestInit = {}): Promise<T> {
  const response = await fetch(url, options)
  if (!response.ok) {
    let message = `请求失败 (${response.status})`
    try {
      const body = await response.json()
      message = body?.error || message
    } catch {
      /* Ignore non-JSON error body */
    }
    throw new Error(`${url}: ${message}`)
  }
  return response.json() as Promise<T>
}

export async function consumeEvents(
  response: Response,
  onEvent: (event: string, data: any) => void | Promise<void>,
): Promise<void> {
  if (!response.ok) {
    let message = `请求失败 (${response.status})`
    try {
      const body = await response.json()
      message = body?.error || message
    } catch {
      /* Ignore non-JSON error body */
    }
    throw new Error(message)
  }

  const reader = response.body?.getReader()
  if (!reader) throw new Error('浏览器不支持流式响应')

  const decoder = new TextDecoder()
  let pending = ''

  const handlers = new Set<Promise<void>>()
  let handlerError: unknown
  let hasHandlerError = false
  const recordError = (error: unknown) => {
    if (!hasHandlerError) handlerError = error
    hasHandlerError = true
  }

  const processFrame = (rawFrame: string) => {
    const frame = rawFrame.split(CRLF).join(LF).trim()
    if (!frame) return
    const eventMatch = frame.match(/^event:\s*(.+)$/m)
    const dataMatch = frame.match(/^data:\s*(.+)$/m)
    if (!eventMatch || !dataMatch) return
    const payload = JSON.parse(dataMatch[1]!)
    try {
      const result = onEvent(eventMatch[1]!.trim(), payload)
      if (result && typeof (result as Promise<void>).then === 'function') {
        const pendingHandler = Promise.resolve(result).catch(recordError)
        handlers.add(pendingHandler)
        void pendingHandler.then(() => handlers.delete(pendingHandler))
      }
    } catch (error) {
      recordError(error)
    }
  }

  try {
    while (true) {
      const { done, value } = await reader.read()
      pending += decoder.decode(value, { stream: !done })
      const normalized = pending.split(CRLF).join(LF)
      const frames = normalized.split(DOUBLE_LF)
      pending = frames.pop() || ''
      for (const frame of frames) {
        processFrame(frame)
      }
      if (done) {
        if (pending.trim()) processFrame(pending)
        break
      }
    }
  } finally {
    // Recovery must never race an editor update already in flight, including
    // when the transport fails before the final SSE frame arrives.
    await Promise.all(handlers)
    reader.releaseLock()
  }
  if (hasHandlerError) throw handlerError
}
