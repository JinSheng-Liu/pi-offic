import { Worker } from 'node:worker_threads'

// Bound CPU and memory use across documents. Storage locks remain in DocumentStore.
const waiting = []
let running = 0
const MAX_RUNNING = 2

function pump() {
  while (running < MAX_RUNNING && waiting.length) waiting.shift()()
}

export function runDocumentJob(kind, data, signal) {
  signal?.throwIfAborted()
  if (waiting.length >= 16) return Promise.reject(Object.assign(new Error('文档计算队列已满，请稍后重试'), { status: 503 }))
  return new Promise((resolve, reject) => {
    let worker, started = false, finished = false
    const finish = async (error, result) => {
      if (finished) return
      finished = true
      signal?.removeEventListener('abort', abort)
      const index = waiting.indexOf(start)
      if (index !== -1) waiting.splice(index, 1)
      try { if (worker) await worker.terminate() }
      catch (terminationError) { error ||= terminationError }
      finally {
        if (started) running--
        pump()
        if (error) reject(error)
        else resolve(result)
      }
    }
    const abort = () => { void finish(signal.reason || new DOMException('已取消', 'AbortError')) }
    const start = () => {
      if (finished) return
      started = true
      running++
      try {
        worker = new Worker(new URL('./document-worker.mjs', import.meta.url), {
          workerData: { kind, data }, execArgv: [],
        })
        worker.once('message', ({ error, result }) => {
          void finish(error ? Object.assign(new Error(error.message), error) : null, result)
        })
        worker.once('error', error => { void finish(error) })
        worker.once('exit', code => {
          if (!finished) void finish(new Error(`文档计算线程提前退出 (${code})`))
        })
      } catch (error) { void finish(error) }
    }
    signal?.addEventListener('abort', abort, { once: true })
    waiting.push(start)
    pump()
  })
}
