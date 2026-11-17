import { randomUUID } from 'node:crypto'
import { problem } from './document/package.mjs'

// A request is scoped to one running document task. No arbitrary editor commands.
export class EditorRequests {
  contexts = new Map()
  pending = new Map()
  attach(id, send, signal) {
    this.contexts.set(id, { send, signal })
    return () => { this.contexts.delete(id) }
  }
  async request(id, payload, bytes) {
    const context = this.contexts.get(id)
    if (!context) throw problem('原生目录需要保持文档编辑页面连接', 409)
    context.signal.throwIfAborted()
    if (this.pending.has(id)) throw problem('已有编辑器操作等待完成', 409)
    const token = randomUUID()
    return new Promise((resolve, reject) => {
      const finish = (error, value) => {
        clearTimeout(timer)
        context.signal.removeEventListener('abort', cancel)
        this.pending.delete(id)
        error ? reject(error) : resolve(value)
      }
      const cancel = () => finish(problem('原生目录操作已取消', 409))
      const timer = setTimeout(() => finish(problem('原生目录操作等待超时', 408)), 90000)
      this.pending.set(id, { token, bytes, payload, finish })
      context.signal.addEventListener('abort', cancel, { once: true })
      try { context.send({ ...payload, token }) } catch (error) { finish(error) }
    })
  }
  get(id, token) {
    const pending = this.pending.get(id)
    if (!pending || pending.token !== token) throw problem('编辑器请求已失效', 409)
    return pending
  }
  complete(id, token, bytes, error) {
    this.get(id, token).finish(error ? problem(String(error).slice(0,500), 422) : null, bytes)
  }
}
