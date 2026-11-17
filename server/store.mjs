import { AsyncLocalStorage } from 'node:async_hooks'
import { mkdir, readFile, writeFile, readdir } from 'node:fs/promises'
import {atomicReplace} from './atomic-replace.mjs'
import { join } from 'node:path'
import { randomUUID } from 'node:crypto'
import { DocumentWorkflow } from './document/workflow.mjs'
import { createBlankDocument } from './document/editor-automation.mjs'
import { unpack, digest, problem, inspect, compareProtected } from './engine.mjs'
import { qualityReport } from './document/quality-report.mjs'

export class DocumentStore {
  constructor(root) { this.taskContext = new AsyncLocalStorage(); this.root = root; this.locks = new Set(); this.workflow = new DocumentWorkflow(this); this.commitListeners = new Set() }
  withTask(task, operation) { return this.taskContext.run(task, operation) }
  withStep(label, operation) { const task=this.taskContext.getStore(); return task?this.withTask({...task,stepLabel:label},operation):operation() }
  async finishTask(id,task,result) {
    const meta=await this.metadata(id)
    return this.exclusive(id,meta.revision,async current=>{
      if(!current.versions.some(v=>v.task?.id===task.id))return
      const next={...current,tasks:{...current.tasks,[task.id]:{...task,state:result.conversationState==='cancelled'?'cancelled':result.conversationState==='failed'||result.plans?.some(p=>p.state!=='executed')||result.checklist?.items.some(i=>['pending','awaitingVerification','blocked','needsDecision'].includes(i.state))?'partial':'finished',finishedAt:result.finishedAt}}}
      if (next.tasks[task.id].state === 'finished') {
        const last = current.versions.filter(v => v.task?.id === task.id).at(-1)
        const previous = (current.savedVersions || []).at(-1)
        const previousHash = current.versions.find(v => v.revision === previous?.revision)?.hash
        next.savedVersions = (current.savedVersions || []).filter(v => v.taskId !== task.id)
        if (last.hash !== previousHash) next.savedVersions.push({ revision: last.revision, summary: task.summary, kind: 'task', taskId: task.id, time: result.finishedAt || last.time })
        next.pendingWork = null
      } else {
        next.pendingWork = { taskId: task.id, startRevision: current.pendingWork?.startRevision ?? task.startRevision, revision: current.revision, summary: task.summary }
      }
      await this.atomicMetadata(next)
    })
  }

  async publishPartial(id, revision) {
    return this.exclusive(id, revision, async meta => {
      if (!meta.pendingWork) throw problem('没有待保存的部分结果', 409)
      const next = { ...meta, pendingWork: null, savedVersions: [...(meta.savedVersions || []), {
        revision, summary: meta.pendingWork.summary, kind: 'partial', time: new Date().toISOString(),
      }] }
      await this.atomicMetadata(next)
      return next
    })
  }
  onCommit(listener) { this.commitListeners.add(listener); return () => this.commitListeners.delete(listener) }
  folder(id) { if (!/^[a-f0-9-]{36}$/.test(id)) throw problem('文档 ID 无效'); return join(this.root, id) }
  async metadata(id) {
    try {
      const meta = JSON.parse(await readFile(join(this.folder(id), 'metadata.json'), 'utf8'))
      if (!meta.savedVersions) {
        meta.savedVersions = meta.versions.filter(v => !v.task || (
          meta.tasks?.[v.task.id]?.state === 'finished' &&
          meta.versions.filter(item => item.task?.id === v.task.id).at(-1)?.revision === v.revision
        )).map(v => ({ revision: v.revision, summary: v.task?.summary || v.summary, kind: v.task ? 'task' : 'manual', time: v.time }))
      }
      return meta
    }
    catch (e) { if (e.code === 'ENOENT') throw problem('文档不存在', 404); throw e }
  }
  async list() {
    await mkdir(this.root, { recursive: true })
    const entries = await readdir(this.root, { withFileTypes: true })
    const docs = []
    for (const e of entries) if (e.isDirectory() && /^[a-f0-9-]{36}$/.test(e.name)) {
      try { docs.push(await this.metadata(e.name)) } catch (error) { if (error.status !== 404) throw error }
    }
    return docs.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt))
  }
  async atomicMetadata(meta) {
    const path = join(this.folder(meta.id), `metadata-${randomUUID()}.tmp`)
    await writeFile(path, JSON.stringify(meta, null, 2), { encoding: 'utf8', flag: 'wx' })
    await atomicReplace(path, join(this.folder(meta.id), 'metadata.json'))
  }
  async create(name, bytes) {
    if (!bytes) {
      bytes = createBlankDocument()
    }
    unpack(bytes)
    const inspected = await inspect(bytes)
    const id = randomUUID(), now = new Date().toISOString()
    await mkdir(this.folder(id), { recursive: true })
    await writeFile(join(this.folder(id), 'original.docx'), bytes, { flag: 'wx' })
    const file = `${randomUUID()}.docx`
    await writeFile(join(this.folder(id), file), bytes, { flag: 'wx' })
    const safeName = String(name || '未命名文档.docx').replace(/[\\/\x00-\x1f]/g, '_').slice(0, 160)
    const meta = { id, name: safeName.endsWith('.docx') ? safeName : `${safeName}.docx`, revision: 1, updatedAt: now,
      savedVersions: [{ revision: 1, summary: '原始文档', kind: 'original', time: now }],
      versions: [{ revision: 1, file, hash: digest(bytes), size: bytes.length, time: now, summary: '原始文档',
        validation: qualityReport(inspected, { checks: ['DOCX 解析成功'], limitations: [] }, {contentVerified:false}) }] }
    await this.atomicMetadata(meta)
    return meta
  }
  async bytes(id, revision) {
    const meta = await this.metadata(id)
    const version = meta.versions.find(v => v.revision === (revision ?? meta.revision))
    if (!version) throw problem('版本不存在', 404)
    const bytes = await readFile(join(this.folder(id), version.file))
    if (digest(bytes) !== version.hash) throw problem('版本文件哈希不匹配，请恢复其他版本', 422)
    return bytes
  }
  async exclusive(id, expected, operation) {
    this.folder(id)
    if (this.locks.has(id)) throw problem('文档正在处理中，请稍后重试', 409)
    this.locks.add(id)
    try {
      const meta = await this.metadata(id)
      if (!Number.isInteger(expected) || expected !== meta.revision) throw problem('文档版本已变化，请重新加载后重试', 409)
      return await operation(meta)
    } finally { this.locks.delete(id) }
  }
  async commit(meta, bytes, summary, validation, replay, task) {
    const file = `${randomUUID()}.docx`, revision = meta.revision + 1, time = new Date().toISOString()
    await writeFile(join(this.folder(meta.id), file), bytes, { flag: 'wx' })
    const saved = await readFile(join(this.folder(meta.id), file))
    if (digest(saved) !== digest(bytes)) throw problem('磁盘写入验证失败', 500)
    const next = { ...meta, revision, updatedAt: time, versions: [...meta.versions,
      { revision, file, hash: digest(bytes), size: bytes.length, time, summary, validation, ...(task?{task}: {}) }] }
    next.pendingWork = task ? { taskId: task.id, startRevision: meta.pendingWork?.startRevision ?? task.startRevision, revision, summary: task.summary } : null
    if (!task && digest(bytes) !== meta.versions.at(-1).hash) {
      next.savedVersions = [...(meta.savedVersions || []), { revision, summary, kind: 'manual', time }]
      next.pendingWork = null
    }
    await this.atomicMetadata(next)
    for (const listener of this.commitListeners) { try { listener({id:meta.id, revision, summary, ...(replay?{previousRevision:meta.revision,replay}:{})}) } catch { /* A disconnected observer cannot roll back a committed version. */ } }
    return next
  }
  async save(id, revision, bytes) {
    return this.exclusive(id, revision, async meta => {
      const before = await this.bytes(id)
      if (digest(before) === digest(bytes)) return meta
      const validation = compareProtected(before, bytes, { formatting: false })
      const data = await inspect(bytes)
      return this.commit(meta, bytes, '手动编辑', qualityReport(data, validation, {contentVerified:false}))
    })
  }
  async apply(id, revision, operations, signal) {
    return this.workflow.execute(id, revision, {kind:'format', operations}, signal)
  }
  async addCaptions(id, revision, captions, signal) {
    return this.workflow.execute(id, revision, {kind:'captions', captions}, signal)
  }
  async restore(id, revision, target) {
    return this.exclusive(id, revision, async meta => {
      const bytes = await this.bytes(id, target)
      unpack(bytes)
      const validation = meta.versions.find(version => version.revision === target).validation
      return this.commit(meta, bytes, `恢复版本 ${target}`, { ...validation, checks: [...validation.checks, '历史文件哈希匹配'], limitations: [...validation.limitations, '恢复不重新验证排版'] })
    })
  }
}
