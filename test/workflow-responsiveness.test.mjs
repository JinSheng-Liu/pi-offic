import test from 'node:test'
import assert from 'node:assert/strict'
import { DocumentWorkflow } from '../server/document/workflow.mjs'
import { paragraphDocument } from './helpers/docx-fixture.mjs'
import { runDocumentJob } from '../server/document/document-jobs.mjs'
import { unpack, withPackageCache } from '../server/document/package.mjs'

test('long document computation yields to main-thread requests before committing', async () => {
  const source = paragraphDocument(Array.from({ length: 1500 }, (_, i) => `正文 English ${i}`))
  let ticks = 0
  const timer = setInterval(() => ticks++, 10)
  const store = {
    exclusive: async (_id, _revision, fn) => fn({ revision: 1 }),
    bytes: async () => source,
    taskContext: { getStore: () => undefined },
    commit: async () => {
      assert.ok(ticks >= 2, `计算期间主线程响应次数只有 ${ticks}`)
      return { revision: 2 }
    },
  }
  try {
    const result = await new DocumentWorkflow(store).execute('test', 1, {
      kind: 'format', operations: [{ scope: 'all', font: { size: 14 } }],
    })
    assert.equal(result.document.revision, 2)
  } finally { clearInterval(timer) }
})

test('cancelling active computation never commits its result', async () => {
  const source = paragraphDocument(Array.from({length:1500},(_,i)=>`正文 ${i}`))
  const controller = new AbortController()
  let commits = 0
  const store = {
    exclusive: async (_id,_revision,fn)=>fn({revision:1}), bytes:async()=>source,
    taskContext:{getStore:()=>undefined}, commit:async()=>{commits++;return {revision:2}},
  }
  const pending = new DocumentWorkflow(store).execute('test',1,{
    kind:'format',operations:[{scope:'all',font:{size:14}}],
  },controller.signal)
  const timer = setTimeout(()=>controller.abort(),50)
  try { await assert.rejects(pending,{name:'AbortError'}); assert.equal(commits,0) }
  finally { clearTimeout(timer) }
  const next=await runDocumentJob('batch',{source:paragraphDocument(['恢复测试']),revision:1,
    batch:{kind:'format',operations:[{scope:'all',font:{size:14}}]}})
  assert.equal(next.result.affected,1)
})

test('package cache isolates mutable arrays and expires at job boundary', async () => {
  const bytes=paragraphDocument(['保护原文'])
  await withPackageCache(async()=>{
    const first=unpack(bytes),expected=first['word/document.xml'].slice()
    first['word/document.xml'].fill(0)
    delete first['word/styles.xml']
    const second=unpack(bytes)
    assert.deepEqual(second['word/document.xml'],expected)
    assert.ok(second['word/styles.xml'])
    await Promise.resolve()
    assert.deepEqual(unpack(bytes)['word/document.xml'],expected)
    const bad=bytes.slice();bad.fill(0)
    assert.throws(()=>unpack(bad))
  })
  assert.ok(unpack(bytes)['word/document.xml'].length)
})

test('queued job cancellation and worker errors release capacity for later work', async () => {
  const source=paragraphDocument(['队列测试'])
  const data={source,revision:1,batch:{kind:'format',operations:[{scope:'all',font:{size:14}}]}}
  const first=runDocumentJob('batch',data),second=runDocumentJob('batch',data)
  const controller=new AbortController()
  const queued=runDocumentJob('batch',data,controller.signal)
  controller.abort()
  await assert.rejects(queued,{name:'AbortError'})
  await Promise.all([first,second])
  await assert.rejects(runDocumentJob('batch',{...data,batch:{kind:'format',operations:[
    {scope:'indices',indices:[999],font:{size:12}},
  ]}}),error=>error.status===400 && /索引/.test(error.message))
  assert.equal((await runDocumentJob('batch',data)).result.affected,1)
})
