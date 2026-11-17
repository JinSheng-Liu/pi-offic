import test from 'node:test'
import assert from 'node:assert/strict'
import {mkdtemp,rm} from 'node:fs/promises'
import {tmpdir} from 'node:os'
import {join} from 'node:path'
import {DocumentStore} from '../server/store.mjs'
import {groupHistory} from '../src/lib/history.ts'
import {executeFormatPlan} from '../server/document/format-plan.mjs'
test('task checkpoints group persistently while manual saves and restores remain separate',async()=>{
 const root=await mkdtemp(join(tmpdir(),'history-'))
 try{
 const store=new DocumentStore(root),doc=await store.create('历史.docx'),events=[]
 store.onCommit(e=>events.push(e))
 const task={id:'task-a',summary:'排版',startRevision:1,startedAt:new Date().toISOString()}
 const plan=await store.withTask(task,()=>executeFormatPlan(store,doc.id,1,[{label:'字体',operations:[{scope:'all',font:{size:14}}]},{label:'对齐',operations:[{scope:'all',paragraph:{alignment:'Centered'}}]}]))
 assert.equal(plan.currentRevision,3);assert.equal(events.length,2)
 await store.finishTask(doc.id,task,{conversationState:'complete',plans:[plan],finishedAt:new Date().toISOString()})
 let meta=await new DocumentStore(root).metadata(doc.id),groups=groupHistory(meta.versions,meta.tasks)
 assert.equal(groups.length,2);assert.equal(groups[0].versions.length,2);assert.equal(groups[0].task.state,'finished')
 assert.equal(groups[0].versions[0].task.stepLabel,'对齐')
 await store.save(doc.id,3,await store.bytes(doc.id))
 await store.restore(doc.id,3,1)
 meta=await store.metadata(doc.id);assert.equal(meta.revision,4)
 assert.equal(groupHistory(meta.versions,meta.tasks).length,3)
 assert.equal(meta.versions[3].task,undefined)
 assert.equal(meta.savedVersions.length,3)
 }finally{await rm(root,{recursive:true,force:true})}
})
test('legacy revisions stay separate and unrelated tasks never merge by time',()=>{
 const v=(revision,task)=>({revision,time:'same',summary:'test',validation:{},task})
 assert.equal(groupHistory([v(1),v(2),v(3)]).length,3)
 assert.equal(groupHistory([v(1,{id:'a'}),v(2,{id:'b'})]).length,2)
})
test('partial and cancelled tasks retain checkpoints; zero-write requests add no history',async()=>{
 const root=await mkdtemp(join(tmpdir(),'history-status-'))
 try{
 const store=new DocumentStore(root),doc=await store.create('状态.docx')
 const task={id:'partial',summary:'测试',startRevision:1,startedAt:new Date().toISOString()}
 const plan=await store.withTask(task,()=>executeFormatPlan(store,doc.id,1,[{label:'成功',operations:[{scope:'all',font:{size:14}}]},{label:'失败',operations:[{scope:'indices',indices:[999],font:{size:14}}]}]))
 await store.finishTask(doc.id,task,{conversationState:'complete',plans:[plan],finishedAt:new Date().toISOString()})
 assert.equal((await store.metadata(doc.id)).tasks.partial.state,'partial')
 await store.finishTask(doc.id,task,{conversationState:'cancelled',plans:[plan],finishedAt:new Date().toISOString()})
 assert.equal((await store.metadata(doc.id)).tasks.partial.state,'cancelled')
 await store.finishTask(doc.id,{...task,id:'empty'},{conversationState:'failed',plans:[]})
 const meta=await store.metadata(doc.id)
 assert.equal(meta.tasks.empty,undefined);assert.equal(meta.revision,2)
 }finally{await rm(root,{recursive:true,force:true})}
})


test('partial work is published only on explicit save and publishing is idempotently refused', async () => {
  const root = await mkdtemp(join(tmpdir(), 'saved-version-'))
  try {
    const store = new DocumentStore(root)
    const doc = await store.create('测试.docx')
    const task = { id: 'partial-save', summary: '排版', startRevision: 1 }
    await store.withTask(task, () => store.apply(doc.id, 1, [{ scope: 'all', font: { size: 14 } }]))
    await store.finishTask(doc.id, task, { conversationState: 'cancelled' })
    let meta = await store.metadata(doc.id)
    assert.equal(meta.savedVersions.length, 1)
    assert.equal(meta.pendingWork.startRevision, 1)
    meta = await store.publishPartial(doc.id, 2)
    assert.equal(meta.savedVersions.length, 2)
    assert.equal(meta.savedVersions[1].kind, 'partial')
    assert.equal(meta.pendingWork, null)
    await assert.rejects(() => store.publishPartial(doc.id, 2), /没有待保存/)
  } finally {
    await rm(root, { recursive: true, force: true })
  }
})
