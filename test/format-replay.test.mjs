import {zipSync} from 'fflate'
import {unpack} from '../server/document/package.mjs'
import test from 'node:test'
import assert from 'node:assert/strict'
import {mkdtemp,rm} from 'node:fs/promises'
import {tmpdir} from 'node:os'
import {join} from 'node:path'
import {readOoxmlPackage,TreePackageStore,writeOoxmlPackage} from '@docx-editor.dev/core/store'
import {DocumentStore} from '../server/store.mjs'
import {createBlankDocument,executeAutomationBatch} from '../server/document/editor-automation.mjs'
import {format} from '../server/engine.mjs'
import {createFormatReplay} from '../server/document/format-replay.mjs'
import {compileFormatOps} from '../shared/editor-format-ops.mjs'
import {applyFormatReplay} from '../src/lib/format-replay.ts'
import {paragraphDocument} from './helpers/docx-fixture.mjs'

function replayEditor(bytes) {
 const pkg=readOoxmlPackage(bytes).package
 const store=new TreePackageStore(pkg,pkg.parts.get(pkg.mainDocumentPart))
 let reads=0,writes=0
 return {
  getMode:()=> 'edit',
  surface:{session:{part(){reads++;return store.bodyStore().part}},
  applyAutomationOps(build){
   const operations=build()
   writes++
   return store.transact({kind:'body'},ctx=>{for(const op of operations)if(!ctx.apply(op))throw Error('操作失败')})
  }},
  counts:()=>({reads,writes}),
 }
}

test('frontend replay preserves successive properties on the same paragraph',async()=>{
 const bytes=source()
 const result=await format(bytes,[
  {scope:'indices',indices:[1],paragraph:{firstLineIndent:24}},
  {scope:'indices',indices:[1],paragraph:{alignment:'Centered'}},
  {scope:'indices',indices:[1],paragraph:{firstLineIndent:0}},
 ])
 await applyFormatReplay(replayEditor(bytes),result.replay)
})

test('frontend replay batches 100 targets with bounded snapshot reads and matching fingerprint',async()=>{
 const bytes=paragraphDocument(Array.from({length:100},(_,i)=>`正文 ${i}`))
 const result=await format(bytes,[{scope:'all',font:{size:14}}])
 const editor=replayEditor(bytes),progress=[]
 await applyFormatReplay(editor,result.replay,(done,total)=>progress.push([done,total]))
 assert.equal(editor.counts().writes,5)
 assert.ok(editor.counts().reads<=8)
 assert.deepEqual(progress,[[20,100],[40,100],[60,100],[80,100],[100,100]])
})

test('replay progress exposes engine node identity without a second document search', async () => {
 const bytes=paragraphDocument(['同名标题','同名标题','正文'])
 const result=await format(bytes,[{scope:'indices',indices:[1],font:{size:14}}])
 const editor=replayEditor(bytes),locations=[]
 await applyFormatReplay(editor,result.replay,(_done,_total,_index,location)=>locations.push(location))
 assert.ok(locations[0]?.blockId)
 assert.equal(locations[0].text,'同名标题')
})
function source() {
 return executeAutomationBatch(createBlankDocument(),host=>{
  const ask=op=>host.execute({operations:[op]}).results[0].value
  const document=ask({op:'getDocument'}).handle,body=ask({op:'getBody',document}).handle
  return [{op:'replaceStoryBlocks',body,paragraphs:['标题','中文正文 English 123','下一段']}]
 }).bytes
}
test('serialized replay matches server results with overlapping operations',async()=>{
 const bytes=source(),result=await format(bytes,[{scope:'all',font:{name:'宋体',latinName:'Times New Roman',size:12}},{scope:'indices',indices:[1],paragraph:{firstLineIndent:24,lineSpacingMultiple:1.5},font:{size:14}}])
 const replay=JSON.parse(JSON.stringify(result.replay))
 assert.equal(replay.kind,'format-v1')
 const pkg=readOoxmlPackage(bytes).package,store=new TreePackageStore(pkg,pkg.parts.get(pkg.mainDocumentPart))
 for(const target of replay.targets)for(const index of target.indices){
  const ops=compileFormatOps(store.bodyStore().part,[index],target.op)
  assert.equal(store.transact({kind:'body'},ctx=>{for(const op of ops)if(!ctx.apply(op))return}).ok,true)
 }
 const compared=createFormatReplay(bytes,writeOoxmlPackage(store.currentPackage()),replay.targets)
 assert.equal(compared.before,replay.before)
 assert.equal(compared.after,replay.after)
})
test('only committed supported formats publish replay; structural formats use snapshots',async()=>{
 const root=await mkdtemp(join(tmpdir(),'replay-'))
 try{
  const store=new DocumentStore(root),doc=await store.create('实时格式.docx',source()),events=[]
  store.onCommit(event=>events.push(event))
  await assert.rejects(store.apply(doc.id,1,[{scope:'indices',indices:[999],font:{size:14}}]))
  assert.equal(events.length,0)
  await store.apply(doc.id,1,[{scope:'all',paragraph:{lineSpacingMultiple:1.5}}])
  assert.equal(events.length,1);assert.equal(events[0].previousRevision,1);assert.equal(events[0].revision,2)
  assert.equal(events[0].replay.kind,'format-v1')
  await store.apply(doc.id,2,[{scope:'sections',sectionIndices:[0],margins:{top:1440}}])
  assert.equal(events.length,2);assert.equal(events[1].replay,undefined)
  assert.equal((await store.metadata(doc.id)).versions[1].replay,undefined)
 }finally{await rm(root,{recursive:true,force:true})}
})

test('replay fingerprints ignore runtime IDs but still reject text and formatting changes',async()=>{
 const bytes=source(),parts=unpack(bytes),main=new TextDecoder().decode(parts['word/document.xml'])
 const make=(id,text='标题')=>{const copy={...parts};copy['word/document.xml']=new TextEncoder().encode(main.replace('<w:p>','<w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="'+id+'" w14:textId="'+id+'">').replace('标题',text));return zipSync(copy)}
 const a=createFormatReplay(make('11111111'),make('11111111'),[])
 const b=createFormatReplay(make('22222222'),make('22222222'),[])
 const formatted=await format(bytes,[{scope:'all',font:{size:18}}])
 assert.notEqual(createFormatReplay(bytes,bytes,[]).before,formatted.replay.after)
 assert.equal(a.before,b.before)
 assert.notEqual(a.before,createFormatReplay(make('22222222','不同内容'),make('22222222'),[]).before)
})
