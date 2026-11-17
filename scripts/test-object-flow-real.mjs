import {readFile,writeFile,mkdir} from 'node:fs/promises'
import assert from 'node:assert/strict'
import {DocumentStore} from '../server/store.mjs'
import {inspect,unpack,digest} from '../server/engine.mjs'
const out='data/validation/object-flow-real'
await mkdir(out,{recursive:true})
const task=JSON.parse(await readFile('data/validation/skills-thesis/task.json','utf8'))
const source=`data/testing/documents/${task.id}/original.docx`,original=await readFile(source)
assert.equal(digest(original),task.hash)
const store=new DocumentStore(out+'/documents')
let meta=await store.create('通用对象处理回放.docx',original)
const apply=async operations=>{meta=(await store.apply(meta.id,meta.revision,operations)).document}
const calls=JSON.parse(await readFile('data/validation/skills-thesis/tool-usage.json','utf8')).calls
for(const call of calls.filter(c=>c.name==='apply_formatting'))await apply(call.args.operations)
const before=await inspect(await store.bytes(meta.id))
const conversions=[]
for(const mode of ['captionInFlow','inlinePicture']){
  const data=await inspect(await store.bytes(meta.id))
  const ids=data.floatingObjects.filter(o=>o.flowCapability.supported&&o.flowCapability.mode===mode).map(o=>o.objectIndex)
  if(ids.length){await apply([{scope:'floatingObjects',objectIndices:ids,objectLayout:mode}]);conversions.push({mode,count:ids.length})}
}
let data=await inspect(await store.bytes(meta.id))
const tocTitles=data.paragraphs.filter(p=>p.text==='目录').map(p=>p.index)
if(tocTitles.length){await apply([{scope:'tocTitle',indices:tocTitles}]);data=await inspect(await store.bytes(meta.id))}
const tocStart=data.paragraphs.findIndex(p=>p.text==='目录')
const tocEnd=data.paragraphs.findIndex((p,i)=>i>tocStart&&p.heading>0)
const entries=data.paragraphs.filter(p=>p.index>tocStart&&p.index<tocEnd&&/^[^\t]+\t[0-9IVXLCDMivxlcdm]+\s*$/.test(p.text)&&!p.tabStops.effective.length).map(p=>p.index)
if(entries.length){await apply([{scope:'indices',indices:entries,paragraph:{tocPageTabPosition:415}}]);data=await inspect(await store.bytes(meta.id))}
const captionIds=data.paragraphs.filter(p=>/^(?:图|表)\s*\d+[.－-]\d+/.test(p.text)&&!p.tableNestingLevel&&!p.inlinePictureCount).map(p=>p.index)
const titles=data.paragraphs.filter(p=>p.text==='目录').map(p=>p.index)
const ops=[]
if(captionIds.length)ops.push({scope:'indices',indices:captionIds,font:{name:'宋体',latinName:'Times New Roman',size:10.5},paragraph:{alignment:'Centered',firstLineIndent:0,lineSpacingMultiple:1,keepLines:true,keepNext:false}})
if(titles.length)ops.push({scope:'indices',indices:titles,font:{name:'黑体',size:16},paragraph:{alignment:'Centered',lineSpacingMultiple:1.5,keepLines:true}})
if(ops.length)await apply(ops)
const output=await store.bytes(meta.id),after=await inspect(output)
const a=unpack(original),b=unpack(output)
for(const name of Object.keys(a).filter(p=>p.startsWith('word/media/')))assert.equal(digest(a[name]),digest(b[name]),name)
assert.equal(after.tables.length,before.tables.length)
assert.equal(before.layoutRisks.floatingCount-after.layoutRisks.floatingCount,conversions.reduce((sum,x)=>sum+x.count,0))
assert.equal(digest(await readFile(source)),task.hash)
await writeFile(out+'/系统处理回放.docx',output)
await writeFile(out+'/report.json',JSON.stringify({sourceHash:task.hash,id:meta.id,revision:meta.revision,conversions,originalUnchanged:true,mediaUnchanged:true,tables:after.tables.length,remainingFloating:after.floatingObjects.map(o=>({text:o.text,reason:o.flowCapability.reason})),visualVerified:false},null,2))
console.log(JSON.stringify({revision:meta.revision,conversions,remainingFloating:after.floatingObjects.length,output:out+'/系统处理回放.docx'}))
