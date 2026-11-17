import {readFile,readdir,mkdir,writeFile} from 'node:fs/promises'
import assert from 'node:assert/strict'
import {format,inspect,unpack,digest} from '../server/engine.mjs'
import {readEditorDocument,findTextRanges} from '../server/document/editor-reader.mjs'
const root='data/documents/d2cecf36-e7a4-4682-afb5-3b29875b0fbe'
const meta=JSON.parse(await readFile(root+'/metadata.json','utf8'))
const original=await readFile(root+'/'+meta.versions[0].file)
let operations
for(const file of await readdir(root+'/agent/sessions')) for(const line of (await readFile(root+'/agent/sessions/'+file,'utf8')).trim().split('\n')) {
  const message=JSON.parse(line).message
  if(message?.role==='assistant'&&Array.isArray(message.content)) for(const content of message.content) if(content.type==='toolCall'&&content.name==='apply_formatting') operations=content.arguments.operations
}
assert.equal(operations?.length,19)
let result
if(process.argv.includes('--api')) {
  const base='http://127.0.0.1:5175'
  assert.equal((await(await fetch(base+'/api/health')).json()).testMode,true)
  const uploaded=await fetch(base+'/api/documents?name='+encodeURIComponent('论文接口验收副本.docx'),{method:'POST',headers:{'content-type':'application/vnd.openxmlformats-officedocument.wordprocessingml.document'},body:original})
  assert.equal(uploaded.status,201)
  const doc=await uploaded.json()
  const response=await fetch(base+`/api/documents/${doc.id}/format`,{method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify({revision:1,operations})})
  const responseBody=await response.json()
  assert.equal(response.status,200,JSON.stringify(responseBody))
  result={bytes:new Uint8Array(await(await fetch(base+`/api/documents/${doc.id}/file`)).arrayBuffer()),validation:responseBody.validation}
  console.log(JSON.stringify({apiDocument:doc.id,status:response.status}))
} else result=await format(original,operations)
const before=unpack(original),after=unpack(result.bytes)
const bodyBefore=readEditorDocument(original)
const bodyAfter=readEditorDocument(result.bytes)
assert.deepEqual(bodyAfter.paragraphs.map(p=>p.text),bodyBefore.paragraphs.map(p=>p.text))
const media=Object.keys(before).filter(name=>name.startsWith('word/media/')&&!name.endsWith('/'))
for(const name of media) assert.equal(digest(after[name]),digest(before[name]),name)
for(const op of operations.filter(op=>op.scope==='text')) {
  const selected=findTextRanges(bodyAfter.paragraphs[op.indices[0]],op.text)[op.occurrence??0]
  for(const [key,value] of Object.entries(op.font)) assert.equal(selected.font[key],value)
}
const styles=new TextDecoder().decode(after['word/styles.xml'])
for(const id of new Set(operations.map(op=>op.style).filter(Boolean))) assert.ok(styles.includes(`w:styleId="${id}"`),`缺少 ${id}`)
const inspection=await inspect(result.bytes)
assert.equal(inspection.paragraphs.filter(p=>p.heading>0).length,29)
await mkdir('data/validation/thesis',{recursive:true})
await writeFile('data/validation/thesis/毕业论文_处理验收副本.docx',result.bytes)
const report={operations:19,paragraphs:bodyAfter.paragraphs.length,textPreserved:true,mediaPreserved:media.length,headings:inspection.paragraphs.filter(p=>p.heading>0).length,validation:result.validation,sourceRevisionUnchanged:JSON.parse(await readFile(root+'/metadata.json','utf8')).revision===meta.revision}
await writeFile('data/validation/thesis/report.json',JSON.stringify(report,null,2))
console.log(JSON.stringify(report))
