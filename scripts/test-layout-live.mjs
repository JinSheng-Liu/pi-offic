import {readFile,writeFile,mkdir} from 'node:fs/promises'
import assert from 'node:assert/strict'
import {inspect} from '../server/engine.mjs'
import {readEditorDocument,findTextRanges} from '../server/document/editor-reader.mjs'
const base='http://127.0.0.1:5175'
const health=await(await fetch(base+'/api/health')).json()
assert.equal(health.testMode,true)
assert.ok(health.capabilities.includes('text-range-format'))
await mkdir('data/validation/layout',{recursive:true})
const original=await readFile('D:/xiangmu/agent-office/agent-office/测试文档/党政机关公文_无格式版.docx')
const created=await fetch(base+'/api/documents?name='+encodeURIComponent('真实模型段内排版验收.docx'),{method:'POST',headers:{'content-type':'application/vnd.openxmlformats-officedocument.wordprocessingml.document'},body:original})
assert.equal(created.status,201)
const doc=await created.json()
const uploaded=await(await fetch(base+`/api/documents/${doc.id}/validate`)).json()
assert.deepEqual(uploaded.paragraphs.map(p=>p.text),(await inspect(original)).paragraphs.map(p=>p.text),'上传内容必须在模型调用前验证')
const response=await fetch(base+`/api/documents/${doc.id}/chat`,{method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify({revision:1,message:'请实际修改当前文档：全文仿宋三号；通知中的“一、提高统筹认识。”至“五、强化项目管理。”这五个段内标题只把标题部分设为黑体三号，后续正文保持仿宋三号。附件中“一、工作目标”至“六、组织保障”六个独立标题设黑体三号，与下段同页并段中不分页。不要增删或改写任何文字，不要把全部正文设为段中不分页。逐项报告结果，不要声称已验证 Word 页面。'}),signal:AbortSignal.timeout(180000)})
assert.equal(response.status,200)
const events=await response.text()
await writeFile('data/validation/layout/live-events.txt',events)
assert.ok(!events.includes('event: error'), '真实模型请求失败，见隔离事件记录')
const output=new Uint8Array(await(await fetch(base+`/api/documents/${doc.id}/file`)).arrayBuffer())
await writeFile('data/validation/layout/真实模型排版验收.docx',output)
const before=await inspect(original),after=await inspect(output)
assert.deepEqual(after.paragraphs.map(p=>p.text),before.paragraphs.map(p=>p.text))
const body=readEditorDocument(output)
for(const i of [5,6,7,8,9]) {
  const p=body.paragraphs[i],prefix=p.text.split('。')[0]+'。'
  assert.equal(findTextRanges(p,prefix)[0].font.name,'黑体',`段内标题 ${i}`)
  assert.equal(findTextRanges(p,p.text.slice(prefix.length))[0].font.name,'仿宋',`正文 ${i}`)
  for(const r of findTextRanges(p,prefix)[0].runs) assert.equal(r.value.rPr.rFonts.eastAsia,'黑体')
}
for(const i of [16,18,20,22,24,26]) {
  assert.equal(after.paragraphs[i].pagination.keepNext,true)
  assert.equal(after.paragraphs[i].pagination.keepLines,true)
}
const report={id:doc.id,mixedHeadings:5,independentHeadings:6,textPreserved:true,wordRenderingVerified:false}
await writeFile('data/validation/layout/live-report.json',JSON.stringify(report,null,2))
console.log(JSON.stringify(report))
