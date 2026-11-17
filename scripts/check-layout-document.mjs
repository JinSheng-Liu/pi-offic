import {readFile, writeFile, mkdir} from 'node:fs/promises'
import assert from 'node:assert/strict'
import {format, inspect} from '../server/engine.mjs'
import {readEditorDocument,findTextRanges} from '../server/document/editor-reader.mjs'

const original = await readFile('D:/xiangmu/agent-office/agent-office/测试文档/党政机关公文_无格式版.docx')
const data = await inspect(original)
const historical = JSON.parse(await readFile('data/documents/4a4aee96-87ef-45f8-ba90-f66204739da1/metadata.json','utf8'))
const formatted = await readFile(`data/documents/${historical.id}/${historical.versions[1].file}`)
const prefixes = [5,6,7,8,9].map(i => ({i,text:data.paragraphs[i].text.split('。')[0]+'。'}))
const result = await format(formatted, [
  {scope:'all',font:{name:'仿宋',size:16},paragraph:{widowControl:true,keepLines:false}},
  {scope:'indices',indices:[0],font:{name:'小标宋',size:22}},
  {scope:'indices',indices:[2,15],font:{name:'方正小标宋简体',size:22},paragraph:{keepNext:true,keepLines:true}},
  ...prefixes.map(({i,text})=>({scope:'text',indices:[i],text,font:{name:'黑体',size:16}})),
  {scope:'indices',indices:[16,18,20,22,24,26],font:{name:'黑体',size:16},paragraph:{keepNext:true,keepLines:true}},
  {scope:'indices',indices:[14,11],paragraph:{keepNext:true}},
])
const after = await inspect(result.bytes)
assert.deepEqual(after.paragraphs.map(p=>p.text),data.paragraphs.map(p=>p.text))
const body = readEditorDocument(result.bytes)
for(const {i,text} of prefixes) {
  const p = body.paragraphs[i]
  assert.equal(findTextRanges(p,text)[0].font.name,'黑体')
  assert.equal(findTextRanges(p,p.text.slice(text.length))[0].font.name,'仿宋')
}
for(const i of [16,18,20,22,24,26]) assert.equal(after.paragraphs[i].pagination.keepNext,true)
await mkdir('data/validation/layout',{recursive:true})
await writeFile('data/validation/layout/公文排版能力验收.docx',result.bytes)
console.log(JSON.stringify({mixedHeadings:5,independentHeadings:6,textPreserved:true,validation:result.validation}))
