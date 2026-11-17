import {readFile,writeFile,mkdir} from 'node:fs/promises'
import assert from 'node:assert/strict'
import {format,unpack,inspect} from '../server/engine.mjs'
import {xml,W} from '../server/document/package.mjs'
const source=await readFile('data/validation/object-flow-real/系统处理回放.docx')
const result=await format(source,[{scope:'footerPageNumber',sectionIndices:[6],alignment:'Centered',font:{name:'Times New Roman',size:10.5}}])
const a=unpack(source),b=unpack(result.bytes)
for(const n of Object.keys(a))if(n!=='word/footer7.xml')assert.deepEqual(b[n],a[n],n)
const f=xml(b['word/footer7.xml'])
assert.equal(f.getElementsByTagNameNS(W,'jc')[0].getAttributeNS(W,'val'),'center')
assert.ok(Array.from(f.getElementsByTagNameNS(W,'sz')).every(n=>n.getAttributeNS(W,'val')==='21'))
assert.ok(Array.from(f.getElementsByTagNameNS(W,'rFonts')).every(n=>n.getAttributeNS(W,'ascii')==='Times New Roman'))
await assert.rejects(format(source,[{scope:'footerPageNumber',sectionIndices:[0],alignment:'Centered'}]),/没有可格式化/)
const d=await inspect(result.bytes),end=d.paragraphs.findIndex(p=>p.text.startsWith('起止时间：'))
assert.ok(end>0&&end<20)
const front=d.paragraphs.slice(0,end+1),ids=front.filter(p=>/^(学生姓名|学\s*号|学\s*院|专业班级|指导教师|起止时间)/.test(p.text)).map(p=>p.index)
const title=front.filter(p=>p.text.startsWith('题    目：')||p.text==='的设计与实现').map(p=>p.index)
// Explicit test-only layout parameters, not a universal academic template.
const output=await format(result.bytes,[
  {scope:'indices',indices:[0],font:{name:'宋体',size:16},paragraph:{alignment:'Centered',lineSpacingMultiple:1.2,spaceAfter:24}},
  {scope:'indices',indices:[front.find(p=>p.text==='本科毕业论文（设计）').index],paragraph:{alignment:'Centered',lineSpacingMultiple:1.2,spaceBefore:18,spaceAfter:36,keepNext:true}},
  {scope:'indices',indices:title,font:{name:'黑体',latinName:'Times New Roman',size:18},paragraph:{alignment:'Centered',lineSpacingMultiple:1.5,spaceAfter:12,keepNext:true}},
  {scope:'indices',indices:ids,font:{name:'宋体',latinName:'Times New Roman',size:14},paragraph:{leftIndent:54,firstLineIndent:0,lineSpacingMultiple:1.5,spaceAfter:8}},
])
await mkdir('data/validation/cover-footer',{recursive:true})
await writeFile('data/validation/cover-footer/封面页码验收.docx',output.bytes)
console.log('真实页码字体、居中、非目标部件保持检查通过；封面验收副本已保存')
