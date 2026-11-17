import {readFile,writeFile} from 'node:fs/promises'
import {isDeepStrictEqual} from 'node:util'
import {createHash} from 'node:crypto'
import {readEditorDocument} from '../server/document/editor-reader.mjs'
import {inspect,compareProtected} from '../server/engine.mjs'
const dir='data/validation/skills-thesis',task=JSON.parse(await readFile(dir+'/task.json','utf8'))
const original=await readFile(task.source),output=await readFile(dir+'/测试结果.docx')
const before=JSON.parse(await readFile(dir+'/before.json','utf8')),after=await inspect(output)
const body=readEditorDocument(output)
const checks=[];const check=(name,pass,details)=>checks.push({name,pass,...(details?{details}:{})})
check('桌面原件哈希不变',createHash('sha256').update(original).digest('hex')===task.hash)
try{compareProtected(original,output);check('包内容与媒体保护',true)}catch(e){check('包内容与媒体保护',false,e.message)}
check('SDK段落文字不变',isDeepStrictEqual(before.paragraphs.map(p=>p.text),after.paragraphs.map(p=>p.text)))
check('封面大标题30pt居中',after.paragraphs[2].font.size===30&&after.paragraphs[2].paragraph.alignment==='Centered')
check('封面题目两行18pt',[5,6].every(i=>after.paragraphs[i].font.size===18))
const unchanged=i=>isDeepStrictEqual(before.paragraphs[i].font,after.paragraphs[i].font)&&isDeepStrictEqual(before.paragraphs[i].paragraph,after.paragraphs[i].paragraph)
const protectedIndices=[...Array.from({length:28},(_,i)=>i+7),39,45,...Array.from({length:51},(_,i)=>i+47)]
check('信息行声明关键词与目录格式保留',protectedIndices.every(unchanged),protectedIndices.filter(i=>!unchanged(i)))
check('摘要与ABSTRACT标题16pt居中',[35,41].every(i=>after.paragraphs[i].font.size===16&&after.paragraphs[i].paragraph.alignment==='Centered'))
const formatted=i=>{const p=after.paragraphs[i];return p.font.size===12&&p.paragraph.firstLineIndent===24&&body.paragraphs[i].p.pPr?.spacing?.lineRule==='auto'&&body.paragraphs[i].p.pPr?.spacing?.line===360&&p.font.asciiDirect.every(v=>v==='Times New Roman')&&p.font.hAnsiDirect.every(v=>v==='Times New Roman')&&p.font.eastAsiaDirect.every(v=>v==='宋体')}
check('六段摘要正文格式',[36,37,38,42,43,44].every(formatted),[36,37,38,42,43,44].map(i=>({index:i,pass:formatted(i),spacing:body.paragraphs[i].p.pPr?.spacing})))
const headingIndices=before.paragraphs.filter(p=>p.index>=101&&p.index<1144&&/^\d+\.(?:\d+\.)?\d*[^\d\s．]/u.test(p.text)&&p.text.length<70&&!p.tableNestingLevel).map(p=>p.index)
check('识别出的章节标题全部设置层级',headingIndices.every(i=>after.paragraphs[i].heading>0),headingIndices.filter(i=>after.paragraphs[i].heading===0))
const samples=[103,104,105,107,108,109,111,112,134,136,138]
check('正文独立样本格式',samples.every(formatted),samples.map(i=>({index:i,pass:formatted(i)})))
const images=before.paragraphs.filter(p=>p.index>=101&&p.inlinePictureCount&&!p.text.trim()).map(p=>p.index)
check('正文独占嵌入图片段落居中',images.every(i=>after.paragraphs[i].paragraph.alignment==='Centered'),{total:images.length,missed:images.filter(i=>after.paragraphs[i].paragraph.alignment!=='Centered')})
const headerLabels=new Set(['字段名','类型','约束','说明','注释','编号','输入/动作','期望的输出','实际情况']); const heads=before.cells.filter(c=>headerLabels.has(c.text.trim())).flatMap(c=>c.paragraphIndices)
check('明确表头单元格居中',heads.every(i=>after.paragraphs[i].paragraph.alignment==='Centered'),{total:heads.length,missed:heads.filter(i=>after.paragraphs[i].paragraph.alignment!=='Centered'),note:'根据文档中字段表和用例表的明确表头文字核对'})
const other=before.cells.flatMap(c=>c.paragraphIndices).filter(i=>!heads.includes(i))
check('非表头单元格格式未改',other.every(unchanged),other.filter(i=>!unchanged(i)))
const report={id:task.id,checks,passed:checks.filter(c=>c.pass).length,total:checks.length,visualVerified:false}
await writeFile(dir+'/verification.json',JSON.stringify(report,null,2),'utf8');await writeFile(dir+'/after.json',JSON.stringify(after),'utf8')
console.log(JSON.stringify(report,null,2))
