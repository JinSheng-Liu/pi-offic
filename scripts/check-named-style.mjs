import {readFile,writeFile,mkdir} from 'node:fs/promises'
import {format} from '../server/engine.mjs'
import {readEditorDocument} from '../server/document/editor-reader.mjs'
import {digest} from '../server/document/package.mjs'
const source='data/documents/ea68ac07-cf3b-495a-9772-608a38afc15f/original.docx'
const bytes=await readFile(source),before=readEditorDocument(bytes)
const indices=before.paragraphs.flatMap((p,i)=>p.text.trim()&&!p.generatedToc&&!p.tableNestingLevel?[i]:[]).slice(0,3)
const result=await format(bytes,[{scope:'namedStyle',indices,styleName:'验证·命名样式',font:{name:'宋体',latinName:'Times New Roman',size:12},paragraph:{lineSpacingMultiple:1.5,spaceBefore:0,spaceAfter:0}}])
const after=readEditorDocument(result.bytes)
await mkdir('output/named-style-regression',{recursive:true})
await writeFile('output/named-style-regression/命名样式验证.docx',result.bytes)
const report={sourceHash:digest(bytes),originalUnchanged:digest(await readFile(source))===digest(bytes),paragraphCount:before.paragraphs.length,textPreserved:before.text===after.text,indices,styles:indices.map(i=>({index:i,id:after.paragraphs[i].styleId,name:after.paragraphs[i].styleName})),validation:result.validation}
await writeFile('output/named-style-regression/report.json',JSON.stringify(report,null,2),'utf8')
console.log(JSON.stringify(report,null,2))
