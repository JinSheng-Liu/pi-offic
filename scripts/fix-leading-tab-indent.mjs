import {readFile,writeFile,mkdir} from 'node:fs/promises'
import assert from 'node:assert/strict'
import {readEditorDocument} from '../server/document/editor-reader.mjs'
import {format,inspect,unpack,digest} from '../server/engine.mjs'
const root='data/documents/3501e5c4-ed8d-4857-96e4-294b79ce0146'
const metadata=JSON.parse(await readFile(`${root}/metadata.json`,'utf8'))
const source=await readFile(`${root}/${metadata.versions.at(-1).file}`)
const body=readEditorDocument(source)
const targets=body.paragraphs.flatMap((p,index)=>{
  const beginsWithTab=p.text.startsWith('\t')
  const first=p.firstLineIndent,left=p.leftIndent
  return beginsWithTab&&p.styleId==='Normal'&&/^\t[^\s（(]/u.test(p.text)&&first>0&&left===0?[{index,text:p.text.slice(0,55),first,left}]:[]
})
assert.ok(targets.length)
const output=await format(source,[{scope:'indices',indices:targets.map(t=>t.index),paragraph:{firstLineIndent:0}}])
const before=await inspect(source),after=await inspect(output.bytes)
assert.deepEqual(after.paragraphs.map(p=>p.text),before.paragraphs.map(p=>p.text))
for(const {index} of targets){assert.equal(after.paragraphs[index].paragraph.firstLineIndent,0);assert.equal(after.paragraphs[index].paragraph.leftIndent,0)}
const old=unpack(source),next=unpack(output.bytes)
for(const name of Object.keys(old).filter(name=>name.startsWith('word/media/')&&!name.endsWith('/')))assert.equal(digest(next[name]),digest(old[name]))
const folder='data/validation/indent';await mkdir(folder,{recursive:true})
await writeFile(`${folder}/论文_首行缩进修正版.docx`,output.bytes)
await writeFile(`${folder}/report.json`,JSON.stringify({sourceRevision:metadata.revision,correctedParagraphs:targets,paragraphCount:after.paragraphs.length,textPreserved:true,mediaPreserved:Object.keys(old).filter(name=>name.startsWith('word/media/')&&!name.endsWith('/')).length},null,2))
console.log(JSON.stringify({corrected:targets.length,targets}))
