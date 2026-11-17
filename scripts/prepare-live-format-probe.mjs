import fs from 'node:fs'
import {createBlankDocument,executeAutomationBatch} from '../server/document/editor-automation.mjs'
import {format} from '../server/engine.mjs'
const out='output/live-format-probe'
fs.mkdirSync(out,{recursive:true})
const source=process.argv[2]?fs.readFileSync(process.argv[2]):executeAutomationBatch(createBlankDocument(),host=>{
 const ask=op=>host.execute({operations:[op]}).results[0].value
 const document=ask({op:'getDocument'}).handle,body=ask({op:'getBody',document}).handle
 return [{op:'replaceStoryBlocks',body,paragraphs:Array.from({length:45},(_,i)=>`第${i+1}段：测试中英文格式 English 123，观察实际行距、字号和缩进变化。`)}]
}).bytes
const formatted=await format(source,[{scope:process.argv[3]||'all',font:{name:'宋体',latinName:'Times New Roman',size:14},paragraph:{lineSpacingMultiple:2,firstLineIndent:28}}])
fs.writeFileSync(`${out}/source.docx`,source)
fs.writeFileSync(`${out}/expected.docx`,formatted.bytes)
fs.writeFileSync(`${out}/replay.json`,JSON.stringify(formatted.replay||null))
console.log({paragraphs:formatted.affected,replay:!!formatted.replay})
