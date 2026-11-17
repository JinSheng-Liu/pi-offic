import assert from 'node:assert/strict'
import {mkdir,readFile,writeFile} from 'node:fs/promises'
import {createBlankDocument,executeAutomationBatch} from '../server/document/editor-automation.mjs'
import {format,inspect} from '../server/engine.mjs'
const directory=new URL('../data/validation/',import.meta.url)
await mkdir(directory,{recursive:true})
const source=executeAutomationBatch(createBlankDocument(),host=>{
  const ask=op=>host.execute({operations:[op]}).results[0].value
  const document=ask({op:'getDocument'}).handle,body=ask({op:'getBody',document}).handle
  return [{op:'replaceStoryBlocks',body,paragraphs:['格式测试标题 DOCX Studio','正文保持原样：中文 English 123。']}]
}).bytes
const original=await format(source,[{scope:'indices',indices:[1],font:{size:11,italic:true}}])
const before=await inspect(original.bytes)
const originalPath=new URL('engine-original.docx',directory),resultPath=new URL('engine-formatted.docx',directory)
await writeFile(originalPath,original.bytes)
const result=await format(original.bytes,[{scope:'indices',indices:[0],font:{bold:true,size:18},paragraph:{alignment:'Centered'}}])
await writeFile(resultPath,result.bytes)
const after=await inspect(await readFile(resultPath))
assert.deepEqual(after.paragraphs.map(p=>p.text),before.paragraphs.map(p=>p.text))
assert.equal(after.paragraphs[0].font.bold,true)
assert.equal(after.paragraphs[0].font.size,18)
assert.equal(after.paragraphs[0].paragraph.alignment,'Centered')
assert.deepEqual(after.paragraphs[1],before.paragraphs[1])
assert.deepEqual(await readFile(originalPath),original.bytes)
console.log(JSON.stringify({status:'passed',engine:'@docx-editor.dev/core@2.33.0',paragraphs:after.paragraphs.length,textPreserved:true,targetFormattingVerified:true,untouchedFormattingVerified:true,originalUnchanged:true,limitation:'合成夹具；未验证 Word/WPS 页面视觉效果'},null,2))
