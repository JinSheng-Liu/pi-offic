import test from 'node:test'
import assert from 'node:assert/strict'
import {format} from '../server/engine.mjs'
import {paragraphDocument} from './helpers/docx-fixture.mjs'
import {readEditorDocument} from '../server/document/editor-reader.mjs'
const op={scope:'characterStyle',indices:[0],text:'一、提高统筹认识。',styleName:'公文·段内一级标题',font:{size:18,bold:true}}
test('inline heading gets character style while body paragraph and remaining text stay unchanged',async()=>{
 const input=paragraphDocument([op.text+'后面是正文','其他段落'])
 const result=await format(input,[op]),p=readEditorDocument(result.bytes).paragraphs[0]
 assert.equal(p.styleId,readEditorDocument(input).paragraphs[0].styleId)
 assert.match(p.runs[0].value.rPr.rStyle.val,/^PiChar/)
 assert.equal(p.runs[0].font.size,18)
 assert.equal(p.runs.at(-1).value.rPr?.rStyle,undefined)
 const again=await format(result.bytes,[op]);assert.equal(readEditorDocument(again.bytes).paragraphs[0].runs[0].value.rPr.rStyle.val,p.runs[0].value.rPr.rStyle.val)
})
test('ambiguous inline headings require occurrence',async()=>{
 const input=paragraphDocument([op.text+op.text])
 await assert.rejects(format(input,[op]),/唯一匹配/)
 const result=await format(input,[{...op,occurrence:1}])
 assert.equal(readEditorDocument(result.bytes).paragraphs[0].runs[0].value.rPr?.rStyle,undefined)
})
test('already formatted single run is split only around selected heading',async()=>{
 const input=await format(paragraphDocument([op.text+'正文']),[{scope:'all',font:op.font}])
 const result=await format(input.bytes,[op]),p=readEditorDocument(result.bytes).paragraphs[0]
 assert.equal(p.runs[0].text,op.text)
 assert.match(p.runs[0].value.rPr.rStyle.val,/^PiChar/)
 assert.equal(p.runs.at(-1).text,'正文')
 assert.equal(p.runs.at(-1).value.rPr?.rStyle,undefined)
})
