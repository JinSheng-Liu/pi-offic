import test from 'node:test'
import assert from 'node:assert/strict'
import {readFileSync} from 'node:fs'
import {EditorRequests} from '../server/editor-requests.mjs'
import {prepareNativeToc,validateNativeToc} from '../server/document/native-toc.mjs'
import {prepareThreeLineTable,applyThreeLineTableOOXML,validateThreeLineTable} from '../server/document/native-table-borders.mjs'
import {documentFixture,richParagraph,table} from './helpers/docx-fixture.mjs'
import {inspect} from '../server/engine.mjs'

test('editor request is scoped, consumed once and cancelled with its task', async()=>{
 const requests=new EditorRequests(),abort=new AbortController();let sent
 const detach=requests.attach('doc',value=>sent=value,abort.signal)
 const result=requests.request('doc',{revision:1},Buffer.from('in'))
 assert.throws(()=>requests.get('other',sent.token))
 requests.complete('doc',sent.token,Buffer.from('out'))
 assert.equal((await result).toString(),'out')
 assert.throws(()=>requests.complete('doc',sent.token,Buffer.from('again')))
 const cancelled=requests.request('doc',{revision:1},Buffer.from('in'))
 abort.abort();await assert.rejects(cancelled,/取消/);detach()
})
test('manual TOC replacement removes only matching contiguous entries',async()=>{
 const bytes=documentFixture(richParagraph('目录')+richParagraph('1. 章节甲')+richParagraph('2. 章节乙')+richParagraph('一、章节甲','<w:pStyle w:val="Heading1"/><w:outlineLvl w:val="0"/>')+richParagraph('保留正文')+richParagraph('二、章节乙','<w:pStyle w:val="Heading1"/><w:outlineLvl w:val="0"/>'))
 const prepared=await prepareNativeToc(bytes,{mode:'replace',indices:[1,2]})
 assert.deepEqual((await inspect(prepared.bytes)).paragraphs.map(p=>p.text),['目录','一、章节甲','保留正文','二、章节乙'])
 await assert.rejects(prepareNativeToc(bytes,{mode:'replace',indices:[3,4]}),/对应/)
 assert.throws(()=>validateNativeToc(prepared.bytes,prepared.bytes),/无效/)
})

test('structural audit follows actual document requirements and rejects static substitutes', async()=>{
 const {structuralTargets,auditStructures}=await import('../server/document/structural-audit.mjs')
 const {format}=await import('../server/engine.mjs')
 const source=documentFixture(richParagraph('问题\t形成原因\t影响')+richParagraph('甲\t乙\t丙'))
 const targets=await structuralTargets(source,'帮我整理这个文档')
 assert.equal(targets.length,1)
 assert.equal((await auditStructures(source,targets))[0].passed,false)
 const result=await format(source,[{scope:'tabularText',indices:[0,1]}])
 assert.equal((await auditStructures(result.bytes,targets))[0].passed,true)
 assert.deepEqual(await structuralTargets(source,'只检查，不要修改'),[])
})

test('three-line table preparation assigns a stable paragraph anchor and validates only editor output',()=>{
 const source=documentFixture(table([['表头一','表头二'],['内容一','内容二']]))
 const prepared=prepareThreeLineTable(source,0)
 assert.match(prepared.paraId,/^[0-9A-F]{8}$/)
 assert.throws(()=>validateThreeLineTable(prepared.bytes,prepared.bytes,0),/边框回读失败/)
 assert.throws(()=>prepareThreeLineTable(source,1),/目标不存在/)
})

test('three-line table OOXML fallback writes and validates a real three-line table',()=>{
 const source=documentFixture(table([['表头一','表头二'],['内容一','内容二']]))
 const prepared=prepareThreeLineTable(source,0)
 const output=applyThreeLineTableOOXML(prepared.bytes,0)
 const validation=validateThreeLineTable(prepared.bytes,output,0)
 assert.equal(validation.status,'passed')
})
