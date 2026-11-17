import test from 'node:test'
import assert from 'node:assert/strict'
import {documentFixture,table} from './helpers/docx-fixture.mjs'
import {zipSync,unzipSync} from 'fflate'
import {format,inspect} from '../server/engine.mjs'
import {xml,W} from '../server/document/package.mjs'

test('automatic row height replaces residual height only on selected rows',async()=>{
  const parts=unzipSync(documentFixture(table([['字段','说明'],['编号','内容']])))
  parts['word/document.xml']=new TextEncoder().encode(new TextDecoder().decode(parts['word/document.xml']).replaceAll('<w:tr>','<w:tr><w:trPr><w:trHeight w:val="433" w:hRule="exact"/></w:trPr>'))
  const source=zipSync(parts),before=await inspect(source)
  const out=await format(source,[{scope:'tableRows',tableIndex:0,rowIndices:[0],table:{autoHeight:true}}])
  const heights=xml(unzipSync(out.bytes)['word/document.xml']).getElementsByTagNameNS(W,'trHeight')
  assert.equal(heights[0].getAttributeNS(W,'val'),'0')
  assert.equal(heights[0].getAttributeNS(W,'hRule'),'auto')
  assert.equal(heights[1].getAttributeNS(W,'val'),'433')
  assert.equal(heights[1].getAttributeNS(W,'hRule'),'exact')
  assert.deepEqual((await inspect(out.bytes)).paragraphs.map(p=>p.text),before.paragraphs.map(p=>p.text))
})

test('native repeated headers can be enabled then removed without changing row height', async () => {
  const source = documentFixture(table([['标题','说明'],['内容','文字']]))
  const enabled = await format(source, [{scope:'tableRows',tableIndex:0,rowIndices:[0],table:{repeatHeader:true}}])
  assert.equal((await inspect(enabled.bytes)).tables[0].rowDetails[0].header, true)
  const disabled = await format(enabled.bytes, [{scope:'tableRows',tableIndex:0,rowIndices:[0],table:{repeatHeader:false}}])
  assert.equal((await inspect(disabled.bytes)).tables[0].rowDetails[0].header, false)
  assert.deepEqual((await inspect(disabled.bytes)).paragraphs.map(p=>p.text),(await inspect(source)).paragraphs.map(p=>p.text))
})

test('native table layout retains cell font and unrelated paragraph properties', async () => {
  const source = documentFixture(table([['标题','说明'],['编号','内容']]))
  const styled = await format(source, [{scope:'tableText',font:{name:'宋体',latinName:'Times New Roman',size:10.5},paragraph:{alignment:'Left',spaceAfter:3}}])
  const before = await inspect(styled.bytes)
  const result = await format(styled.bytes,[{scope:'tableRows',tableIndex:0,rowIndices:[0],table:{repeatHeader:true,cantSplit:true,autoHeight:true}},
    {scope:'tableCells',tableIndex:0,cellTargets:[{rowIndex:1,cellIndex:0}],table:{verticalAlignment:'Center'}}])
  const after = await inspect(result.bytes)
  assert.deepEqual(after.paragraphs, before.paragraphs)
})
