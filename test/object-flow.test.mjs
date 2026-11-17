import test from 'node:test'
import assert from 'node:assert/strict'
import {paragraphDocument,withInlinePicture} from './helpers/docx-fixture.mjs'
import {zipSync,unzipSync} from 'fflate'
import {format,inspect} from '../server/engine.mjs'
import {xml,W,A} from '../server/document/package.mjs'
import {XMLSerializer} from '@xmldom/xmldom'
import {WP} from '../server/document/object-flow.mjs'
import {findMixedTocTitles} from '../server/document/toc-layout.mjs'

async function fixture(mismatch=false) {
  const parts=unzipSync(paragraphDocument(['保留正文']))
  const box=text=>`<w:txbxContent><w:p><w:pPr/><w:r><w:t>${text}</w:t></w:r></w:p><w:p><w:pPr/></w:p></w:txbxContent>`
  const shape=`<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" Requires="wps"><w:drawing><wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:wrapNone/>${box('表1 结果')}</wp:anchor></w:drawing></mc:Choice><mc:Fallback>${box(mismatch?'不同题注':'表1 结果')}</mc:Fallback></mc:AlternateContent>`
  parts['word/document.xml']=new TextEncoder().encode(new TextDecoder().decode(parts['word/document.xml']).replace('<w:body>',`<w:body><w:p><w:r>${shape}</w:r></w:p><w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:p><w:r><w:t>值</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`))
  return zipSync(parts)
}
test('floating caption moves into flow once, preserving other content and rejecting divergent fallback',async()=>{
  const source=await fixture()
  const result=await format(source,[{scope:'floatingObjects',objectIndices:[0],objectLayout:'captionInFlow'}])
  const data=await inspect(result.bytes)
  assert.equal(data.layoutRisks.floatingCount,0)
  assert.equal(data.paragraphs.filter(p=>p.text==='表1 结果').length,1)
  assert.ok(data.paragraphs.some(p=>p.text==='保留正文'))
  assert.equal(data.tables.length,1)
  assert.equal(data.paragraphs.length,(await inspect(source)).paragraphs.length+1)
  await assert.rejects(format(await fixture(true),[{scope:'floatingObjects',objectIndices:[0],objectLayout:'captionInFlow'}]),/兼容分支/)
  const parts=unzipSync(source)
  parts['word/document.xml']=new TextEncoder().encode(new TextDecoder().decode(parts['word/document.xml']).replaceAll('<w:p><w:pPr/></w:p></w:txbxContent>','<w:p><w:r><w:t>表0 历史题注</w:t></w:r></w:p></w:txbxContent>'))
  await assert.rejects(format(zipSync(parts),[{scope:'floatingObjects',objectIndices:[0],objectLayout:'captionInFlow'}]),/累积题注/)
})
test('floating picture becomes inline without changing media, extent or surrounding text',async()=>{
  const parts=unzipSync(withInlinePicture()),doc=xml(parts['word/document.xml'])
  const run=doc.getElementsByTagNameNS(W,'r')[0],text=doc.createElementNS(W,'w:t');text.textContent='校名';run.insertBefore(text,run.firstChild)
  const inline=doc.getElementsByTagNameNS(WP,'inline')[0],anchor=doc.createElementNS(WP,'wp:anchor')
  for(const child of Array.from(inline.childNodes))anchor.appendChild(child.cloneNode(true))
  inline.parentNode.replaceChild(anchor,inline)
  parts['word/document.xml']=new TextEncoder().encode(new XMLSerializer().serializeToString(doc))
  const out=await format(zipSync(parts),[{scope:'floatingObjects',objectIndices:[0],objectLayout:'inlinePicture'}])
  const saved=unzipSync(out.bytes),after=xml(saved['word/document.xml'])
  assert.equal(after.getElementsByTagNameNS(WP,'anchor').length,0)
  assert.equal(after.getElementsByTagNameNS(WP,'inline').length,1)
  assert.equal(after.getElementsByTagNameNS(W,'t')[0].textContent,'校名')
  assert.equal(after.getElementsByTagNameNS(WP,'extent')[0].getAttribute('cx'),doc.getElementsByTagNameNS(WP,'extent')[0].getAttribute('cx'))
  assert.equal(after.getElementsByTagNameNS(A,'blip').length,1)
  for(const name of Object.keys(parts).filter(n=>n!=='word/document.xml'))assert.deepEqual(saved[name],parts[name])
})

test('table caption anchored in explanatory text moves before adjacent table without moving body text',async()=>{
  const parts=unzipSync(await fixture())
  const source=new TextDecoder().decode(parts['word/document.xml']).replace('</mc:AlternateContent></w:r></w:p>','</mc:AlternateContent></w:r><w:r><w:t>说明正文，如表1所示。</w:t></w:r></w:p>')
  parts['word/document.xml']=new TextEncoder().encode(source)
  const result=await format(zipSync(parts),[{scope:'floatingObjects',objectIndices:[0],objectLayout:'captionInFlow'}])
  const doc=xml(unzipSync(result.bytes)['word/document.xml'])
  const blocks=Array.from(doc.getElementsByTagNameNS(W,'body')[0].childNodes).filter(n=>n.nodeType===1)
  assert.equal(blocks[0].textContent,'说明正文，如表1所示。')
  assert.equal(blocks[1].textContent,'表1 结果')
  assert.equal(blocks[2].localName,'tbl')
  parts['word/document.xml']=new TextEncoder().encode(source.replace('<w:tbl>','<w:p><w:pPr/></w:p><w:tbl>'))
  const spaced=await format(zipSync(parts),[{scope:'floatingObjects',objectIndices:[0],objectLayout:'captionInFlow'}])
  const spacedDoc=xml(unzipSync(spaced.bytes)['word/document.xml'])
  const table=spacedDoc.getElementsByTagNameNS(W,'tbl')[0]
  assert.equal(table.previousSibling.textContent,'表1 结果')
  assert.equal(table.previousSibling.previousSibling.textContent,'')
  parts['word/document.xml']=new TextEncoder().encode(source.replace('<w:tbl>','<w:p><w:r><w:br w:type="page"/></w:r></w:p><w:tbl>'))
  await assert.rejects(format(zipSync(parts),[{scope:'floatingObjects',objectIndices:[0],objectLayout:'captionInFlow'}]),/宿主|相邻/)
  parts['word/document.xml']=new TextEncoder().encode(source.replace('<w:tbl>','<w:p><w:r><w:t>无关正文</w:t></w:r></w:p><w:tbl>'))
  await assert.rejects(format(zipSync(parts),[{scope:'floatingObjects',objectIndices:[0],objectLayout:'captionInFlow'}]),/宿主|相邻/)
})
test('TOC title separation preserves field instructions and page tabs preserve text',async()=>{
  const parts=unzipSync(paragraphDocument(['正文']))
  parts['word/document.xml']=new TextEncoder().encode(new TextDecoder().decode(parts['word/document.xml']).replace('<w:body>','<w:body><w:sdt><w:sdtPr/><w:sdtContent><w:p><w:pPr/><w:r><w:t>目录</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \\o "1-3"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p><w:p><w:r><w:t>第一章</w:t><w:tab/><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:sdtContent></w:sdt>'))
  const source=zipSync(parts),out=await format(source,[{scope:'tocTitle',indices:[0]}])
  assert.deepEqual(findMixedTocTitles(xml(unzipSync(source)['word/document.xml'])),[0])
  assert.deepEqual(findMixedTocTitles(xml(unzipSync(out.bytes)['word/document.xml'])),[])
  const doc=xml(unzipSync(out.bytes)['word/document.xml'])
  const first=Array.from(doc.getElementsByTagNameNS(W,'body')[0].childNodes).find(n=>n.nodeType===1)
  assert.equal(first.localName,'p')
  assert.equal(first.getElementsByTagNameNS(W,'t')[0].textContent,'目录')
  assert.equal(doc.getElementsByTagNameNS(W,'instrText')[0].textContent,'TOC \\o "1-3"')
  assert.equal(doc.getElementsByTagNameNS(W,'fldChar').length,3)
  const data=await inspect(out.bytes),entry=data.paragraphs.find(p=>p.text==='第一章\t1')
  const tabs=await format(out.bytes,[{scope:'indices',indices:[entry.index],paragraph:{tocPageTabPosition:400}}])
  const after=await inspect(tabs.bytes)
  assert.deepEqual(after.paragraphs.map(p=>p.text),data.paragraphs.map(p=>p.text))
  assert.deepEqual(after.paragraphs[entry.index].tabStops.direct,[{positionTwips:8000,alignment:'right',leader:'dot'}])
  await assert.rejects(format(out.bytes,[{scope:'tocTitle',indices:[0]}]),/内容控件/)
})

