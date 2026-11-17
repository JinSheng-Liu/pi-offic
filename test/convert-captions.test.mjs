import test from 'node:test'
import assert from 'node:assert/strict'
import {paragraphDocument} from './helpers/docx-fixture.mjs'
import {zipSync,unzipSync} from 'fflate'
import {format,inspect} from '../server/engine.mjs'
import {xml,W} from '../server/document/package.mjs'
async function fixture(){
  const parts=unzipSync(paragraphDocument(['正文保持','图3.1 系统结构','图3.2 流程']))
  parts['word/document.xml']=new TextEncoder().encode(new TextDecoder().decode(parts['word/document.xml']).replace('<w:t>图3.1 系统结构</w:t>','<w:t>图3.</w:t></w:r><w:bookmarkStart w:id="44" w:name="_RefCaption"/><w:r><w:t>1</w:t></w:r><w:bookmarkEnd w:id="44"/><w:r><w:t> 系统结构</w:t>'))
  return zipSync(parts)
}
test('existing caption numbers become SEQ fields without changing text, bookmarks or other parts',async()=>{
  const source=await fixture(),before=await inspect(source)
  const out=await format(source,[{scope:'convertCaptions',indices:[1,2]}]),after=await inspect(out.bytes)
  assert.deepEqual(after.paragraphs.map(p=>p.text),before.paragraphs.map(p=>p.text))
  assert.equal(after.captions.filter(c=>c.numbering==='sequence-field').length,2)
  const a=unzipSync(source),b=unzipSync(out.bytes)
  for(const name of Object.keys(a))if(name!=='word/document.xml')assert.deepEqual(b[name],a[name])
  const d=xml(b['word/document.xml'])
  assert.equal(d.getElementsByTagNameNS(W,'bookmarkStart')[0].getAttributeNS(W,'name'),'_RefCaption')
  assert.equal(d.getElementsByTagNameNS(W,'instrText').length,2)
  assert.match(d.getElementsByTagNameNS(W,'instrText')[0].textContent,/\\r 1/)
  assert.doesNotMatch(d.getElementsByTagNameNS(W,'instrText')[1].textContent,/\\r/)
  await assert.rejects(format(out.bytes,[{scope:'convertCaptions',indices:[1,2]}]),/已有域/)
  await assert.rejects(format(source,[{scope:'convertCaptions',indices:[1]}]),/完整选择/)
})
test('noncontinuous or ambiguous numbering is rejected without renumbering content',async()=>{
  const a=unzipSync(await fixture())
  a['word/document.xml']=new TextEncoder().encode(new TextDecoder().decode(a['word/document.xml']).replace('图3.2 流程','图3.4 流程'))
  await assert.rejects(format(zipSync(a),[{scope:'convertCaptions',indices:[1,2]}]),/连续递增/)
})

test('split and quoted existing SEQ instructions cannot collide with conversion',async()=>{
  for(const field of [
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> SE</w:instrText></w:r><w:r><w:instrText>Q "PiFigure3" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>',
    '<w:fldSimple w:instr=" SEQ PiFigure3 "><w:r><w:t>1</w:t></w:r></w:fldSimple>'
  ]){
    const parts=unzipSync(await fixture())
    parts['word/document.xml']=new TextEncoder().encode(new TextDecoder().decode(parts['word/document.xml']).replace('</w:body>',`<w:p>${field}</w:p></w:body>`))
    await assert.rejects(format(zipSync(parts),[{scope:'convertCaptions',indices:[1,2]}]),/同名编号序列/)
  }
})

test('multi-run caption number retains text and run formatting',async()=>{
  const parts=unzipSync(paragraphDocument(Array.from({length:12},(_,i)=>`图3.${i+1} 图片说明`)))
  parts['word/document.xml']=new TextEncoder().encode(new TextDecoder().decode(parts['word/document.xml']).replace('<w:t>图3.12 图片说明</w:t>','<w:t>图3.1</w:t></w:r><w:r w:rsidR="12345678"><w:rPr><w:b/></w:rPr><w:t>2 图片说明</w:t>'))
  const source=zipSync(parts),out=await format(source,[{scope:'convertCaptions',indices:Array.from({length:12},(_,i)=>i)}])
  const d=xml(unzipSync(out.bytes)['word/document.xml']),p=d.getElementsByTagNameNS(W,'p')[11]
  assert.equal(Array.from(p.getElementsByTagNameNS(W,'t')).map(t=>t.textContent).join(''),'图3.12 图片说明')
  const two=Array.from(p.getElementsByTagNameNS(W,'t')).find(t=>t.textContent==='2')
  assert.ok(two)
  assert.equal(two.parentNode.getAttributeNS(W,'rsidR'),'12345678')
  assert.equal(two.parentNode.getElementsByTagNameNS(W,'b').length,1)
  assert.equal((await inspect(out.bytes)).captions.filter(c=>c.numbering==='sequence-field').length,12)
})
