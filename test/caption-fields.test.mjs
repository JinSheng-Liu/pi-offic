import test from 'node:test'
import assert from 'node:assert/strict'
import {captionFieldDetails} from '../server/document/caption-fields.mjs'
import {xml,W} from '../server/document/package.mjs'
test('manual labels and Caption style do not imply automatic captions; split SEQ is recognized',()=>{
  const doc=xml(new TextEncoder().encode(`<w:document xmlns:w="${W}"><w:body>
  <w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr><w:r><w:t>图1 手工编号</w:t></w:r></w:p>
  <w:p><w:r><w:fldChar w:fldCharType="begin"/><w:instrText> SE</w:instrText></w:r><w:r><w:instrText>Q Figure </w:instrText><w:fldChar w:fldCharType="separate"/><w:t>1</w:t><w:fldChar w:fldCharType="end"/></w:r></w:p>
  <w:p><w:fldSimple w:instr="SEQ 表"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>
  <w:p><w:r><w:fldChar w:fldCharType="begin"/><w:instrText>SEQ Broken</w:instrText></w:r></w:p>
  </w:body></w:document>`))
  const result=captionFieldDetails(doc)
  assert.deepEqual(result.map(r=>r.numbering),['manual-or-unverified','sequence-field','sequence-field','manual-or-unverified'])
  assert.deepEqual(result[1].sequenceIdentifiers,['Figure'])
  assert.deepEqual(result[2].sequenceIdentifiers,['表'])
})
