import test from 'node:test'
import assert from 'node:assert/strict'
import {paragraphDocument} from './helpers/docx-fixture.mjs'
import {unzipSync,zipSync} from 'fflate'
import {mkdtemp,rm} from 'node:fs/promises'
import {tmpdir} from 'node:os'
import {join} from 'node:path'
import {restoreReferenceCover} from '../server/document/reference-cover.mjs'
import {DocumentStore} from '../server/store.mjs'
import {xml,W,canonical} from '../server/document/package.mjs'
import {auditFrontMatter} from '../server/document/front-matter-audit.mjs'
async function fixture(){
  const parts=unzipSync(paragraphDocument(['学生姓名：张三','学号：123','正文保持'])),s=new TextDecoder().decode(parts['word/document.xml'])
  parts['word/document.xml']=new TextEncoder().encode(s.replace('<w:p>','<w:p><w:pPr><w:sectPr/></w:pPr>'))
  return parts
}
test('reference cover preserves body and parts, refuses different personal data and records a version',async()=>{
  const parts=await fixture(),target=zipSync(parts)
  const changed={...parts,'word/document.xml':new TextEncoder().encode(new TextDecoder().decode(parts['word/document.xml']).replace('<w:r>','<w:r><w:rPr><w:sz w:val="36"/><w:u w:val="single"/></w:rPr>'))}
  const reference=zipSync(changed),out=restoreReferenceCover(target,reference)
  const a=xml(parts['word/document.xml']),b=xml(unzipSync(out.bytes)['word/document.xml'])
  assert.equal(b.getElementsByTagNameNS(W,'sz')[0].getAttributeNS(W,'val'),'36')
  assert.equal(canonical(a.getElementsByTagNameNS(W,'p')[2]),canonical(b.getElementsByTagNameNS(W,'p')[2]))
  for(const name of Object.keys(parts))if(name!=='word/document.xml')assert.deepEqual(unzipSync(out.bytes)[name],parts[name])
  const mismatch={...changed,'word/document.xml':new TextEncoder().encode(new TextDecoder().decode(changed['word/document.xml']).replace('张三','李四'))}
  assert.throws(()=>restoreReferenceCover(target,zipSync(mismatch)),/文字不同/)
  const root=await mkdtemp(join(tmpdir(),'reference-cover-'))
  try{
    const store=new DocumentStore(root),doc=await store.create('目标.docx',target),ref=await store.create('参考.docx',reference)
    const result=await store.workflow.execute(doc.id,1,{kind:'referenceCover',referenceId:ref.id,referenceRevision:1})
    assert.equal(result.document.revision,2)
    assert.equal(result.validation.operationReceipt.operations[0].referenceId,ref.id)
    assert.equal(result.validation.operationReceipt.kind,'referenceCover')
  }finally{await rm(root,{recursive:true,force:true})}
})

test('declaration audit finds missing formatting and restore leaves following body intact',async()=>{
  const parts=await fixture()
  const prefix='<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>封面</w:t></w:r></w:p>'
  const region='<w:p><w:r><w:t>学位论文原创性声明</w:t></w:r></w:p><w:p><w:r><w:t>本人郑重声明：内容不改。</w:t></w:r></w:p><w:p><w:r><w:t>学位论文作者签名：</w:t></w:r></w:p><w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>日期：2026年5月1日</w:t></w:r></w:p>'
  const source=new TextDecoder().decode(parts['word/document.xml']).replace(/<w:body>[\s\S]*<\/w:body>/,`<w:body>${prefix}${region}<w:p><w:r><w:t>正文保持</w:t></w:r></w:p><w:sectPr/></w:body>`)
  const target=zipSync({...parts,'word/document.xml':new TextEncoder().encode(source)})
  const reference=zipSync({...parts,'word/document.xml':new TextEncoder().encode(source.replace('<w:p><w:r><w:t>学位论文原创性声明','<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>学位论文原创性声明'))})
  assert.equal(auditFrontMatter(target,reference).regions[2].status,'format-differences')
  const result=restoreReferenceCover(target,reference,undefined,'declaration')
  assert.equal(auditFrontMatter(result.bytes,reference).regions[2].status,'reference-properties-matched')
  const before=xml(unzipSync(target)['word/document.xml']),after=xml(unzipSync(result.bytes)['word/document.xml'])
  assert.equal(canonical(before.getElementsByTagNameNS(W,'p')[5]),canonical(after.getElementsByTagNameNS(W,'p')[5]))
})
