import test from 'node:test'
import assert from 'node:assert/strict'
import {paragraphDocument} from './helpers/docx-fixture.mjs'
import {zipSync,unzipSync} from 'fflate'
import {format} from '../server/engine.mjs'
import {xml,W} from '../server/document/package.mjs'
test('formats PAGE inside content controls, preserves fields and refuses inherited shared footer',async()=>{
  const parts=unzipSync(paragraphDocument(['正文'])),decode=n=>new TextDecoder().decode(parts[n]),encode=s=>new TextEncoder().encode(s)
  parts['word/document.xml']=encode(decode('word/document.xml').replace('<w:sectPr','<w:sectPr').replace(/(<w:sectPr[^>]*>)/,'$1<w:footerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="footerTest" w:type="default"/>'))
  parts['word/_rels/document.xml.rels']=encode(decode('word/_rels/document.xml.rels').replace('</Relationships>','<Relationship Id="footerTest" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footerTest.xml"/></Relationships>'))
  parts['word/footerTest.xml']=encode(`<w:ftr xmlns:w="${W}"><w:sdt><w:sdtPr/><w:sdtContent><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:sdtContent></w:sdt></w:ftr>`)
  parts['[Content_Types].xml']=encode(decode('[Content_Types].xml').replace('</Types>','<Override PartName="/word/footerTest.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>'))
  const op={scope:'footerPageNumber',sectionIndices:[0],alignment:'Centered',font:{name:'Times New Roman',size:10.5}}
  const result=unzipSync((await format(zipSync(parts),[op])).bytes),footer=xml(result['word/footerTest.xml'])
  assert.equal(footer.getElementsByTagNameNS(W,'jc')[0].getAttributeNS(W,'val'),'center')
  assert.equal(footer.getElementsByTagNameNS(W,'instrText')[0].textContent,'PAGE')
  assert.equal(footer.getElementsByTagNameNS(W,'sz')[0].getAttributeNS(W,'val'),'21')
  for(const n of Object.keys(parts))if(n!=='word/footerTest.xml')assert.deepEqual(result[n],parts[n])
  parts['word/document.xml']=encode(decode('word/document.xml').replace('<w:body>','<w:body><w:p><w:pPr><w:sectPr><w:footerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="footerTest" w:type="default"/></w:sectPr></w:pPr></w:p>'))
  await assert.rejects(format(zipSync(parts),[op]),/共享/)
})
