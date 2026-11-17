import test from 'node:test'
import assert from 'node:assert/strict'
import {paragraphDocument} from './helpers/docx-fixture.mjs'
import {unzipSync,zipSync} from 'fflate'
import {format,inspect} from '../server/engine.mjs'
import {xml,W,canonical} from '../server/document/package.mjs'
test('cover layout joins title and aligns fields without changing other content',async()=>{
  const source=paragraphDocument(['本科毕业论文（设计）','题    目：基于Spring Boot的系统','设计与实现','学生姓名：     张三   ','学    号：   123456','正文必须保持']),out=await format(source,[{scope:'coverLayout',indices:[1,2,3,4],cover:{widthPt:420,labelWidthPt:154}}])
  const a=unzipSync(source),b=unzipSync(out.bytes),doc=xml(b['word/document.xml'])
  const table=doc.getElementsByTagNameNS(W,'tbl')[0]
  assert.equal(table.getElementsByTagNameNS(W,'tr').length,3)
  const data=await inspect(out.bytes)
  assert.ok(data.paragraphs.some(p=>p.text==='基于Spring Boot的系统设计与实现'))
  assert.ok(data.paragraphs.some(p=>p.text==='张三'))
  assert.equal(canonical(doc.getElementsByTagNameNS(W,'body')[0].lastChild),canonical(xml(a['word/document.xml']).getElementsByTagNameNS(W,'body')[0].lastChild))
  for(const name of Object.keys(a))if(name!=='word/document.xml')assert.deepEqual(a[name],b[name])
  await assert.rejects(format(source,[{scope:'coverLayout',indices:[1,3],cover:{widthPt:420,labelWidthPt:154}}]),/连续/)
  await assert.rejects(format(out.bytes,[{scope:'coverLayout',indices:[1,2],cover:{widthPt:420,labelWidthPt:154}}]),/顶层/)
})

test('cover keeps final section and trailing bookmark boundaries outside the table',async()=>{
  const parts=unzipSync(paragraphDocument(['学生姓名： 张三','学号： 123','声明页']))
  let source=new TextDecoder().decode(parts['word/document.xml'])
  source=source.replace('<w:p>','<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:left="1440" w:right="1440"/></w:sectPr></w:pPr>')
  // A section on an intermediate selected paragraph is explicitly rejected.
  parts['word/document.xml']=new TextEncoder().encode(source)
  await assert.rejects(format(zipSync(parts),[{scope:'coverLayout',indices:[0,1],cover:{widthPt:420,labelWidthPt:154}}]),/复杂结构/)
  source=source.replace(/<w:pPr>.*?<\/w:pPr>/,'').replace('<w:t>学号： 123</w:t></w:r></w:p>','<w:t>学号： 123</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>')
  source=source.replace('<w:p><w:r><w:t>学号', '<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:left="1440" w:right="1440"/></w:sectPr></w:pPr><w:r><w:t>学号')
  parts['word/document.xml']=new TextEncoder().encode(source)
  const out=await format(zipSync(parts),[{scope:'coverLayout',indices:[0,1],cover:{widthPt:420,labelWidthPt:154}}])
  const doc=xml(unzipSync(out.bytes)['word/document.xml']),table=doc.getElementsByTagNameNS(W,'tbl')[0]
  assert.equal(table.getElementsByTagNameNS(W,'sectPr').length,0)
  assert.equal(table.nextSibling.getElementsByTagNameNS(W,'sectPr').length,1)
  assert.equal(table.nextSibling.getElementsByTagNameNS(W,'bookmarkEnd')[0].getAttributeNS(W,'id'),'7')
})
