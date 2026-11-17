import fs from 'node:fs'
import assert from 'node:assert/strict'
import {zipSync,unzipSync} from 'fflate'
import {DOMParser,XMLSerializer} from '@xmldom/xmldom'
import {createServerAutomationHost} from '@docx-editor.dev/core/automation'
const out='output/editor-only-probe/boundaries';fs.mkdirSync(out,{recursive:true})
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',enc=s=>new TextEncoder().encode(s)
const fixture=zipSync({
 '[Content_Types].xml':enc('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
 '_rels/.rels':enc('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
 'word/document.xml':enc(`<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="ＭＳ 明朝"/><w:sz w:val="22"/></w:rPr><w:t>中文正文 Spring Boot 2026</w:t></w:r></w:p><w:p><w:r><w:t>未修改内容</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`)
})
const opened=createServerAutomationHost(fixture);assert.ok(opened.ok);const h=opened.host
const q=op=>{const r=h.execute({operations:[op]});assert.ok(r.ok,JSON.stringify(r));return r.results[0].value}
const body=q({op:'getBody',document:q({op:'getDocument'}).handle}).handle,p=q({op:'getParagraphs',body}).handles[0]
q({op:'setFont',span:{paragraph:p},font:{name:'宋体',size:12}})
const read=b=>new DOMParser().parseFromString(new TextDecoder().decode(unzipSync(b)['word/document.xml']),'application/xml')
const slots=d=>Array.from(d.getElementsByTagNameNS(W,'rFonts')).map(n=>Object.fromEntries(['ascii','hAnsi','eastAsia'].map(k=>[k,n.getAttributeNS(W,k)])))
const saved=h.save();assert.ok(saved.ok)
const rawSlots=slots(read(saved.bytes));assert.ok(rawSlots.some(f=>f.eastAsia==='ＭＳ 明朝'))
const current=h.revision()
const stale=h.execute({expectedRevision:current-1,operations:[{op:'setFont',span:{paragraph:p},font:{size:30}}]})
assert.equal(stale.ok,false);assert.equal(h.revision(),current)
const failed=h.execute({expectedRevision:current,operations:[{op:'setFont',span:{paragraph:p},font:{size:30}},{op:'setStyle',span:{paragraph:p},name:'STYLE_NOT_DEFINED'}]})
assert.equal(failed.ok,false);assert.equal(h.revision(),current);assert.equal(q({op:'getFont',span:{paragraph:p}}).font.size,12)
// Narrow fallback proof only: set the three font slots on a known fixture node.
const parts=unzipSync(saved.bytes),xml=read(saved.bytes)
for(const n of Array.from(xml.getElementsByTagNameNS(W,'rFonts'))){for(const [k,v] of Object.entries({eastAsia:'宋体',ascii:'Times New Roman',hAnsi:'Times New Roman'}))n.setAttributeNS(W,'w:'+k,v)}
parts['word/document.xml']=enc(new XMLSerializer().serializeToString(xml))
const patched=zipSync(parts),again=createServerAutomationHost(patched);assert.ok(again.ok)
const exported=again.host.save();assert.ok(exported.ok)
const patchedSlots=slots(read(exported.bytes));assert.ok(patchedSlots.every(f=>f.eastAsia==='宋体'&&f.ascii==='Times New Roman'&&f.hAnsi==='Times New Roman'))
fs.writeFileSync(`${out}/mixed-fonts.docx`,exported.bytes)
const report={rawSlots,staleRevisionRejected:!stale.ok,failedBatchAtomic:!failed.ok,patchedSlots,conclusion:'Native name write retains existing East Asian font; narrow OOXML slot supplement survives editor reopen/save. This does not prove visual font availability.'}
fs.writeFileSync(`${out}/report.json`,JSON.stringify(report,null,2),'utf8');console.log(JSON.stringify(report,null,2));h.dispose();again.host.dispose()
