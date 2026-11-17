import fs from 'node:fs'
import assert from 'node:assert/strict'
import {unzipSync} from 'fflate'
import {DOMParser} from '@xmldom/xmldom'
import {createServerAutomationHost} from '@docx-editor.dev/core/automation'
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const M='http://schemas.openxmlformats.org/officeDocument/2006/math'
const out='output/editor-only-probe/objects';fs.mkdirSync(out,{recursive:true})
const bytes=fs.readFileSync('data/documents/ea68ac07-cf3b-495a-9772-608a38afc15f/original.docx')
const opened=createServerAutomationHost(bytes);assert.ok(opened.ok)
const h=opened.host,results={}
const q=op=>{const r=h.execute({operations:[op],expectedRevision:h.revision()});if(!r.ok)throw Error(JSON.stringify(r));return r.results[0].value}
const doc=q({op:'getDocument'}).handle,body=q({op:'getBody',document:doc}).handle
const tables=q({op:'getTables',scope:{body}}).handles
results.tables=tables.length
results.repeatHeader=q({op:'updateTable',table:tables[0],mutation:{kind:'properties',headerRowCount:1}})
const cell=q({op:'getTableCell',table:tables[0],rowIndex:1,cellIndex:0}).handle
results.cellVertical=q({op:'updateTableCell',cell,properties:{verticalAlignment:'Center'}})
const cellBody=q({op:'getTableCellBody',cell}).handle
results.cellHorizontal=q({op:'setParagraphFormat',paragraph:{body:cellBody,at:'first'},format:{alignment:'Centered'}})
const sections=q({op:'getSections',document:doc}).handles
results.sections=sections.length
results.pageSetup=q({op:'setPageSetup',section:sections[0],setup:{topMargin:72,bottomMargin:72}})
const shapes=q({op:'getShapes',span:{body},types:['TextBox']})
results.shapesResponseKind=shapes.kind
results.shapesCount=shapes.handles?.length
if(shapes.handles?.length){
 const shapeBody=q({op:'getShapeBody',shape:shapes.handles[0]}).handle
 const before=q({op:'getText',target:shapeBody}).text
 results.textboxFormatting=q({op:'setParagraphFormat',paragraph:{body:shapeBody,at:'first'},format:{alignment:'Centered'}})
 assert.equal(q({op:'getText',target:shapeBody}).text,before)
 results.textboxTextPreserved=true
}
const footer=q({op:'getFurniture',section:sections.at(-1),kind:'footer',variant:'default'}).handle
const footerPs=q({op:'getParagraphs',body:footer}).handles
results.footerParagraphs=footerPs.length
if(footerPs.length)results.footerAlignment=q({op:'setParagraphFormat',paragraph:{paragraph:footerPs[0]},format:{alignment:'Centered'}})
const fields=q({op:'getFields',span:{body}}).handles
results.fields=fields.length
results.fieldRead=fields.length?q({op:'getField',field:fields[0]}):null
const saved=h.save();assert.ok(saved.ok);fs.writeFileSync(`${out}/edited.docx`,saved.bytes)
const parse=b=>new DOMParser().parseFromString(new TextDecoder().decode(b),'application/xml')
const all=(d,n,ns=W)=>Array.from(d.getElementsByTagNameNS(ns,n))
const source=unzipSync(bytes),target=unzipSync(saved.bytes),xml=parse(target['word/document.xml'])
const table=all(xml,'tbl')[0]
results.xml={repeatHeader:all(table,'tblHeader').length,cellAlignment:all(all(table,'tr')[1],'vAlign')[0]?.getAttributeNS(W,'val'),topMargin:all(xml,'pgMar')[0]?.getAttributeNS(W,'top')}
assert.equal(results.xml.repeatHeader,1);assert.equal(results.xml.cellAlignment,'center');assert.equal(results.xml.topMargin,'1440')
results.contentChanges=[]
for(const name of Object.keys(source).filter(n=>/^word\/.*\.xml$/.test(n))){
 const a=parse(source[name]),b=parse(target[name]);
 for(const [ns,tag] of [[W,'t'],[W,'instrText'],[M,'oMath']]){
  const read=d=>all(d,tag,ns).map(x=>x.textContent)
  if(JSON.stringify(read(a))!==JSON.stringify(read(b)))results.contentChanges.push({name,tag})
 }
}
assert.deepEqual(results.contentChanges,[])
const reopened=createServerAutomationHost(saved.bytes);results.reopened=reopened.ok;assert.ok(reopened.ok)
fs.writeFileSync(`${out}/report.json`,JSON.stringify(results,null,2),'utf8');console.log(JSON.stringify(results,null,2));h.dispose();reopened.host.dispose()
