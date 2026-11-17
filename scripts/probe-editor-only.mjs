// Independent architecture probe: no application engine or core-ts imports.
import fs from 'node:fs'
import path from 'node:path'
import assert from 'node:assert/strict'
import {createHash} from 'node:crypto'
import {unzipSync} from 'fflate'
import {DOMParser} from '@xmldom/xmldom'
import {createServerAutomationHost} from '@docx-editor.dev/core/automation'

const source=process.argv[2]||'data/documents/ea68ac07-cf3b-495a-9772-608a38afc15f/original.docx'
const out=process.argv[3]||'output/editor-only-probe'
fs.mkdirSync(out,{recursive:true})
const bytes=fs.readFileSync(source),hash=b=>createHash('sha256').update(b).digest('hex')
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const parse=b=>new DOMParser().parseFromString(new TextDecoder().decode(b),'application/xml')
const nodes=(d,n)=>Array.from(d.getElementsByTagNameNS(W,n))
function inventory(b){
  const parts=unzipSync(b),stories={}
  for(const [name,value] of Object.entries(parts))if(/^word\/.*\.xml$/.test(name)){
    const d=parse(value)
    stories[name]={text:nodes(d,'t').map(n=>n.textContent),fields:nodes(d,'instrText').map(n=>n.textContent),counts:Object.fromEntries(['tbl','drawing','pict','txbxContent','sectPr','fldChar','oMath'].map(n=>[n,nodes(d,n).length]))}
  }
  return {parts:Object.keys(parts).sort(),stories,media:Object.fromEntries(Object.entries(parts).filter(([n])=>n.startsWith('word/media/')).map(([n,b])=>[n,hash(b)])),relationships:Object.fromEntries(Object.entries(parts).filter(([n])=>n.endsWith('.rels')).map(([n,b])=>[n,Array.from(parse(b).documentElement.childNodes).filter(x=>x.nodeType===1).map(x=>[x.getAttribute('Id'),x.getAttribute('Type'),x.getAttribute('Target'),x.getAttribute('TargetMode')]).sort()]))}
}
function diff(a,b){return {samePartNames:JSON.stringify(a.parts)===JSON.stringify(b.parts),sameMedia:JSON.stringify(a.media)===JSON.stringify(b.media),sameRelationships:JSON.stringify(a.relationships)===JSON.stringify(b.relationships),changedStories:Object.keys(a.stories).filter(n=>JSON.stringify(a.stories[n])!==JSON.stringify(b.stories[n]))}}
const opened=createServerAutomationHost(bytes)
if(!opened.ok)throw Error(JSON.stringify(opened))
const host=opened.host
function q(op){const r=host.execute({operations:[op],expectedRevision:host.revision()});if(!r.ok)throw Error(JSON.stringify(r));return r.results[0].value}
const document=q({op:'getDocument'}).handle,body=q({op:'getBody',document}).handle
const ps=q({op:'getParagraphs',body}).handles
const text=q({op:'getText',target:body}).text
const originalInventory=inventory(bytes),saved=host.save()
if(!saved.ok)throw Error(JSON.stringify(saved))
fs.writeFileSync(path.join(out,'roundtrip.docx'),saved.bytes)
const target=ps.find(p=>q({op:'getText',target:p}).text.trim().length>100)
if(!target)throw Error('No prose paragraph')
const beforeText=q({op:'getText',target}).text
const fontWrite=q({op:'setFont',span:{paragraph:target},font:{name:'宋体',size:12}})
const paragraphWrite=q({op:'setParagraphFormat',paragraph:{paragraph:target},format:{alignment:'Justified',firstLineIndent:24,spaceBefore:0,spaceAfter:0}})
const edited=host.save();if(!edited.ok)throw Error(JSON.stringify(edited))
fs.writeFileSync(path.join(out,'edited.docx'),edited.bytes)
const reopen=createServerAutomationHost(edited.bytes)
const rq=op=>{const r=reopen.host.execute({operations:[op]});if(!r.ok)throw Error(JSON.stringify(r));return r.results[0].value}
const reopenedBody=reopen.ok?rq({op:'getBody',document:rq({op:'getDocument'}).handle}).handle:null
const reopenedTarget=reopen.ok?rq({op:'getParagraphs',body:reopenedBody}).handles[ps.indexOf(target)]:null
const savedDoc=parse(unzipSync(edited.bytes)['word/document.xml'])
const savedTarget=nodes(savedDoc,'p').find(p=>nodes(p,'t').map(n=>n.textContent).join('')===beforeText)
const fontSlots=savedTarget?nodes(savedTarget,'rFonts').map(n=>Object.fromEntries(['ascii','hAnsi','eastAsia','cs','asciiTheme','hAnsiTheme','eastAsiaTheme'].map(k=>[k,n.getAttributeNS(W,k)]))):[]
const report={source:path.resolve(source),sourceHash:hash(bytes),sourceUnchanged:hash(fs.readFileSync(source))===hash(bytes),paragraphs:ps.length,bodyCharacters:text.length,roundtrip:diff(originalInventory,inventory(saved.bytes)),edited:diff(originalInventory,inventory(edited.bytes)),targetIndex:ps.indexOf(target),targetTextPreserved:q({op:'getText',target}).text===beforeText,fontWrite,paragraphWrite,fontRead:q({op:'getFont',span:{paragraph:target}}),paragraphRead:q({op:'getParagraphFormat',paragraph:{paragraph:target}}),reopened:reopen.ok,limitations:['Headless only; no browser paint or page-layout verification','Inventory checks do not prove complete semantic equivalence','Mixed Chinese/Latin font slots not yet verified']}
report.savedFontSlots=fontSlots
report.reopenedFont=reopen.ok?rq({op:'getFont',span:{paragraph:reopenedTarget}}):null
report.reopenedParagraph=reopen.ok?rq({op:'getParagraphFormat',paragraph:{paragraph:reopenedTarget}}):null
assert.equal(report.sourceUnchanged,true)
for(const result of [report.roundtrip,report.edited]){assert.equal(result.sameMedia,true);assert.equal(result.sameRelationships,true);assert.equal(result.samePartNames,true);assert.deepEqual(result.changedStories,[])}
assert.equal(report.reopenedFont.font.size,12)
assert.equal(report.reopenedParagraph.format.firstLineIndent,24)
assert.equal(report.reopenedParagraph.format.alignment,'Justified')
assert.equal(report.targetTextPreserved,true)
report.explicitEastAsiaFontWritten=fontSlots.length>0&&fontSlots.every(f=>f.eastAsia==='宋体')
fs.writeFileSync(path.join(out,'report.json'),JSON.stringify(report,null,2),'utf8')
console.log(JSON.stringify(report,null,2));host.dispose();if(reopen.ok)reopen.host.dispose()
