import fs from 'node:fs'
import assert from 'node:assert/strict'
import {unzipSync} from 'fflate'
import {DOMParser} from '@xmldom/xmldom'
const source=unzipSync(fs.readFileSync('data/documents/ea68ac07-cf3b-495a-9772-608a38afc15f/original.docx'))
const reports=[]
const examples=[]
let normalizeCompatibility=false
const parse=b=>new DOMParser().parseFromString(new TextDecoder().decode(b),'application/xml')
function canonical(n){
 if(n.nodeType===3||n.nodeType===4)return n.data.trim()?n.data:''
 if(n.nodeType!==1)return ''
 const attrs=Array.from(n.attributes).filter(a=>a.namespaceURI!=='http://www.w3.org/2000/xmlns/').map(a=>[a.namespaceURI||'',a.localName,normalizeCompatibility&&a.namespaceURI==='http://schemas.openxmlformats.org/markup-compatibility/2006'&&a.localName==='Ignorable'?a.value.trim().split(/\s+/).sort().join(' '):a.value]).sort((a,b)=>JSON.stringify(a).localeCompare(JSON.stringify(b)))
 return [n.namespaceURI,n.localName,attrs,Array.from(n.childNodes).map(canonical).filter(Boolean)]
}
function differences(a,b,path){
 if(JSON.stringify(a)===JSON.stringify(b))return
 if(!Array.isArray(a)||!Array.isArray(b)){if(examples.length<60)examples.push({path,before:a,after:b});return}
 if(a.length!==b.length){if(examples.length<60)examples.push({path,lengthBefore:a.length,lengthAfter:b.length});return}
 for(let i=0;i<a.length;i++)differences(a[i],b[i],path+'/'+i)
}
for(const file of ['output/editor-only-probe/roundtrip.docx','output/editor-browser-probe/roundtrip.docx']){
 const after=unzipSync(fs.readFileSync(file)),changedXml=[],changedBinary=[]
 for(const [name,bytes]of Object.entries(source)){
  if(!after[name])continue
  if(/\.(xml|rels)$/.test(name)){const a=canonical(parse(bytes).documentElement),b=canonical(parse(after[name]).documentElement);if(JSON.stringify(a)!==JSON.stringify(b)){changedXml.push(name);if(reports.length===0)differences(a,b,name)}}
  else if(!Buffer.from(bytes).equals(Buffer.from(after[name])))changedBinary.push(name)
 }
 reports.push({file,missing:Object.keys(source).filter(n=>!after[n]),added:Object.keys(after).filter(n=>!source[n]),changedXml,changedBinary})
 normalizeCompatibility=true
 const substantive=changedXml.filter(n=>JSON.stringify(canonical(parse(source[n]).documentElement))!==JSON.stringify(canonical(parse(after[n]).documentElement)))
 reports.at(-1).differencesAfterIgnorableSorting=substantive
 assert.deepEqual(substantive,[])
 normalizeCompatibility=false
}
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
function spacing(file){const d=parse(unzipSync(fs.readFileSync(file))['word/document.xml']),ps=[];const walk=n=>{if(n.namespaceURI===W&&n.localName==='p'){ps.push(n);return}for(const c of Array.from(n.childNodes||[]))walk(c)};walk(d.getElementsByTagNameNS(W,'body')[0]);const n=ps[16].getElementsByTagNameNS(W,'spacing')[0];return {line:n?.getAttributeNS(W,'line'),rule:n?.getAttributeNS(W,'lineRule')}}
const browser={formatted:spacing('output/editor-browser-probe/formatted.docx'),undone:spacing('output/editor-browser-probe/undone.docx')}
assert.deepEqual(browser.formatted,{line:'480',rule:'auto'});assert.deepEqual(browser.undone,{line:'360',rule:'auto'})
fs.writeFileSync('output/editor-only-probe/preservation.json',JSON.stringify({reports,browser,examples,limitations:['Canonical comparison ignores whitespace-only text and namespace prefix spellings; XML differences require classification, not automatic loss claims']},null,2),'utf8')
console.log(JSON.stringify({reports,browser},null,2))
