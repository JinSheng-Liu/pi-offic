import {isDeepStrictEqual} from 'node:util'
import {zipSync} from 'fflate'
import {XMLSerializer} from '@xmldom/xmldom'
import {unpack,xml,W,digest,canonical,compareProtected,problem} from './package.mjs'
import {readEditorDocument,findTextRanges} from './editor-reader.mjs'
import {fontProperties} from '../../shared/editor-format-ops.mjs'
const children=n=>Array.from(n?.childNodes||[]).filter(n=>n.nodeType===1)
const child=(n,key)=>children(n).find(n=>n.namespaceURI===W&&n.localName===key)
function add(parent,key,attrs={}){const n=parent.ownerDocument.createElementNS(W,'w:'+key);for(const [k,v]of Object.entries(attrs))n.setAttributeNS(W,'w:'+k,String(v));parent.appendChild(n);return n}
const chars=p=>p.runs.flatMap(r=>Array.from({length:r.text.length},()=>({font:r.font,fonts:r.effectiveFonts})))
export async function applyCharacterStyle(bytes,op,signal,formatText,paragraphNodes){
 const formatted=await formatText(bytes,{...op,scope:'text'},signal)
 const expected=readEditorDocument(formatted.bytes),index=op.indices[0],range=findTextRanges(expected.paragraphs[index],op.text)[op.occurrence??0]
 const parts=unpack(formatted.bytes),doc=xml(parts['word/document.xml']),styles=xml(parts['word/styles.xml']),p=paragraphNodes(doc)[index]
 const properties=fontProperties(op.font),name=op.styleName.trim(),ids=new Set()
 let offset=0
 for(const run of children(p).filter(n=>n.localName==='r'&&n.namespaceURI===W)){
  const content=children(run).filter(n=>n.localName!=='rPr'),length=content.reduce((sum,n)=>sum+(n.localName==='t'?(n.textContent||'').length:1),0),start=offset;offset+=length
  if(start>=range.end||offset<=range.start)continue
  let target=run
  if(start<range.start||offset>range.end){
   if(content.some(n=>n.localName!=='t'))throw problem('字符样式边界跨越控制字符，未提交',422)
   const text=content.map(n=>n.textContent).join(''),a=Math.max(0,range.start-start),b=Math.min(length,range.end-start)
   for(const [value,selected]of [[text.slice(0,a),false],[text.slice(a,b),true],[text.slice(b),false]])if(value){
    const split=run.cloneNode(false),pr=child(run,'rPr');if(pr)split.appendChild(pr.cloneNode(true))
    const t=add(split,'t');t.setAttributeNS('http://www.w3.org/XML/1998/namespace','xml:space','preserve');t.textContent=value
    p.insertBefore(split,run);if(selected)target=split
   }
   p.removeChild(run)
  }
  const pr=child(target,'rPr')||target.insertBefore(doc.createElementNS(W,'w:rPr'),target.firstChild)
  const base=child(pr,'rStyle')?.getAttributeNS(W,'val')||''
  const existing=children(styles.documentElement).filter(n=>n.localName==='style'),old=existing.find(n=>n.getAttributeNS(W,'styleId')===base)
  const signature=digest(new TextEncoder().encode(JSON.stringify({name,properties}))).slice(0,24)
  const baseId=base.startsWith('PiChar')&&child(old,'aliases')?.getAttributeNS(W,'val')===signature?child(old,'basedOn')?.getAttributeNS(W,'val')||'':base
  const id='PiChar'+digest(new TextEncoder().encode(JSON.stringify({signature,baseId}))).slice(0,24)
  if(!existing.some(n=>n.getAttributeNS(W,'styleId')===id)){
   const style=add(styles.documentElement,'style',{type:'character',customStyle:'1',styleId:id})
   add(style,'name',{val:existing.some(n=>child(n,'name')?.getAttributeNS(W,'val')===name)?`${name} · ${id.slice(-6)}`:name});add(style,'aliases',{val:signature})
   if(baseId)add(style,'basedOn',{val:baseId})
   add(style,'uiPriority',{val:10});add(style,'qFormat');const rpr=add(style,'rPr')
   for(const prop of properties)add(rpr,prop.localName,prop.attributes)
  }
  let link=child(pr,'rStyle');if(!link){link=doc.createElementNS(W,'w:rStyle');pr.insertBefore(link,pr.firstChild)}link.setAttributeNS(W,'w:val',id)
  for(const prop of properties){const direct=child(pr,prop.localName);if(!direct)continue;for(const key of Object.keys(prop.attributes))direct.removeAttributeNS(W,key);if(!direct.attributes.length)pr.removeChild(direct)}
  ids.add(id)
 }
 signal?.throwIfAborted()
 const encode=d=>new TextEncoder().encode(new XMLSerializer().serializeToString(d))
 parts['word/document.xml']=encode(doc);parts['word/styles.xml']=encode(styles)
 const output=Buffer.from(zipSync(parts)),actual=readEditorDocument(output)
 if(actual.text!==expected.text||!isDeepStrictEqual(chars(actual.paragraphs[index]),chars(expected.paragraphs[index])))throw problem('字符样式继承或文字范围校验失败',422)
 const before=xml(unpack(formatted.bytes)['word/document.xml']),source=paragraphNodes(before)[index]
 if((child(source,'pPr')?canonical(child(source,'pPr')):'')!==(child(p,'pPr')?canonical(child(p,'pPr')):''))throw problem('字符样式改变段落样式',422)
 p.parentNode.replaceChild(doc.importNode(source,true),p)
 if(canonical(doc.documentElement)!==canonical(before.documentElement))throw problem('字符样式改变非目标段落',422)
 // Compare untouched characters' direct properties as well as effective values.
 const direct=p=>p.runs.flatMap(r=>Array.from({length:r.text.length},()=>r.value))
 const oldDirect=direct(expected.paragraphs[index]),newDirect=direct(actual.paragraphs[index])
 for(let i=0;i<oldDirect.length;i++)if((i<range.start||i>=range.end)&&!isDeepStrictEqual(oldDirect[i],newDirect[i]))throw problem('字符样式影响未选中文字',422)
 const comparison=unpack(output),checkDoc=xml(comparison['word/document.xml'])
 const checkP=paragraphNodes(checkDoc)[index],originalP=paragraphNodes(xml(unpack(bytes)['word/document.xml']))[index]
 checkP.parentNode.replaceChild(checkDoc.importNode(originalP,true),checkP)
 comparison['word/document.xml']=encode(checkDoc)
 const validation=compareProtected(bytes,zipSync(comparison))
 validation.checks.push('字符样式关联与有效格式一致','正文段落样式保持','非目标段落保持')
 return {bytes:output,affected:1,validation}
}
