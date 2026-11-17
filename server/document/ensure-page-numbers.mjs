import {zipSync} from 'fflate'
import {XMLSerializer} from '@xmldom/xmldom'
import {unpack,xml,W,problem,compareProtected} from './package.mjs'
import {sectionNodes} from './structure.mjs'
import {formatFooterPageNumber} from './footer-page-number.mjs'
const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PR='http://schemas.openxmlformats.org/package/2006/relationships'
const CT='http://schemas.openxmlformats.org/package/2006/content-types'
const children=(n,name)=>Array.from(n?.childNodes||[]).filter(c=>c.namespaceURI===W&&c.localName===name)
const encode=d=>new TextEncoder().encode(new XMLSerializer().serializeToString(d))

// Explicit whole-document page-number request only. Preserve all existing footer
// text and shapes; add a field only to empty footers. Never invent header text.
export function ensurePageNumbers(bytes,signal){
  const parts=unpack(bytes),output={...parts},doc=xml(parts['word/document.xml'])
  const rels=xml(parts['word/_rels/document.xml.rels']),types=xml(parts['[Content_Types].xml'])
  const sections=sectionNodes(doc.getElementsByTagNameNS(W,'body')[0]),undo=[],added=[]
  const targets=new Set();let inherited
  const set=(n,k,v)=>{const had=n.hasAttributeNS(W,k),old=n.getAttributeNS(W,k);n.setAttributeNS(W,'w:'+k,v);undo.push(()=>had?n.setAttributeNS(W,'w:'+k,old):n.removeAttributeNS(W,k))}
  for(const section of sections){
    signal?.throwIfAborted()
    let numbering=children(section,'pgNumType')[0]
    if(!numbering){numbering=doc.createElementNS(W,'w:pgNumType');const next=Array.from(section.childNodes).find(n=>['cols','formProt','vAlign','noEndnote','titlePg','textDirection','docGrid','sectPrChange'].includes(n.localName));section.insertBefore(numbering,next||null);undo.push(()=>section.removeChild(numbering))}
    set(numbering,'fmt','decimal')
    for(const ref of children(section,'footerReference')){
      const rel=Array.from(rels.documentElement.childNodes).find(n=>n.nodeType===1&&n.getAttribute('Id')===ref.getAttributeNS(R,'id'))
      const name=rel?.getAttribute('Target')
      if(!name||!/^([^/]+\.xml|footer\d+\.xml)$/.test(name)||rel.getAttribute('TargetMode')==='External')throw problem('页脚引用需单独检查',422)
      const path='word/'+name
      if(!parts[path])throw problem('页脚部件缺失',422)
      targets.add(path)
      if(ref.getAttributeNS(W,'type')==='default')inherited=path
    }
    if(!inherited){
      let i=1;while(output[`word/piFooter${i}.xml`]||Array.from(rels.documentElement.childNodes).some(n=>n.nodeType===1&&n.getAttribute('Id')===`piFooter${i}`))i++
      inherited=`word/piFooter${i}.xml`;added.push(inherited)
      output[inherited]=new TextEncoder().encode(`<w:ftr xmlns:w="${W}"/>`)
      const rel=rels.createElementNS(PR,'Relationship');rel.setAttribute('Id',`piFooter${i}`);rel.setAttribute('Type',R+'/footer');rel.setAttribute('Target',`piFooter${i}.xml`);rels.documentElement.appendChild(rel);undo.push(()=>rels.documentElement.removeChild(rel))
      const type=types.createElementNS(CT,'Override');type.setAttribute('PartName','/'+inherited);type.setAttribute('ContentType','application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml');types.documentElement.appendChild(type);undo.push(()=>types.documentElement.removeChild(type))
      const ref=doc.createElementNS(W,'w:footerReference');ref.setAttributeNS(W,'w:type','default');ref.setAttributeNS(R,'r:id',`piFooter${i}`);section.insertBefore(ref,section.firstChild);undo.push(()=>section.removeChild(ref))
    }
    targets.add(inherited)
  }
  const footerDocs=[]
  for(const path of targets){
    const footer=xml(output[path]);footerDocs.push([path,footer])
    const fields=[...Array.from(footer.getElementsByTagNameNS(W,'instrText')).map(n=>n.textContent),...Array.from(footer.getElementsByTagNameNS(W,'fldSimple')).map(n=>n.getAttributeNS(W,'instr'))]
    if(fields.some(t=>/^\s*PAGE(?:\s|$)/i.test(t)))continue
    const text=Array.from(footer.getElementsByTagNameNS(W,'t')).map(n=>n.textContent).join('').trim()
    if(text||fields.length)throw problem('无PAGE的页脚含其他文字或域，不能自行决定页码布局',422)
    const p=footer.createElementNS(W,'w:p'),r=footer.createElementNS(W,'w:r'),field=footer.createElementNS(W,'w:fldSimple')
    field.setAttributeNS(W,'w:instr',' PAGE ');field.setAttributeNS(W,'w:dirty','true')
    const t=footer.createElementNS(W,'w:t');t.textContent='1';r.appendChild(t);field.appendChild(r);p.appendChild(field);footer.documentElement.appendChild(p);undo.push(()=>footer.documentElement.removeChild(p))
    output[path]=encode(footer)
  }
  output['word/document.xml']=encode(doc);output['word/_rels/document.xml.rels']=encode(rels);output['[Content_Types].xml']=encode(types)
  const result=zipSync(output)
  for(const restore of undo.reverse())restore()
  const restored={...output,'word/document.xml':encode(doc),'word/_rels/document.xml.rels':encode(rels),'[Content_Types].xml':encode(types)}
  for(const [path,d]of footerDocs)restored[path]=encode(d)
  for(const path of added)delete restored[path]
  compareProtected(bytes,zipSync(restored),{formatting:false})
  const formatted=formatFooterPageNumber(result,{scope:'footerPageNumber',sectionIndices:sections.map((_,i)=>i),alignment:'Centered'},signal)
  formatted.validation.checks.push('空白页脚补充PAGE且分节编号设为decimal；原有起点保留')
  formatted.validation.limitations=['页码域和居中属性已处理，缓存数字及页面位置需实际布局更新；不重算目录页码']
  return formatted
}
