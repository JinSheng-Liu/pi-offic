import {zipSync} from 'fflate'
import {XMLSerializer} from '@xmldom/xmldom'
import {posix} from 'node:path'
import {unpack,xml,W,canonical,digest,problem} from './package.mjs'
const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const elements=(n,name)=>Array.from(n.getElementsByTagNameNS(W,name))
const normalized=blocks=>blocks.flatMap(b=>elements(b,'t')).map(t=>t.textContent||'').join('').replace(/\s/g,'')
export function referenceRegion(doc,region='cover'){
  const body=elements(doc,'body')[0],blocks=Array.from(body.childNodes).filter(n=>n.nodeType===1),end=blocks.findIndex(n=>n.localName==='p'&&elements(n,'sectPr').length)
  if(['authorization','declaration'].includes(region)){
    const titles=blocks.map((n,i)=>({n,i})).filter(({n})=>normalized([n])===(region==='authorization'?'学位论文版权使用授权书':'学位论文原创性声明'))
    if(titles.length!==1)throw problem('前置页标题缺失或重复，无法确定范围',422)
    const title=titles[0].i
    let start=title
    while(start>0&&!elements(blocks[start-1],'sectPr').length)start--
    const finish=blocks.findIndex((n,i)=>i>=title&&elements(n,'sectPr').length)
    if(finish<0||finish-start>30||!blocks.slice(title,finish+1).some(n=>normalized([n]).includes('学位论文作者签名')))throw problem('前置页签署区或分节边界不明确',422)
    return {body,blocks:blocks.slice(start,finish+1)}
  }
  if(end<0||end>30)throw problem('参考区域需要明确的首节分节边界',422)
  return {body,blocks:blocks.slice(0,end+1)}
}
// Restore an explicitly supplied reference, never copy names or dates into a
// different document. Dependencies must already exist identically in the target.
export function restoreReferenceCover(bytes,reference,signal,region='cover'){
  if(!['cover','authorization','declaration'].includes(region))throw problem('不支持的参考区域',422)
  signal?.throwIfAborted()
  const parts=unpack(bytes),refParts=unpack(reference),doc=xml(parts['word/document.xml']),refDoc=xml(refParts['word/document.xml']),target=referenceRegion(doc,region),source=referenceRegion(refDoc,region)
  if(normalized(target.blocks)!==normalized(source.blocks))throw problem('参考区域与当前区域文字不同，不能直接替换；需要字段映射',422)
  const symbols=blocks=>blocks.flatMap(b=>elements(b,'sym').map(canonical))
  if(JSON.stringify(symbols(target.blocks))!==JSON.stringify(symbols(source.blocks)))throw problem('参考区域符号内容不一致',422)
  const copies=source.blocks.map(n=>n.cloneNode(true))
  const styles=parts['word/styles.xml']?xml(parts['word/styles.xml']):null,refStyles=refParts['word/styles.xml']?xml(refParts['word/styles.xml']):null
  let stylesChanged=false
  const mapped=new Map(),visiting=new Set()
  const importStyle=id=>{
    if(mapped.has(id))return mapped.get(id)
    const a=styles&&elements(styles,'style').find(s=>s.getAttributeNS(W,'styleId')===id),b=refStyles&&elements(refStyles,'style').find(s=>s.getAttributeNS(W,'styleId')===id)
    if(!styles||!b)throw problem('参考区域样式依赖缺失',422)
    if(visiting.has(id))throw problem('参考样式依赖循环',422)
    visiting.add(id)
    const clone=b.cloneNode(true)
    for(const tag of ['basedOn','next','link'])for(const dep of elements(clone,tag)){
      const depId=dep.getAttributeNS(W,'val')
      if(depId===id)continue
      dep.setAttributeNS(W,'w:val',importStyle(depId))
    }
    visiting.delete(id)
    if(a&&canonical(a)===canonical(clone)){mapped.set(id,id);return id}
    const newId='PiRef'+digest(new TextEncoder().encode(canonical(clone))).slice(0,16)
    clone.setAttributeNS(W,'w:styleId',newId)
    for(const dep of elements(clone,'next'))if(dep.getAttributeNS(W,'val')===id)dep.setAttributeNS(W,'w:val',newId)
    const name=elements(clone,'name')[0];if(name)name.setAttributeNS(W,'w:val','参考区域 '+newId)
    if(!elements(styles,'style').some(n=>n.getAttributeNS(W,'styleId')===newId)){styles.documentElement.appendChild(clone);stylesChanged=true}
    mapped.set(id,newId);return newId
  }
  for(const block of copies)for(const tag of ['pStyle','rStyle','tblStyle'])for(const n of elements(block,tag))n.setAttributeNS(W,'w:val',importStyle(n.getAttributeNS(W,'val')))
  const relations=parts['word/_rels/document.xml.rels']?Array.from(xml(parts['word/_rels/document.xml.rels']).documentElement.childNodes).filter(n=>n.nodeType===1):[]
  const refRelations=refParts['word/_rels/document.xml.rels']?Array.from(xml(refParts['word/_rels/document.xml.rels']).documentElement.childNodes).filter(n=>n.nodeType===1):[]
  // Keep the target section, including its header/footer relationships.
  const oldSection=elements(target.blocks.at(-1),'sectPr')[0],newSection=elements(copies.at(-1),'sectPr')[0]
  const keptSection=oldSection.cloneNode(true)
  if(['authorization','declaration'].includes(region)){
    for(const key of ['pgSz','pgMar','docGrid','cols']){
      const old=elements(keptSection,key)[0],replacement=elements(newSection,key)[0]
      if(old){if(replacement)old.parentNode.replaceChild(replacement.cloneNode(true),old);else old.parentNode.removeChild(old)}
      else if(replacement)keptSection.appendChild(replacement.cloneNode(true))
    }
  }
  newSection.parentNode.replaceChild(keptSection,newSection)
  for(const block of copies){
    const visit=n=>{
      if(n.nodeType===1)for(const attr of Array.from(n.attributes||[]).filter(a=>a.namespaceURI===R)){
        // Section refs were retained from the target and need no remapping.
        if(['headerReference','footerReference'].includes(n.localName))continue
        const rel=refRelations.find(r=>r.getAttribute('Id')===attr.value)
        if(!rel||rel.getAttribute('TargetMode')==='External'||!rel.getAttribute('Type').endsWith('/image'))throw problem('参考区域含不支持的关系依赖',422)
        const data=refParts[posix.normalize('word/'+rel.getAttribute('Target'))]
        const match=data&&relations.find(r=>r.getAttribute('Type')===rel.getAttribute('Type')&&parts[posix.normalize('word/'+r.getAttribute('Target'))]&&digest(parts[posix.normalize('word/'+r.getAttribute('Target'))])===digest(data))
        if(!match)throw problem('目标文档没有与参考区域相同的图片资源',422)
        attr.value=match.getAttribute('Id')
      }
      for(const child of Array.from(n.childNodes||[]))visit(child)
    };visit(block)
  }
  const bookmarks=blocks=>blocks.flatMap(b=>['bookmarkStart','bookmarkEnd'].flatMap(tag=>elements(b,tag).map(n=>`${tag}:${n.getAttributeNS(W,'id')}:${n.getAttributeNS(W,'name')}`))).sort()
  if(JSON.stringify(bookmarks(target.blocks))!==JSON.stringify(bookmarks(copies)))throw problem('封面书签集合不同，不能安全替换',422)
  const tail=target.blocks.at(-1).nextSibling
  for(const block of target.blocks)target.body.removeChild(block)
  for(const block of copies)target.body.insertBefore(block,tail)
  const output=new TextEncoder().encode(new XMLSerializer().serializeToString(doc))
  for(const block of copies)target.body.removeChild(block)
  for(const block of target.blocks)target.body.insertBefore(block,tail)
  if(canonical(doc.documentElement)!==canonical(xml(parts['word/document.xml']).documentElement))throw problem('参考区域恢复影响了非目标正文',422)
  if(normalized(referenceRegion(xml(output),region).blocks)!==normalized(target.blocks))throw problem('参考区域回读文字不一致',422)
  signal?.throwIfAborted()
  return {bytes:zipSync({...parts,'word/document.xml':output,...(stylesChanged?{'word/styles.xml':new TextEncoder().encode(new XMLSerializer().serializeToString(styles))}:{})}),affected:copies.length,validation:{status:'partial',checks:['参考区域文字匹配','原有分节及书签集合保持','图片资源按内容匹配','非目标正文与其他包部件保持'],limitations:['仅支持文字相同且依赖兼容的参考区域；需要实际渲染确认，不能据此认定全文排版通过']}}
}
