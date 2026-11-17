import {XMLSerializer} from '@xmldom/xmldom'
import {zipSync} from 'fflate'
import {W,unpack,xml,problem,compareProtected} from './package.mjs'
export function findMixedTocTitles(doc) {
  const ps=[]
  const walk=n=>{if(n.namespaceURI===W&&n.localName==='p'){ps.push(n);return}for(const c of Array.from(n.childNodes||[]))walk(c)}
  walk(doc.getElementsByTagNameNS(W,'body')[0])
  return ps.flatMap((p,index)=>{
    const text=Array.from(p.getElementsByTagNameNS(W,'t')).map(n=>n.textContent||'').join('').trim()
    const toc=Array.from(p.getElementsByTagNameNS(W,'instrText')).some(n=>/^\s*TOC\s/.test(n.textContent||''))
    return toc&&/^(目录|目\s+录|contents|table of contents)$/i.test(text)?[index]:[]
  })
}
export function separateTocTitle(bytes,op,signal) {
  const parts=unpack(bytes),doc=xml(parts['word/document.xml']),ps=[]
  const body=doc.getElementsByTagNameNS(W,'body')[0]
  const walk=n=>{if(n.namespaceURI===W&&n.localName==='p'){ps.push(n);return}for(const c of Array.from(n.childNodes||[]))walk(c)}
  walk(body)
  const undo=[]
  for(const index of op.indices){
    signal?.throwIfAborted()
    const p=ps[index],content=p?.parentNode,sdt=content?.parentNode
    if(content?.namespaceURI!==W||content.localName!=='sdtContent'||sdt?.namespaceURI!==W||sdt.localName!=='sdt'||sdt.parentNode!==body||Array.from(content.childNodes).filter(c=>c.nodeType===1)[0]!==p)throw problem('目录标题必须位于顶层目录内容控件首段',422)
    if(!Array.from(p.getElementsByTagNameNS(W,'instrText')).some(n=>/^\s*TOC\s/.test(n.textContent||'')))throw problem('目标没有原生目录域',422)
    const start=Array.from(p.childNodes).find(n=>n.nodeType===1&&Array.from(n.getElementsByTagNameNS(W,'fldChar')).some(f=>f.getAttributeNS(W,'fldCharType')==='begin'))
    if(!start)throw problem('目录域起点不明确',422)
    const prefix=[]
    for(const n of Array.from(p.childNodes)){
      if(n===start)break
      if(n.nodeType!==1||n.namespaceURI===W&&n.localName==='pPr')continue
      if(n.namespaceURI!==W||n.localName!=='r'||Array.from(n.childNodes).some(c=>c.nodeType===1&&(c.namespaceURI!==W||!['rPr','t'].includes(c.localName))))throw problem('目录标题含复杂结构，不能安全分离',422)
      prefix.push(n)
    }
    const titleText=prefix.flatMap(n=>Array.from(n.getElementsByTagNameNS(W,'t')).map(t=>t.textContent)).join('').trim()
    if(!/^(目录|目\s+录|contents|table of contents)$/i.test(titleText))throw problem('未确认目录标题文字',422)
    if(p.getElementsByTagNameNS(W,'sectPr').length)throw problem('目录标题含分节属性',422)
    const title=doc.createElementNS(W,'w:p'),pr=Array.from(p.childNodes).find(n=>n.namespaceURI===W&&n.localName==='pPr')
    if(pr)title.appendChild(pr.cloneNode(true))
    for(const n of prefix)title.appendChild(n)
    body.insertBefore(title,sdt)
    undo.push(()=>{for(const n of prefix)p.insertBefore(n,start);body.removeChild(title)})
  }
  const serialize=()=>new TextEncoder().encode(new XMLSerializer().serializeToString(doc))
  const output=serialize()
  for(const restore of undo.reverse())restore()
  compareProtected(bytes,zipSync({...parts,'word/document.xml':serialize()}),{formatting:false})
  signal?.throwIfAborted()
  return {bytes:zipSync({...parts,'word/document.xml':output}),affected:op.indices.length,validation:{status:'passed',checks:['目录标题与域分离','目录域及非目标内容保持'],limitations:['目录页码未重算；分离后须重新读取段落索引，分页仍需实际页面验收']}}
}

// Inspect actual field instructions; a TOC paragraph style alone is not a TOC field.
export function hasNativeToc(bytes) {
  const document = xml(unpack(bytes)['word/document.xml'])
  const instructions = [...document.getElementsByTagNameNS(W, 'instrText')].map(node => node.textContent || '').join('')
  return /(?:^|\s)TOC(?:\s|$)/i.test(instructions) ||
    [...document.getElementsByTagNameNS(W, 'fldSimple')].some(node => /^\s*TOC(?:\s|$)/i.test(node.getAttributeNS(W, 'instr') || ''))
}
