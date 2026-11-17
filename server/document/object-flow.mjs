import {zipSync} from 'fflate'
import {XMLSerializer} from '@xmldom/xmldom'
import {W,A,unpack,xml,canonical,compareProtected,problem} from './package.mjs'
export const WP='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
const MC='http://schemas.openxmlformats.org/markup-compatibility/2006'
const direct=(n,ns,name)=>Array.from(n?.childNodes||[]).filter(c=>c.namespaceURI===ns&&c.localName===name)
const descendants=(n,ns,name)=>Array.from(n.getElementsByTagNameNS(ns,name))
const text=n=>descendants(n,W,'t').map(t=>t.textContent||'').join('')

export function flowTarget(anchor,mode) {
  if(!anchor)throw problem('浮动对象索引超出范围',422)
  let host=anchor
  while(host&&!(host.namespaceURI===W&&host.localName==='p'))host=host.parentNode
  if(host?.parentNode?.namespaceURI!==W||host.parentNode.localName!=='body')throw problem('只支持正文顶层对象，表内、修订及包装对象需单独处理',422)
  if(descendants(host,WP,'anchor').length!==1)throw problem('同段多个浮动对象无法确定布局关系',422)
  if(mode==='inlinePicture') {
    if(anchor.parentNode.localName!=='drawing'||anchor.parentNode.parentNode.localName!=='r'||anchor.parentNode.parentNode.parentNode!==host||descendants(anchor,W,'txbxContent').length||descendants(anchor,A,'graphicData')[0]?.getAttribute('uri')!=='http://schemas.openxmlformats.org/drawingml/2006/picture')throw problem('只支持无兼容分支的独立浮动图片',422)
    if(!direct(anchor,WP,'extent').length||!direct(anchor,WP,'docPr').length)throw problem('浮动图片结构不完整',422)
    return {host,source:anchor}
  }
  let source=anchor
  while(source.parentNode&&!(source.namespaceURI===MC&&source.localName==='AlternateContent'))source=source.parentNode
  if(source.namespaceURI!==MC||source.localName!=='AlternateContent'||source.parentNode.localName!=='r'||source.parentNode.parentNode!==host)throw problem('仅支持明确的题注兼容分支',422)
  const choices=direct(source,MC,'Choice'),fallbacks=direct(source,MC,'Fallback')
  if(choices.length!==1||fallbacks.length!==1)throw problem('题注兼容分支不唯一',422)
  const boxes=descendants(choices[0],W,'txbxContent'),fallback=descendants(fallbacks[0],W,'txbxContent')
  if(boxes.length!==1||fallback.length!==1||canonical(boxes[0])!==canonical(fallback[0]))throw problem('题注兼容分支内容不一致，不能合并',422)
  const paragraphs=direct(boxes[0],W,'p'),label=text(boxes[0]).trim()
  if(paragraphs.filter(p=>text(p).trim()).length!==1)throw problem('文本框含多段文字或累积题注，需要确认保留内容，不能全部展开',422)
  if(!/^[图表]\s*\d+(?:[.－-]\d+)*\s*\S/u.test(label)||!paragraphs.length)throw problem('不是可确认的图表题注',422)
  // Only plain text paragraphs: no fields, bookmarks, hyperlinks, revisions or embedded objects.
  const allowed=new Set(['txbxContent','p','pPr','r','rPr','t'])
  const scan=n=>{for(const child of Array.from(n.childNodes||[]).filter(c=>c.nodeType===1)){
    if(n.namespaceURI===W&&['pPr','rPr'].includes(n.localName))continue
    if(child.namespaceURI!==W||!allowed.has(child.localName))throw problem('复杂题注内容暂不支持移入文档流',422)
    scan(child)
  }}
  scan(boxes[0])
  if(descendants(boxes[0],W,'sectPr').length)throw problem('题注含分节设置',422)
  const copy=host.cloneNode(true)
  for(const branch of descendants(copy,MC,'AlternateContent'))branch.parentNode.removeChild(branch)
  const pictures=descendants(host,WP,'inline').length
  let next=host.nextSibling;while(next&&next.nodeType!==1)next=next.nextSibling
  if(label.startsWith('表'))while(next?.namespaceURI===W&&next.localName==='p'&&Array.from(next.childNodes).filter(n=>n.nodeType===1).every(n=>n.namespaceURI===W&&n.localName==='pPr'&&!Array.from(n.childNodes).some(c=>c.nodeType===1))){
    next=next.nextSibling;while(next&&next.nodeType!==1)next=next.nextSibling
  }
  // An adjacent table establishes placement even when its caption is anchored
  // in the preceding explanatory paragraph. Preserve that paragraph in place.
  const adjacentTable=label.startsWith('表')&&pictures===0&&next?.namespaceURI===W&&next.localName==='tbl'
  if(text(copy).trim()&&!adjacentTable)throw problem('宿主段落含其他正文，无法安全决定题注位置',422)
  let previous=host.previousSibling;while(previous&&previous.nodeType!==1)previous=previous.previousSibling
  const precedingPicture=pictures===0&&previous?.namespaceURI===W&&previous.localName==='p'&&!text(previous).trim()&&descendants(previous,WP,'inline').length===1&&descendants(previous,WP,'anchor').length===0
  if(!(label.startsWith('图')&&(pictures===1||precedingPicture))&&!(label.startsWith('表')&&pictures===0&&next?.namespaceURI===W&&next.localName==='tbl'))throw problem('题注与图片或表格没有明确的相邻关系',422)
  // Empty textbox padding is not body content; keep all nonempty or formatted paragraphs.
  while(paragraphs.length>1){
    const last=paragraphs.at(-1),elements=Array.from(last.childNodes).filter(n=>n.nodeType===1)
    if(elements.some(n=>n.namespaceURI!==W||n.localName!=='pPr'||Array.from(n.childNodes).some(c=>c.nodeType===1)))break
    paragraphs.pop()
  }
  return {host,source,paragraphs,before:adjacentTable?next:host.nextSibling}
}

export function flowObjects(bytes,op,signal) {
  const parts=unpack(bytes),document=xml(parts['word/document.xml'])
  const anchors=descendants(document,WP,'anchor')
  const targets=op.objectIndices.map(index=>flowTarget(anchors[index],op.objectLayout))
  const undo=[]
  for(const target of targets){
    signal?.throwIfAborted()
    const {source,host}=target
    if(op.objectLayout==='inlinePicture'){
      const inline=document.createElementNS(WP,'wp:inline')
      for(const key of ['distT','distB','distL','distR'])inline.setAttribute(key,'0')
      for(const name of ['extent','effectExtent','docPr','cNvGraphicFramePr','graphic']){
        const child=direct(source,name==='graphic'?A:WP,name)[0]
        if(child)inline.appendChild(child.cloneNode(true))
      }
      source.parentNode.replaceChild(inline,source)
      undo.push(()=>inline.parentNode.replaceChild(source,inline))
    }else{
      const placeholder=document.createComment('pi-flow-verification')
      source.parentNode.replaceChild(placeholder,source)
      const added=target.paragraphs.map(p=>p.cloneNode(true)),next=target.before
      for(const p of added)host.parentNode.insertBefore(p,next)
      undo.push(()=>{for(const p of added)p.parentNode.removeChild(p);placeholder.parentNode.replaceChild(source,placeholder)})
    }
  }
  const serialize=()=>new TextEncoder().encode(new XMLSerializer().serializeToString(document))
  // Keep placeholders only in the reversible verification DOM, never in the output.
  const outputDocument=xml(serialize())
  const clean=n=>{for(const c of Array.from(n.childNodes||[])){if(c.nodeType===8&&c.data==='pi-flow-verification')n.removeChild(c);else clean(c)}}
  clean(outputDocument)
  const output=new TextEncoder().encode(new XMLSerializer().serializeToString(outputDocument))
  for(const restore of undo.reverse())restore()
  compareProtected(bytes,zipSync({...parts,'word/document.xml':serialize()}),{formatting:false})
  const remaining=descendants(outputDocument,WP,'anchor').length
  if(remaining!==anchors.length-targets.length)throw problem('浮动对象转换数量验证失败',422)
  signal?.throwIfAborted()
  return {bytes:zipSync({...parts,'word/document.xml':output}),affected:targets.length,validation:{status:'passed',checks:['对象文档流转换数量验证','非目标包内容保持','兼容题注单份保留'],limitations:['对象已进入文档流，实际分页与位置仍需页面验收；重新读取段落和对象索引后才能继续操作']}}
}
