import {posix} from 'node:path'
import {zipSync} from 'fflate'
import {XMLSerializer} from '@xmldom/xmldom'
import {W,unpack,xml,problem,compareProtected} from './package.mjs'
import {sectionNodes} from './structure.mjs'
const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const children=(n,name)=>Array.from(n.childNodes||[]).filter(c=>c.namespaceURI===W&&c.localName===name)
// Deliberately limited to a footer paragraph containing only a PAGE field.
// Shared footer parts are refused unless every referencing section is selected.
export function formatFooterPageNumber(bytes,op,signal) {
  const parts=unpack(bytes),document=xml(parts['word/document.xml'])
  const sections=sectionNodes(document.getElementsByTagNameNS(W,'body')[0])
  if(op.sectionIndices.some(i=>!sections[i]))throw problem('页脚分节索引超出范围',422)
  const rels=xml(parts['word/_rels/document.xml.rels'])
  const relationships=new Map(Array.from(rels.documentElement.childNodes).filter(n=>n.nodeType===1).map(n=>[n.getAttribute('Id'),n]))
  const inherited=new Map(),uses=new Map()
  sections.forEach((section,index)=>{
    for(const ref of children(section,'footerReference')){
      const rel=relationships.get(ref.getAttributeNS(R,'id'))
      if(!rel||rel.getAttribute('TargetMode')==='External'||!rel.getAttribute('Type').endsWith('/footer'))throw problem('页脚关系无效',422)
      const target=rel.getAttribute('Target'),path=target.startsWith('/')?target.slice(1):posix.normalize(posix.join('word',target))
      if(!/^word\/[^/]+\.xml$/.test(path)||!parts[path])throw problem('页脚路径无效',422)
      inherited.set(ref.getAttributeNS(W,'type')||'default',path)
    }
    for(const path of new Set(inherited.values())){if(!uses.has(path))uses.set(path,new Set());uses.get(path).add(index)}
  })
  const selected=new Set(op.sectionIndices),targets=[...uses].filter(([,indices])=>[...indices].some(i=>selected.has(i)))
  const output={...parts},undoParts={},checks=[]
  let affected=0
  for(const [path,indices]of targets){
    signal?.throwIfAborted()
    const doc=xml(parts[path]),undo=[]
    const property=(owner,group,key,attrs)=>{
      let pr=children(owner,group)[0];const created=!pr
      if(!pr){pr=doc.createElementNS(W,'w:'+group);owner.insertBefore(pr,owner.firstChild)}
      const old=children(pr,key)[0],node=old?old.cloneNode(true):doc.createElementNS(W,'w:'+key)
      for(const [k,v]of Object.entries(attrs))node.setAttributeNS(W,'w:'+k,String(v))
      if(key==='rFonts')for(const k of ['asciiTheme','hAnsiTheme','eastAsiaTheme','cstheme','csTheme'])node.removeAttributeNS(W,k)
      if(old)pr.replaceChild(node,old);else{
        const order=group==='rPr'?['rStyle','rFonts','b','bCs','i','iCs','caps','smallCaps','strike','dstrike','outline','shadow','emboss','imprint','noProof','snapToGrid','vanish','webHidden','color','spacing','w','kern','position','sz','szCs','highlight','u','effect','bdr','shd','fitText','vertAlign','rtl','cs','em','lang','eastAsianLayout','specVanish','oMath','rPrChange']:['pStyle','keepNext','keepLines','pageBreakBefore','framePr','widowControl','numPr','suppressLineNumbers','pBdr','shd','tabs','suppressAutoHyphens','kinsoku','wordWrap','overflowPunct','topLinePunct','autoSpaceDE','autoSpaceDN','bidi','adjustRightInd','snapToGrid','spacing','ind','contextualSpacing','mirrorIndents','suppressOverlap','jc','textDirection','textAlignment','textboxTightWrap','outlineLvl','divId','cnfStyle','rPr','sectPr','pPrChange']
        const next=Array.from(pr.childNodes).find(c=>c.namespaceURI===W&&order.indexOf(c.localName)>order.indexOf(key));pr.insertBefore(node,next||null)
      }
      undo.push(()=>{if(old)pr.replaceChild(old,node);else pr.removeChild(node);if(created&&!pr.childNodes.length)owner.removeChild(pr)})
    }
    for(const p of Array.from(doc.getElementsByTagNameNS(W,'p'))){
      const instructions=Array.from(p.getElementsByTagNameNS(W,'instrText'))
      const simple=Array.from(p.getElementsByTagNameNS(W,'fldSimple'))
      if(!instructions.some(n=>/^\s*PAGE(?:\s|$)/i.test(n.textContent))&&!simple.some(n=>/^\s*PAGE(?:\s|$)/i.test(n.getAttributeNS(W,'instr'))))continue
      if([...indices].some(i=>!selected.has(i)))throw problem('页脚被未选分节共享，请同时选中或先解除链接',422)
      for(let parent=p.parentNode;parent!==doc.documentElement;parent=parent?.parentNode){
        if(!parent||parent.namespaceURI!==W||!['sdt','sdtContent'].includes(parent.localName))throw problem('页码位于表格或复杂页脚结构中，暂不支持',422)
      }
      const elements=Array.from(p.childNodes).filter(n=>n.nodeType===1)
      if(elements.some(n=>n.namespaceURI!==W||!['pPr','r','fldSimple'].includes(n.localName)))throw problem('页码段落含其他结构',422)
      const texts=Array.from(p.getElementsByTagNameNS(W,'t')).map(n=>n.textContent).join('')
      if(!/^\s*[0-9IVXLCDMivxlcdm]*\s*$/.test(texts)||instructions.some(n=>!/^\s*PAGE\s*(?:\\\*\s*MERGEFORMAT)?\s*$/i.test(n.textContent))||simple.length>1)throw problem('仅支持独立 PAGE 页码段落，不能覆盖页脚其他文字或域',422)
      const runs=Array.from(p.getElementsByTagNameNS(W,'r'))
      if(runs.some(r=>Array.from(r.childNodes).some(n=>n.nodeType===1&&(n.namespaceURI!==W||!['rPr','t','instrText','fldChar'].includes(n.localName)))))throw problem('页码包含复杂运行内容',422)
      const marks=Array.from(p.getElementsByTagNameNS(W,'fldChar')).map(n=>n.getAttributeNS(W,'fldCharType'))
      if(simple.length?instructions.length||marks.length:instructions.length!==1||marks.join(',')!=='begin,separate,end')throw problem('PAGE 域结构不完整或存在多个域',422)
      if(op.alignment)property(p,'pPr','jc',{val:op.alignment==='Centered'?'center':op.alignment.toLowerCase()})
      for(const r of runs){
        if(op.font?.name)property(r,'rPr','rFonts',{ascii:op.font.name,hAnsi:op.font.name,eastAsia:op.font.name,cs:op.font.name})
        if(op.font?.size!==undefined){property(r,'rPr','sz',{val:Math.round(op.font.size*2)});property(r,'rPr','szCs',{val:Math.round(op.font.size*2)})}
      }
      affected++
    }
    output[path]=new TextEncoder().encode(new XMLSerializer().serializeToString(doc))
    for(const restore of undo.reverse())restore()
    undoParts[path]=new TextEncoder().encode(new XMLSerializer().serializeToString(doc))
  }
  if(!affected)throw problem('选中分节没有可格式化的 PAGE 页码，不自动新增',422)
  compareProtected(bytes,zipSync({...output,...undoParts}),{formatting:false})
  checks.push('页码域与非目标内容保持','指定页脚页码格式写入')
  signal?.throwIfAborted()
  return {bytes:zipSync(output),affected,validation:{status:'passed',checks,limitations:['仅修改已有页码的字体和对齐；不新增编号、不重算目录页码，页面位置仍需渲染检查']}}
}
