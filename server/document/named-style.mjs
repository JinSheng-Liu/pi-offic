import {isDeepStrictEqual} from 'node:util'
import {zipSync} from 'fflate'
import {XMLSerializer} from '@xmldom/xmldom'
import {unpack,xml,W,digest,canonical,compareProtected,problem} from './package.mjs'
import {readEditorDocument} from './editor-reader.mjs'
import {fontProperties,paragraphProperties} from '../../shared/editor-format-ops.mjs'
const children=n=>Array.from(n?.childNodes||[]).filter(n=>n.nodeType===1)
const child=(n,name)=>children(n).find(n=>n.namespaceURI===W&&n.localName===name)
function add(parent,name,attributes={}) {
 const node=parent.ownerDocument.createElementNS(W,`w:${name}`)
 for(const [key,value] of Object.entries(attributes))node.setAttributeNS(W,`w:${key}`,String(value))
 parent.appendChild(node);return node
}
function removeProperties(container,properties) {
 for(const prop of properties) {
  const node=child(container,prop.localName)
  if(!node)continue
  for(const key of Object.keys(prop.attributes))node.removeAttributeNS(W,key)
  if(!node.attributes.length&&!children(node).length)container.removeChild(node)
 }
}
// Immutable derived styles avoid changing any paragraph outside the selected scope.
export async function applyNamedStyle(bytes,op,signal,formatParagraphs,paragraphNodes) {
 const original=readEditorDocument(bytes)
 const indices=[...new Set(op.indices)]
 if(indices.some(i=>!original.paragraphs[i]||original.paragraphs[i].generatedToc))throw problem('命名样式不能用于无效段落或自动目录',422)
 const formatted=await formatParagraphs(bytes,[{scope:'indices',indices,font:op.font,paragraph:op.paragraph}],signal)
 const expected=readEditorDocument(formatted.bytes),parts=unpack(formatted.bytes)
 if(!parts['word/styles.xml'])throw problem('缺少样式定义',422)
 const styles=xml(parts['word/styles.xml']),document=xml(parts['word/document.xml']),nodes=paragraphNodes(document)
 if(nodes.length!==expected.paragraphs.length)throw problem('样式段落定位不一致',422)
 const fonts=fontProperties(op.font),paragraphs=paragraphProperties(op.paragraph)
 if(op.paragraph?.outlineLevel!==undefined)paragraphs.push({localName:'outlineLvl',attributes:{val:String(op.paragraph.outlineLevel-1)}})
 const applied=new Map()
 for(const index of indices) {
  signal?.throwIfAborted()
  const p=nodes[index],pr=child(p,'pPr')||add(p,'pPr')
  // Use the original base, not navigation-only styles introduced by formatting.
  let base=original.paragraphs[index].styleId
  const existing=children(styles.documentElement).filter(n=>n.localName==='style')
  const old=existing.find(n=>n.getAttributeNS(W,'styleId')===base)
  const props={name:op.styleName.trim(),base,font:fonts,paragraph:paragraphs}
  // Repeated application of an identical definition reuses the existing style.
  const signature=digest(new TextEncoder().encode(JSON.stringify({font:fonts,paragraph:paragraphs}))).slice(0,16)
  if(base.startsWith('PiStyle')&&child(old,'aliases')?.getAttributeNS(W,'val')===`PiFormat_${signature}`&&[props.name,`${props.name} · ${base.slice(-6)}`].includes(child(old,'name')?.getAttributeNS(W,'val')))base=child(old,'basedOn')?.getAttributeNS(W,'val')||base
  props.base=base
  const id=`PiStyle${digest(new TextEncoder().encode(JSON.stringify(props))).slice(0,24)}`
  let style=existing.find(n=>n.getAttributeNS(W,'styleId')===id)
  if(!style) {
   style=add(styles.documentElement,'style',{type:'paragraph',customStyle:'1',styleId:id})
   const displayName=existing.some(n=>child(n,'name')?.getAttributeNS(W,'val')===props.name)?`${props.name} · ${id.slice(-6)}`:props.name
   add(style,'name',{val:displayName});add(style,'aliases',{val:`PiFormat_${signature}`})
   if(existing.some(n=>n.getAttributeNS(W,'styleId')===base))add(style,'basedOn',{val:base})
   add(style,'uiPriority',{val:10});add(style,'qFormat')
   for(const [name,properties] of [['pPr',paragraphs],['rPr',fonts]])if(properties.length){const holder=add(style,name);for(const prop of properties)add(holder,prop.localName,prop.attributes)}
  }
  let link=child(pr,'pStyle');if(!link){link=document.createElementNS(W,'w:pStyle');pr.insertBefore(link,pr.firstChild)}
  link.setAttributeNS(W,'w:val',id)
  removeProperties(pr,paragraphs)
  // Visit runs and paragraph marks, but never enter nested drawing text boxes.
  const clear=n=>{
   if(n!==p&&n.namespaceURI===W&&['p','drawing','pict'].includes(n.localName))return
   if(n.namespaceURI===W&&['r','pPr'].includes(n.localName))removeProperties(child(n,'rPr'),fonts)
   for(const c of children(n))clear(c)
  }
  clear(p);applied.set(index,id)
 }
 const serialize=doc=>new TextEncoder().encode(new XMLSerializer().serializeToString(doc))
 parts['word/document.xml']=serialize(document);parts['word/styles.xml']=serialize(styles)
 const output=Buffer.from(zipSync(parts)),actual=readEditorDocument(output)
 const validation=compareProtected(bytes,output)
 if(actual.text!==original.text)throw problem('命名样式改变正文内容',422)
 for(const [index,id] of applied) {
  const before=expected.paragraphs[index],after=actual.paragraphs[index]
  if(after.styleId!==id)throw problem('命名样式关联重读失败',422)
  if(!isDeepStrictEqual(before.font,after.font)||!isDeepStrictEqual(before.format,after.format)||!isDeepStrictEqual(before.runs.map(r=>[r.font,r.effectiveFonts]),after.runs.map(r=>[r.font,r.effectiveFonts])))throw problem('样式继承与直接格式结果不一致，未提交',422)
 }
 // All non-target paragraphs retain their exact XML, in addition to package protection.
 const oldNodes=paragraphNodes(xml(unpack(bytes)['word/document.xml']))
 for(let i=0;i<nodes.length;i++)if(!applied.has(i)&&canonical(oldNodes[i])!==canonical(nodes[i]))throw problem('命名样式影响非目标段落',422)
 validation.checks.push('命名样式定义与段落关联','样式继承结果重读一致','非目标段落保持')
 return {bytes:output,affected:applied.size,validation}
}
