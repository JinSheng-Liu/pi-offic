import { zipSync } from 'fflate'
import { XMLSerializer } from '@xmldom/xmldom'
import { unpack, xml, W, V, canonical, compareProtected, problem } from './package.mjs'
const captionNumber = text => /^(?:图|表)\s*(\d+)[-－.](\d+)/u.exec(text)
const chapterNumber = text => {
  const match = /^\s*(?:第\s*)?([0-9]+|[一二三四五六七八九十]+)\s*章/u.exec(text)
  if (!match) return undefined
  const value = match[1]
  if (/^\d+$/.test(value)) return String(Number(value))
  const digits = {一:1,二:2,三:3,四:4,五:5,六:6,七:7,八:8,九:9,十:10}
  if(digits[value]) return String(digits[value])
  if(/^[一二三四五六七八九]?十[一二三四五六七八九]?$/u.test(value)) {
    const [tens,ones]=value.split('十'); return String((digits[tens]||1)*10+(digits[ones]||0))
  }
  return undefined
}

function wordText(node) { return Array.from(node.getElementsByTagNameNS(W,'t')).map(t=>t.textContent||'').join('') }

function fieldRun(document, type) {
  const run=document.createElementNS(W,'w:r'), field=document.createElementNS(W,'w:fldChar')
  field.setAttributeNS(W,'w:fldCharType',type); run.appendChild(field); return run
}

function textRun(document, text, kind='t') {
  const run=document.createElementNS(W,'w:r'), node=document.createElementNS(W,`w:${kind}`)
  if (kind==='instrText') node.setAttributeNS('http://www.w3.org/XML/1998/namespace','xml:space','preserve')
  node.appendChild(document.createTextNode(text)); run.appendChild(node); return run
}

function makeCaption(document, kind, chapter, sequence, text, reset) {
  const paragraph=document.createElementNS(W,'w:p'), pPr=document.createElementNS(W,'w:pPr')
  const style=document.createElementNS(W,'w:pStyle'); style.setAttributeNS(W,'w:val','Caption'); pPr.appendChild(style)
  const align=document.createElementNS(W,'w:jc'); align.setAttributeNS(W,'w:val','center'); pPr.appendChild(align)
  if(kind==='table')pPr.appendChild(document.createElementNS(W,'w:keepNext'))
  pPr.appendChild(document.createElementNS(W,'w:keepLines'))
  paragraph.appendChild(pPr)
  paragraph.appendChild(textRun(document,`${kind==='figure'?'图':'表'}${chapter}-`))
  paragraph.appendChild(fieldRun(document,'begin'))
  paragraph.appendChild(textRun(document,` SEQ Pi${kind==='figure'?'Figure':'Table'}${chapter}${reset?` \\r ${sequence}`:''} \\* ARABIC `,'instrText'))
  paragraph.appendChild(fieldRun(document,'separate'))
  paragraph.appendChild(textRun(document,String(sequence)))
  paragraph.appendChild(fieldRun(document,'end'))
  paragraph.appendChild(textRun(document,` ${text}`))
  return paragraph
}

export function insertCaptions(bytes, captions, signal) {
  if (!Array.isArray(captions)||!captions.length||captions.length>20) throw problem('需要1–20个题注')
  for (const caption of captions) if (!caption||!['figure','table'].includes(caption.kind)||!Number.isInteger(caption.targetIndex)||caption.targetIndex<0||typeof caption.text!=='string'||!caption.text.trim()||caption.text.length>300||/[\r\n\u0000-\u001f]/u.test(caption.text)) throw problem('题注类型、目标索引或题注文字无效')
  signal?.throwIfAborted()
  const parts=unpack(bytes), document=xml(parts['word/document.xml']), body=document.getElementsByTagNameNS(W,'body')[0]
  const blocks=Array.from(body.childNodes).filter(node=>node.nodeType===1&&node.namespaceURI===W)
  const topParagraphs=blocks.filter(node=>node.localName==='p'), topTables=blocks.filter(node=>node.localName==='tbl')
  const stylesDocument=parts['word/styles.xml']?xml(parts['word/styles.xml']):null
  const headingStyles=new Set(Array.from(stylesDocument?.getElementsByTagNameNS(W,'style')||[])
    .filter(style=>style.getAttributeNS(W,'styleId')==='Heading1').map(style=>style.getAttributeNS(W,'styleId')))
  const headingTextBefore = target => {
    const index=blocks.indexOf(target)
    for(let i=index-1;i>=0;i--) if(blocks[i].localName==='p') {
      const pPr=Array.from(blocks[i].childNodes).find(n=>n.namespaceURI===W&&n.localName==='pPr')
      const style=Array.from(pPr?.childNodes||[]).find(n=>n.namespaceURI===W&&n.localName==='pStyle')?.getAttributeNS(W,'val')
      const level=Array.from(pPr?.childNodes||[]).find(n=>n.namespaceURI===W&&n.localName==='outlineLvl')?.getAttributeNS(W,'val')
      if(headingStyles.has(style)||level==='0') return chapterNumber(wordText(blocks[i]))
    }
  }
  const actions=[], staged=new Map(), inserted=[]
  const oldFields=Array.from(document.getElementsByTagNameNS(W,'instrText')).filter(n=>/SEQ Pi(?:Figure|Table)\d+/.test(n.textContent||'')).length
  const requestedTargets=new Set()
  const ordered=[...captions].sort((a,b)=>blocks.indexOf(a.kind==='figure'?topParagraphs[a.targetIndex]:topTables[a.targetIndex])-blocks.indexOf(b.kind==='figure'?topParagraphs[b.targetIndex]:topTables[b.targetIndex]))
  for(const item of ordered) {
    const target=item.kind==='figure'?topParagraphs[item.targetIndex]:topTables[item.targetIndex]
    if(!target) throw problem('题注目标索引超出范围',422)
    const targetKey=`${item.kind}:${item.targetIndex}`
    if(requestedTargets.has(targetKey)) throw problem('同一图表不能重复添加题注',422)
    requestedTargets.add(targetKey)
    if(item.kind==='figure'&&target.getElementsByTagNameNS(W,'drawing').length!==1) throw problem('目标不是含单个图片的正文段落',422)
    if(item.kind==='figure'&&target.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing','anchor').length) throw problem('暂不支持浮动图片题注',422)
    if(item.kind==='figure'&&target.getElementsByTagNameNS(V,'imagedata').length) throw problem('暂不支持VML图片题注',422)
    const chapter=headingTextBefore(target)
    if(!chapter) throw problem('无法从前置一级标题可靠识别章节号，未插入题注',422)
    const label=item.kind==='figure'?'图':'表'
    const candidates=blocks.filter(node=>node.localName==='p'&&node!==target&&captionNumber(wordText(node))?.[1]===chapter&&wordText(node).startsWith(label))
    for(const node of candidates) {
      const cmp=captionNumber(wordText(node)), order=blocks.indexOf(node), targetOrder=blocks.indexOf(target)
      if(cmp&&Number(cmp[2])>0&&order>targetOrder) throw problem('目标位于已有同章题注之前；为避免旧编号失序，请先将现有题注转换为自动编号',422)
    }
    const adjacent=item.kind==='figure'?blocks[blocks.indexOf(target)+1]:blocks[blocks.indexOf(target)-1]
    if(adjacent?.localName==='p'&&(captionNumber(wordText(adjacent))||Array.from(adjacent.getElementsByTagNameNS(W,'pStyle')).some(n=>n.getAttributeNS(W,'val')==='Caption'))) throw problem('目标附近已有题注，未重复插入',422)
    const key=`${item.kind}:${chapter}`
    const sequence=candidates.reduce((max,node)=>Math.max(max,Number(captionNumber(wordText(node))?.[2]||0)),0)+(staged.get(key)||0)+1
    staged.set(key,(staged.get(key)||0)+1)
    const managed=candidates.some(node=>Array.from(node.getElementsByTagNameNS(W,'instrText')).some(n=>(n.textContent||'').includes(`SEQ Pi${item.kind==='figure'?'Figure':'Table'}${chapter} `)))
    actions.push({target,item,chapter,sequence,reset:!managed&&staged.get(key)===1})
  }
  for(const action of actions) {
    const node=makeCaption(document,action.item.kind,action.chapter,action.sequence,action.item.text,action.reset)
    const target=action.target, reference=action.item.kind==='figure'?target.nextSibling:target
    body.insertBefore(node,reference)
    inserted.push(node)
  }
  const captionParagraphs=Array.from(document.getElementsByTagNameNS(W,'p')).filter(p=>Array.from(p.getElementsByTagNameNS(W,'instrText')).some(n=>/SEQ Pi(?:Figure|Table)\d+/.test(n.textContent||'')))
  if(captionParagraphs.length!==actions.length+oldFields) throw problem('题注字段写入校验失败',422)
  let styles=stylesDocument
  if(!styles) throw problem('缺少 Word 样式定义，无法安全创建题注样式',422)
  if(!Array.from(styles.getElementsByTagNameNS(W,'style')).some(style=>style.getAttributeNS(W,'styleId')==='Caption')) {
    const style=styles.createElementNS(W,'w:style'); style.setAttributeNS(W,'w:type','paragraph'); style.setAttributeNS(W,'w:styleId','Caption')
    const name=styles.createElementNS(W,'w:name'); name.setAttributeNS(W,'w:val','caption'); style.appendChild(name)
    const basedOn=styles.createElementNS(W,'w:basedOn'); basedOn.setAttributeNS(W,'w:val','Normal'); style.appendChild(basedOn)
    const pPr=styles.createElementNS(W,'w:pPr')
    pPr.appendChild(styles.createElementNS(W,'w:keepLines'))
    style.appendChild(pPr); styles.documentElement.appendChild(style)
    parts['word/styles.xml']=new TextEncoder().encode(new XMLSerializer().serializeToString(styles))
  }
  const oldDocument=xml(parts['word/document.xml']), outputDocument=document.cloneNode(true)
  const positions=inserted.map(node=>Array.from(body.childNodes).indexOf(node)).sort((a,b)=>b-a)
  const outputBody=outputDocument.getElementsByTagNameNS(W,'body')[0]
  for(const position of positions)outputBody.removeChild(outputBody.childNodes[position])
  if(canonical(oldDocument.documentElement)!==canonical(outputDocument.documentElement)) throw problem('插入题注时发现非目标正文结构变化',422)
  parts['word/document.xml']=new TextEncoder().encode(new XMLSerializer().serializeToString(document))
  const output=zipSync(parts), protectedParts={...parts,'word/document.xml':new TextEncoder().encode(new XMLSerializer().serializeToString(oldDocument))}
  const result=compareProtected(bytes,zipSync(protectedParts))
  signal?.throwIfAborted()
  result.checks.push('题注定位与章节编号验证','Word SEQ 域保存验证','非题注正文结构保持')
  return {bytes:output,affected:actions.length,validation:result,captions:actions.map(({item,chapter,sequence})=>({kind:item.kind,targetIndex:item.targetIndex,chapter,sequence,text:item.text}))}
}

