import {isCaptionLabel} from './caption-label.mjs'
import {zipSync} from 'fflate'
import {XMLSerializer} from '@xmldom/xmldom'
import {W,unpack,xml,canonical,compareProtected,problem} from './package.mjs'
const children=n=>Array.from(n.childNodes||[]).filter(c=>c.nodeType===1)
const text=n=>Array.from(n.getElementsByTagNameNS(W,'t')).map(t=>t.textContent||'').join('')
const labelPattern=/^(\s*)(图|表)(\s*)(?:(\d+)([.－-]))?(\d+)(?=\s|[^\d.－-]|$)/u

// Conversion changes only existing numeric text into a field, not caption content,
// paragraph formatting, bookmarks, or object positions. Chapter prefixes stay literal.
export function convertCaptions(bytes,op,signal) {
  const parts=unpack(bytes),doc=xml(parts['word/document.xml']),body=doc.getElementsByTagNameNS(W,'body')[0]
  const paragraphs=[]
  const walk=n=>{if(n.namespaceURI===W&&n.localName==='p'){paragraphs.push(n);return}for(const c of children(n))walk(c)}
  walk(body)
  const records=paragraphs.map((p,index)=>{const match=isCaptionLabel(text(p))?labelPattern.exec(text(p)):null;return {p,index,match,key:match?`${match[2]}:${match[4]||'global'}`:null}})
  const selected=op.indices.map(index=>records[index])
  const existingSequences=new Set()
  for(const [name,data] of Object.entries(parts))if(/^word\/.*\.xml$/.test(name)){
    const root=xml(data),stack=[]
    const remember=instruction=>{const match=/^\s*SEQ\s+(?:"([^"]+)"|([^\s\\]+))/i.exec(instruction);if(match)existingSequences.add((match[1]||match[2]).toLowerCase())}
    const visit=node=>{
      if(node.namespaceURI===W){
        if(node.localName==='fldSimple')remember(node.getAttributeNS(W,'instr')||'')
        if(node.localName==='fldChar'){
          const type=node.getAttributeNS(W,'fldCharType')
          if(type==='begin')stack.push({instruction:'',reading:true})
          else if(type==='separate'&&stack.length)stack.at(-1).reading=false
          else if(type==='end'&&stack.length)remember(stack.pop().instruction)
        }
        if(node.localName==='instrText'&&stack.at(-1)?.reading)stack.at(-1).instruction+=node.textContent||''
      }
      for(const child of children(node))visit(child)
    }
    visit(root)
    for(const field of stack)remember(field.instruction)
  }
  if(selected.some(r=>!r?.match))throw problem('题注索引不存在或没有明确的图表编号',422)
  const selectedIds=new Set(op.indices),groups=new Map()
  for(const r of selected){if(!groups.has(r.key))groups.set(r.key,records.filter(other=>other.key===r.key))}
  for(const group of groups.values()){
    if(group.some(r=>!selectedIds.has(r.index)))throw problem('必须完整选择同一图表编号组，不能转换一部分后让编号失序',422)
    group.forEach((r,i)=>{
      const number=Number(r.match[6])
      if(number!==i+1||String(number)!==r.match[6])throw problem('题注编号不从1连续递增或包含前导零，需先确认编号规则',422)
    })
  }
  const originals=[],converted=[]
  for(const record of selected.sort((a,b)=>a.index-b.index)){
    signal?.throwIfAborted()
    const {p,index,match}=record
    if(p.parentNode!==body)throw problem('只支持正文顶层手工题注，表格和内容控件内题注需单独处理',422)
    const nodes=children(p)
    if(nodes.some(n=>n.namespaceURI!==W||!['pPr','r','bookmarkStart','bookmarkEnd'].includes(n.localName)))throw problem('题注包含复杂结构，未转换',422)
    const runs=nodes.filter(n=>n.localName==='r')
    if(runs.some(r=>children(r).some(n=>n.namespaceURI!==W||!['rPr','t'].includes(n.localName))))throw problem('题注已有域或复杂内容，未重复转换',422)
    if(p.getElementsByTagNameNS(W,'sectPr').length)throw problem('题注包含分节信息',422)
    const prefix=match[1]+match[2]+match[3]+(match[4]?match[4]+match[5]:'')
    const start=prefix.length,end=start+match[6].length
    let offset=0;const segments=[]
    for(const node of nodes){
      if(['bookmarkStart','bookmarkEnd'].includes(node.localName)&&offset>start&&offset<end)throw problem('书签边界落在编号内部，不能安全转换',422)
      if(node.localName!=='r')continue
      for(const t of children(node).filter(n=>n.localName==='t')){
        const length=t.textContent.length
        if(offset<end&&offset+length>start)segments.push({run:node,t,from:Math.max(0,start-offset),to:Math.min(length,end-offset)})
        offset+=length
      }
    }
    if(!segments.length)throw problem('题注编号无法定位',422)
    const identifier=`Pi${match[2]==='图'?'Figure':'Table'}${match[4]||'Global'}`
    // Existing matching fields anywhere in the package require explicit migration.
    if(existingSequences.has(identifier.toLowerCase()))throw problem('文档已有同名编号序列，不能混入手工转换',422)
    originals.push({index,node:p.cloneNode(true)})
    const first=segments[0],last=segments.at(-1)
    // Split the boundary runs while retaining formatting and run metadata.
    const splitAt=(segment,at)=>{
      const run=segment.run,right=run.cloneNode(false),pr=children(run).find(n=>n.localName==='rPr')
      if(pr)right.appendChild(pr.cloneNode(true))
      const texts=children(run).filter(n=>n.localName==='t'),position=texts.indexOf(segment.t)
      const suffix=segment.t.cloneNode(true);suffix.textContent=segment.t.textContent.slice(at)
      segment.t.textContent=segment.t.textContent.slice(0,at)
      segment.t.setAttributeNS('http://www.w3.org/XML/1998/namespace','xml:space','preserve')
      suffix.setAttributeNS('http://www.w3.org/XML/1998/namespace','xml:space','preserve')
      right.appendChild(suffix)
      for(const t of texts.slice(position+1))right.appendChild(t)
      p.insertBefore(right,run.nextSibling)
      return right
    }
    const tail=splitAt(last,last.to)
    const numberRun=splitAt(first,first.from)
    const make=(kind,value)=>{const run=doc.createElementNS(W,'w:r'),node=doc.createElementNS(W,'w:'+kind);if(kind==='fldChar')node.setAttributeNS(W,'w:fldCharType',value);else{node.setAttributeNS('http://www.w3.org/XML/1998/namespace','xml:space','preserve');node.textContent=value}run.appendChild(node);return run}
    p.insertBefore(make('fldChar','begin'),numberRun)
    p.insertBefore(make('instrText',` SEQ ${identifier}${Number(match[6])===1?' \\r 1':''} \\* ARABIC `),numberRun)
    p.insertBefore(make('fldChar','separate'),numberRun)
    p.insertBefore(make('fldChar','end'),tail)
    converted.push({index,identifier,number:Number(match[6]),chapterPrefix:match[4]||null})
  }
  const serialize=()=>new TextEncoder().encode(new XMLSerializer().serializeToString(doc))
  const output=serialize(),reopened=xml(output)
  const reopenedParagraphs=[]
  const collect=n=>{if(n.namespaceURI===W&&n.localName==='p'){reopenedParagraphs.push(n);return}for(const c of children(n))collect(c)}
  collect(reopened.getElementsByTagNameNS(W,'body')[0])
  for(const original of originals){
    const after=reopenedParagraphs[original.index]
    if(text(after)!==text(original.node)||Array.from(after.getElementsByTagNameNS(W,'fldChar')).map(n=>n.getAttributeNS(W,'fldCharType')).join(',')!=='begin,separate,end')throw problem('题注文字或编号域回读验证失败',422)
    const expected=converted.find(c=>c.index===original.index)
    if(!after.getElementsByTagNameNS(W,'instrText')[0]?.textContent.includes(`SEQ ${expected.identifier} `))throw problem('题注序列回读验证失败',422)
  }
  for(const original of originals)paragraphs[original.index].parentNode.replaceChild(original.node,paragraphs[original.index])
  if(canonical(doc.documentElement)!==canonical(xml(parts['word/document.xml']).documentElement))throw problem('非目标题注正文发生变化，已取消保存',422)
  compareProtected(bytes,zipSync({...parts,'word/document.xml':serialize()}),{formatting:false})
  signal?.throwIfAborted()
  return {bytes:zipSync({...parts,'word/document.xml':output}),affected:converted.length,validation:{status:'passed',checks:['Word SEQ 域回读验证','题注文字与编号显示保持','非目标题注及包内容保持'],limitations:['仅转换编号域，图表对应关系和交叉引用更新未验证；章节前缀保持原文，不随章节移动自动更新；需在Word更新域验证编号']},captions:converted}
}
