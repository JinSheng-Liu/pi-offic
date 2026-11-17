import {XMLSerializer} from '@xmldom/xmldom'
import {zipSync} from 'fflate'
import {unpack,xml,W,problem,canonical,compareProtected} from './package.mjs'
import {readEditorDocument} from './editor-reader.mjs'
const children=(n,name)=>Array.from(n?.childNodes||[]).filter(c=>c.nodeType===1&&(!name||c.namespaceURI===W&&c.localName===name))
function context(bytes){
  const parts=unpack(bytes),doc=xml(parts['word/document.xml']),body=doc.getElementsByTagNameNS(W,'body')[0],ps=[]
  const walk=n=>{if(n.namespaceURI===W&&n.localName==='p'){ps.push(n);return}for(const c of children(n))walk(c)}
  walk(body)
  return {parts,doc,body,ps,reader:readEditorDocument(bytes)}
}
function widthAt(p,body){
  let block=p
  while(block.parentNode!==body&&block.parentNode)block=block.parentNode
  let section
  for(let n=block;n;n=n.nextSibling){
    section=n.namespaceURI===W&&n.localName==='sectPr'?n:children(children(n,'pPr')[0],'sectPr')[0]
    if(section)break
  }
  const size=children(section,'pgSz')[0],margin=children(section,'pgMar')[0],cols=children(section,'cols')[0]
  const num=(n,k)=>Number(n?.getAttributeNS(W,k)||0)
  const width=num(size,'w')-num(margin,'left')-num(margin,'right')-num(margin,'gutter')
  if(!size||!margin||num(cols,'num')>1||children(cols,'col').length>1||width<100||width>14400)throw problem('无法确认目标所在分节的单栏版心宽度',422)
  return width
}

// Reuse the existing native paragraph formatting boundary; never manufacture page numbers.
export async function formatTocEntries(bytes,signal,formatParagraphs){
  const {ps,body,reader}=context(bytes)
  const operations=[]
  for(const [index,p] of reader.paragraphs.entries()){
    if(!(p.generatedToc||/^TOC\d+$/i.test(p.styleId||''))||p.tableNestingLevel||!p.text.trim())continue
    if(!/^[^\t]+\t[0-9IVXLCDMivxlcdm]+\s*$/.test(p.text))continue
    const position=widthAt(ps[index],body)/20
    operations.push({scope:'indices',indices:[index],paragraph:{tocPageTabPosition:position,alignment:'Left',firstLineIndent:0,rightIndent:0}})
  }
  if(!operations.length)return {bytes,affected:0,validation:{status:'unverified',checks:['未找到可安全格式化的原生目录页码条目'],limitations:['未重算目录页码；非标准条目须单独检查']}}
  let result={bytes,affected:0}
  for(let start=0;start<operations.length;start+=30){
    const next=await formatParagraphs(result.bytes,operations.slice(start,start+30),signal)
    result={...next,affected:result.affected+next.affected}
  }
  result.validation.limitations.push('目录页码未重算；实际分页需页面验收')
  return result
}

// Narrow OOXML fallback: the editor has insertTable, but no lossless text-to-table API.
// Only plain runs are accepted. Each original run property is cloned into its cell.
export function tabularTextToTable(bytes,op,signal){
  const {parts,doc,body,ps,reader}=context(bytes),selected=op.indices.map(i=>ps[i])
  const original=canonical(doc.documentElement)
  const rows=[]
  for(const [rowIndex,p] of selected.entries()){
    signal?.throwIfAborted()
    const info=reader.paragraphs[op.indices[rowIndex]]
    if(!p||p.parentNode!==body||info?.generatedToc||/^TOC\d*$/i.test(info?.styleId||'')||info?.effectivePPr?.numPr||p.getElementsByTagNameNS(W,'sectPr').length)throw problem('仅支持正文顶层普通段落，不能转换目录、列表或分节段落',422)
    if(rowIndex&&children(body).indexOf(p)!==children(body).indexOf(selected[rowIndex-1])+1)throw problem('Tab 文本必须为相邻连续段落',422)
    if(children(p).some(n=>n.namespaceURI!==W||!['pPr','r'].includes(n.localName))||p.getElementsByTagNameNS(W,'pPrChange').length)throw problem('Tab 段落包含书签、域或复杂结构，不能安全转换',422)
    const cells=[[]]
    for(const run of children(p,'r')){
      if(children(run).some(n=>n.namespaceURI!==W||!['rPr','t','tab'].includes(n.localName))||run.getElementsByTagNameNS(W,'rPrChange').length)throw problem('Tab 段落含非纯文字内容，不能安全转换',422)
      let fragment
      for(const child of children(run).filter(n=>n.localName!=='rPr')){
        if(child.localName==='tab'){cells.push([]);fragment=undefined;continue}
        if((child.textContent||'').includes('\t'))throw problem('文本中的 Tab 编码不明确，不能安全转换',422)
        if(!fragment){fragment=run.cloneNode(false);const pr=children(run,'rPr')[0];if(pr)fragment.appendChild(pr.cloneNode(true));cells.at(-1).push(fragment)}
        fragment.appendChild(child.cloneNode(true))
      }
    }
    if(cells.length<2||cells.length>8||rows.length&&cells.length!==rows[0].length)throw problem('Tab 列数必须一致，且为 2–8 列',422)
    rows.push(cells)
  }
  const width=widthAt(selected.at(-1),body),count=rows[0].length
  const widths=Array.from({length:count},(_,i)=>Math.floor(width/count)+(i<width%count?1:0))
  const make=(name,attrs={})=>{const n=doc.createElementNS(W,'w:'+name);for(const [k,v] of Object.entries(attrs))n.setAttributeNS(W,'w:'+k,String(v));return n}
  const table=make('tbl'),pr=make('tblPr');table.appendChild(pr)
  pr.appendChild(make('tblW',{w:width,type:'dxa'}))
  const borders=make('tblBorders');for(const side of ['top','left','bottom','right','insideH','insideV'])borders.appendChild(make(side,{val:'nil'}));pr.appendChild(borders)
  pr.appendChild(make('tblLayout',{type:'fixed'}))
  const grid=make('tblGrid');widths.forEach(w=>grid.appendChild(make('gridCol',{w})));table.appendChild(grid)
  rows.forEach((cells,r)=>{
    const tr=make('tr');table.appendChild(tr)
    cells.forEach((runs,c)=>{
      const tc=make('tc'),tcPr=make('tcPr');tcPr.appendChild(make('tcW',{w:widths[c],type:'dxa'}));tc.appendChild(tcPr)
      const p=make('p'),oldPr=children(selected[r],'pPr')[0],pPr=oldPr?oldPr.cloneNode(true):make('pPr')
      for(const n of children(pPr).filter(n=>['tabs','ind','jc','keepNext','pageBreakBefore'].includes(n.localName)))pPr.removeChild(n)
      pPr.appendChild(make('ind',{left:0,right:0,firstLine:0}));pPr.appendChild(make('jc',{val:'left'}));p.appendChild(pPr)
      runs.forEach(run=>p.appendChild(run));tc.appendChild(p);tr.appendChild(tc)
    })
  })
  const next=selected.at(-1).nextSibling
  body.insertBefore(table,selected[0]);selected.forEach(p=>body.removeChild(p))
  const output=zipSync({...parts,'word/document.xml':new TextEncoder().encode(new XMLSerializer().serializeToString(doc))})
  const tableIndex=Array.from(body.getElementsByTagNameNS(W,'tbl')).indexOf(table)
  body.removeChild(table);selected.forEach(p=>body.insertBefore(p,next))
  if(canonical(doc.documentElement)!==original)throw problem('Tab 转表格改变了非目标结构',422)
  compareProtected(bytes,output,{formatting:false})
  signal?.throwIfAborted()
  return {bytes:output,tableIndex,affected:selected.length,validation:{status:'passed',checks:['连续等列 Tab 转为独立换行单元格','保留各列文字和原有字符格式','表格总宽等于所在分节版心','非目标结构保持'],limitations:['结构转换后必须重新读取段落索引','列宽均分；实际视觉和长词溢出仍需页面验收']}}
}
