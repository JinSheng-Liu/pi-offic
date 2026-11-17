import {zipSync} from 'fflate'
import {XMLSerializer} from '@xmldom/xmldom'
import {unpack,xml,W,problem,compareProtected} from './package.mjs'
const children=(n,name)=>Array.from(n?.childNodes||[]).filter(c=>c.namespaceURI===W&&c.localName===name)
const number=(n,key)=>Number(n?.getAttributeNS(W,key)||0)

// Narrow fallback for table/grid widths missing from the SDK's table API.
// Every mutation has an undo used by the unchanged-content guard.
export function fitTableWidth(bytes,op,signal){
  signal?.throwIfAborted()
  const parts=unpack(bytes),doc=xml(parts['word/document.xml']),body=doc.getElementsByTagNameNS(W,'body')[0]
  const table=Array.from(body.getElementsByTagNameNS(W,'tbl'))[op.tableIndex]
  if(!table||table.parentNode!==body)throw problem('版心适配仅支持正文顶层表格',422)
  const rows=children(table,'tr'),grid=children(table,'tblGrid')[0],cols=children(grid,'gridCol')
  if(!cols.length||rows.length!==table.getElementsByTagNameNS(W,'tr').length)throw problem('嵌套或缺少列网格的表格需单独处理',422)
  let section
  for(let n=table.nextSibling;n;n=n.nextSibling){
    section=n.namespaceURI===W&&n.localName==='sectPr'?n:children(children(n,'pPr')[0],'sectPr')[0]
    if(section)break
  }
  const page=children(section,'pgSz')[0],margin=children(section,'pgMar')[0],columns=children(section,'cols')[0]
  if(!page||!margin||number(columns,'num')>1)throw problem('无法确认单栏版心',422)
  const pr=children(table,'tblPr')[0]
  if(!pr||children(pr,'tblpPr').length||children(pr,'tblCellSpacing').some(n=>number(n,'w')>0))throw problem('浮动表格或带单元格间距的表格需单独处理',422)
  const indent=children(pr,'tblInd')[0]
  if(indent&&indent.getAttributeNS(W,'type')!=='dxa')throw problem('表格缩进单位不明确',422)
  const available=number(page,'w')-number(margin,'left')-number(margin,'right')-number(margin,'gutter')-Math.max(0,number(indent,'w'))
  const widths=cols.map(c=>number(c,'w')),total=widths.reduce((a,b)=>a+b,0)
  if(available<100||widths.some(w=>w<=0))throw problem('表格网格或版心宽度无效',422)
  const ratios=op.columnRatios || widths
  if(ratios.length!==cols.length||ratios.some(v=>!Number.isFinite(v)||v<=0))throw problem('列宽比例与表格网格不匹配',422)
  const ratioTotal=ratios.reduce((a,b)=>a+b,0)
  const target=Math.min(total,available),scaled=ratios.map(w=>Math.floor(w*target/ratioTotal))
  if(scaled.some(w=>w<240))throw problem('调整后列宽过窄，不能安全排版',422)
  scaled[scaled.length-1]+=target-scaled.reduce((a,b)=>a+b,0)
  const undo=[]
  const attr=(n,key,value)=>{const old=n.getAttributeNS(W,key),had=n.hasAttributeNS(W,key);n.setAttributeNS(W,'w:'+key,String(value));undo.push(()=>had?n.setAttributeNS(W,'w:'+key,old):n.removeAttributeNS(W,key))}
  const width=(owner,group,name,value)=>{
    let p=group?children(owner,group)[0]:owner
    if(!p){p=doc.createElementNS(W,'w:'+group);owner.insertBefore(p,owner.firstChild);undo.push(()=>owner.removeChild(p))}
    let w=children(p,name)[0]
    if(!w){w=doc.createElementNS(W,'w:'+name);const before=group==='tblPr'?['jc','tblCellSpacing','tblInd','tblBorders','shd','tblLayout','tblCellMar','tblLook','tblCaption','tblDescription','tblPrChange']:['gridSpan','hMerge','vMerge','tcBorders','shd','noWrap','tcMar','textDirection','tcFitText','vAlign','hideMark','headers','cellIns','cellDel','cellMerge','tcPrChange'];const next=Array.from(p.childNodes).find(n=>n.namespaceURI===W&&before.includes(n.localName));p.insertBefore(w,next||null);undo.push(()=>p.removeChild(w))}
    attr(w,'w',value);attr(w,'type','dxa')
  }
  cols.forEach((c,i)=>attr(c,'w',scaled[i]))
  width(table,'tblPr','tblW',target)
  for(const row of rows){
    const rowPr=children(row,'trPr')[0]
    if(children(rowPr,'gridBefore').length||children(rowPr,'gridAfter').length)throw problem('带前后省略网格的表格暂不支持适配',422)
    let offset=0
    for(const cell of children(row,'tc')){
      const span=number(children(children(cell,'tcPr')[0],'gridSpan')[0],'val')||1
      if(span<1||offset+span>cols.length)throw problem('表格合并网格不一致',422)
      width(cell,'tcPr','tcW',scaled.slice(offset,offset+span).reduce((a,b)=>a+b,0));offset+=span
    }
    if(offset!==cols.length)throw problem('表格行未覆盖全部网格',422)
  }
  const serialize=()=>new TextEncoder().encode(new XMLSerializer().serializeToString(doc))
  const output=zipSync({...parts,'word/document.xml':serialize()})
  for(const restore of undo.reverse())restore()
  compareProtected(bytes,zipSync({...parts,'word/document.xml':serialize()}),{formatting:false})
  signal?.throwIfAborted()
  return {bytes:output,affected:1,validation:{status:'passed',checks:['表格列宽与单元格宽度按版心协调','表格非目标内容保持'],limitations:['宽度声明已适配；不可断行内容及实际溢出仍需页面核验']}}
}
