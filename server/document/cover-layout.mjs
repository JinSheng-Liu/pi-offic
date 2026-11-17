import {zipSync} from 'fflate'
import {XMLSerializer} from '@xmldom/xmldom'
import {unpack,xml,W,canonical,problem} from './package.mjs'
const children=n=>Array.from(n.childNodes||[]).filter(c=>c.nodeType===1)
const text=n=>Array.from(n.getElementsByTagNameNS(W,'t')).map(t=>t.textContent||'').join('')
const normalized=s=>s.replace(/[ \u3000\t\r\n]/g,'')

// Explicit structural operation: only plain label/value paragraphs may be moved.
// Whitespace used as positioning is intentionally removed; every other character
// and its run properties are retained. No content is generated from a template.
export function layoutCover(bytes,op,signal){
  signal?.throwIfAborted()
  const parts=unpack(bytes),doc=xml(parts['word/document.xml']),body=doc.getElementsByTagNameNS(W,'body')[0],ps=[]
  const collect=n=>{if(n.namespaceURI===W&&n.localName==='p'){ps.push(n);return}for(const c of children(n))collect(c)}
  collect(body)
  const selected=op.indices.map(i=>ps[i])
  if(selected.some(p=>!p||p.parentNode!==body))throw problem('封面必须选择正文顶层普通段落',422)
  const blocks=children(body),start=blocks.indexOf(selected[0])
  if(selected.some((p,i)=>blocks[start+i]!==p))throw problem('封面布局段落必须连续并按文档顺序选择',422)
  const activeSection=blocks.slice(start).flatMap(n=>n.localName==='sectPr'?[n]:Array.from(n.getElementsByTagNameNS(W,'sectPr')))[0]
  const page=activeSection?.getElementsByTagNameNS(W,'pgSz')[0],margin=activeSection?.getElementsByTagNameNS(W,'pgMar')[0]
  if(page&&margin){const available=Number(page.getAttributeNS(W,'w'))-['left','right','gutter'].reduce((sum,k)=>sum+Number(margin.getAttributeNS(W,k)||0),0);if(op.cover.widthPt*20>available)throw problem('封面布局宽度超过当前节正文区域',422)}
  const rows=[]
  let boundaryParagraph=null
  const last=selected.at(-1),lastNodes=children(last),section=last.getElementsByTagNameNS(W,'sectPr')[0]
  const endings=lastNodes.filter(n=>n.namespaceURI===W&&n.localName==='bookmarkEnd')
  if(section||endings.length){
    if(endings.length&&lastNodes.slice(lastNodes.indexOf(endings[0])).some(n=>n.localName!=='bookmarkEnd'))throw problem('封面书签边界位于内容内部，不能重排',422)
    boundaryParagraph=doc.createElementNS(W,'w:p')
    const boundaryPr=doc.createElementNS(W,'w:pPr')
    // A small empty paragraph carries the original section and bookmark endings.
    const spacing=doc.createElementNS(W,'w:spacing');for(const [k,v]of Object.entries({before:0,after:0,line:20,lineRule:'exact'}))spacing.setAttributeNS(W,'w:'+k,String(v));boundaryPr.appendChild(spacing)
    if(section)boundaryPr.appendChild(section.cloneNode(true))
    boundaryParagraph.appendChild(boundaryPr);for(const n of endings)boundaryParagraph.appendChild(n.cloneNode(true))
  }
  for(const p of selected){
    if(children(p).some(n=>n.namespaceURI!==W||!['pPr','r',...(p===last?['bookmarkEnd']:[])].includes(n.localName))||(p!==last&&p.getElementsByTagNameNS(W,'sectPr').length)||p.getElementsByTagNameNS(W,'numPr').length||p.getElementsByTagNameNS(W,'pPrChange').length)throw problem('封面包含复杂结构、编号或分节，不能安全重排',422)
    if(children(p).filter(n=>n.localName==='r').some(r=>children(r).some(n=>n.namespaceURI!==W||!['rPr','t'].includes(n.localName))))throw problem('封面含域、图片、制表符或修订，需单独处理',422)
    const value=text(p),match=/^\s*(题\s*目|学生姓名|学\s*号|学\s*院|专业班级|指导教师[^：:]*|起止时间)[：:]/u.exec(value)
    if(match){
      const colon=value.indexOf(match[0].includes('：')?'：':':')
      if(!value.slice(colon+1).trim())throw problem('封面信息值为空，不能自动重排',422)
      rows.push({label:{p,start:0,end:colon+1},values:[{p,start:colon+1,end:value.length}],title:/题\s*目/.test(match[1])})
    }else if(rows.length===1&&rows[0].title&&value.trim())rows[0].values.push({p,start:0,end:value.length})
    else throw problem('封面段落不是明确标签或题名续行',422)
  }
  if(rows.length<2)throw problem('封面重排至少需要两项信息',422)
  const make=(name,attrs={})=>{const n=doc.createElementNS(W,'w:'+name);for(const [k,v]of Object.entries(attrs))n.setAttributeNS(W,'w:'+k,String(v));return n}
  const tbl=make('tbl'),pr=make('tblPr'),width=Math.round(op.cover.widthPt*20),label=Math.round(op.cover.labelWidthPt*20)
  pr.appendChild(make('tblW',{w:width,type:'dxa'}));pr.appendChild(make('jc',{val:'center'}))
  const borders=make('tblBorders');for(const edge of ['top','left','bottom','right','insideH','insideV'])borders.appendChild(make(edge,{val:'nil'}));pr.appendChild(borders)
  pr.appendChild(make('tblLayout',{type:'fixed'}));const margins=make('tblCellMar');for(const edge of ['top','left','bottom','right'])margins.appendChild(make(edge,{w:0,type:'dxa'}));pr.appendChild(margins);tbl.appendChild(pr)
  const grid=make('tblGrid');for(const w of [label,width-label])grid.appendChild(make('gridCol',{w}));tbl.appendChild(grid)
  const sliceRuns=(part,removeSpaces=false)=>{
    let offset=0;const output=[];const source=text(part.p),segment=source.slice(part.start,part.end),leading=segment.length-segment.trimStart().length,trailing=segment.length-segment.trimEnd().length
    for(const run of children(part.p).filter(n=>n.localName==='r'))for(const t of children(run).filter(n=>n.localName==='t')){
      const value=t.textContent||'',from=Math.max(0,part.start+leading-offset),to=Math.min(value.length,part.end-trailing-offset);offset+=value.length
      if(to<=from)continue
      let fragment=value.slice(from,to);if(removeSpaces)fragment=fragment.replace(/[ \u3000]/g,'');if(!fragment)continue
      const r=run.cloneNode(false),rpr=children(run).find(n=>n.localName==='rPr');if(rpr)r.appendChild(rpr.cloneNode(true))
      const node=make('t');node.setAttributeNS('http://www.w3.org/XML/1998/namespace','xml:space','preserve');node.textContent=fragment;r.appendChild(node);output.push(r)
    }return output
  }
  for(const [i,row]of rows.entries()){
    const tr=make('tr'),trPr=make('trPr');trPr.appendChild(make('cantSplit'));tr.appendChild(trPr)
    for(const [j,w]of [label,width-label].entries()){
      const tc=make('tc'),tcPr=make('tcPr');tcPr.appendChild(make('tcW',{w,type:'dxa'}));tcPr.appendChild(make('vAlign',{val:'top'}));tc.appendChild(tcPr)
      const p=make('p'),ppr=make('pPr'),originalPr=children((j===0?row.label:row.values[0]).p).find(n=>n.localName==='pPr')
      const style=originalPr&&children(originalPr).find(n=>n.localName==='pStyle'),font=originalPr&&children(originalPr).find(n=>n.localName==='rPr')
      if(style)ppr.appendChild(style.cloneNode(true));if(i<rows.length-1)ppr.appendChild(make('keepNext'));ppr.appendChild(make('keepLines'));ppr.appendChild(make('spacing',{before:0,after:row.title?240:160,line:300,lineRule:'auto'}));ppr.appendChild(make('ind',{left:0,right:0,firstLine:0}));ppr.appendChild(make('jc',{val:'left'}));if(font)ppr.appendChild(font.cloneNode(true));p.appendChild(ppr)
      for(const part of j===0?[row.label]:row.values)for(const r of sliceRuns(part,j===0))p.appendChild(r)
      tc.appendChild(p);tr.appendChild(tc)
    }tbl.appendChild(tr)
  }
  if(normalized(text(tbl))!==normalized(selected.map(text).join('')))throw problem('封面重排文字校验失败',422)
  body.insertBefore(tbl,selected[0]);if(boundaryParagraph)body.insertBefore(boundaryParagraph,selected[0]);for(const p of selected)body.removeChild(p)
  const output=new TextEncoder().encode(new XMLSerializer().serializeToString(doc))
  // Restore original nodes to prove that all surrounding content is untouched.
  for(const p of selected)body.insertBefore(p,tbl);body.removeChild(tbl);if(boundaryParagraph)body.removeChild(boundaryParagraph)
  if(canonical(doc.documentElement)!==canonical(xml(parts['word/document.xml']).documentElement))throw problem('封面重排影响非目标内容',422)
  const reopened=xml(output);if(normalized(text(reopened))!==normalized(text(doc)))throw problem('封面保存回读内容不一致',422)
  signal?.throwIfAborted()
  return {bytes:zipSync({...parts,'word/document.xml':output}),affected:selected.length,validation:{status:'partial',checks:['封面标签与内容两列布局','除定位空白外文字保持','非目标正文与包部件保持','保存回读验证'],limitations:['封面定位空格被移除，题名续段合并；段落与表格索引已变化，必须重新读取','页面布局需渲染确认；本操作不是学校模板合规认证']}}
}
