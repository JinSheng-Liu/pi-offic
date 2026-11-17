import {
  readOoxmlPackage,bodyStoryRoot,storyParagraphs,stylesPartOf,propertyContainer,detectBodyTocs,
  authoredProperties,AUTHORABLE_RUN_PROPERTIES,
} from '@docx-editor.dev/core/store'
import {buildStyleCascadeTable,cascadeParagraphFormatting,cascadeRunProperties,resolveRunStyle} from '@docx-editor.dev/core/layout'
import {problem,W} from './package.mjs'

const children=node=>node?.children||[]
const attr=(node,key)=>node?.attributes?.find(a=>a.namespaceUri===W&&a.localName===key)?.value
const child=(node,name)=>children(node).find(n=>n.namespaceUri===W&&n.localName===name)
const on=value=>value!==undefined&&!['false','0','off','no'].includes(String(value))
function attrs(node){return Object.fromEntries((node?.attributes||[]).filter(a=>a.namespaceUri===W).map(a=>[a.localName,a.value]))}
function bag(node) {
  const out={}
  for(const n of children(node))if(n.namespaceUri===W) {
    const values=attrs(n)
    for(const [k,v] of Object.entries(values))if(/^-?\d+(\.\d+)?$/.test(v))values[k]=Number(v)
    out[n.localName]=children(n).length?{...values,...bag(n)}:Object.keys(values).length?values:{val:true}
    if(['keepNext','keepLines','widowControl','pageBreakBefore'].includes(n.localName))out[n.localName]={val:on(attr(n,'val')??true)}
    if(n.localName==='rFonts')out.rFonts=attrs(n)
    if(n.localName==='tabs')out.tabs={tab:children(n).filter(t=>t.localName==='tab').map(attrs).map(t=>({...t,pos:Number(t.pos)}))}
  }
  return out
}
function flatBag(properties) {
  const out={}
  for(const p of properties) {
    if(p.localName==='rFonts'&&out.rFonts){
      // Font slots form explicit/theme pairs. A later explicit slot replaces
      // its inherited theme; flattening both makes an inactive theme look live.
      for(const [slot,theme] of [['ascii','asciiTheme'],['hAnsi','hAnsiTheme'],['eastAsia','eastAsiaTheme'],['cs','cstheme']]){
        if(p.attributes?.[slot]!==undefined&&p.attributes?.[theme]===undefined)delete out.rFonts[theme]
        if(p.attributes?.[theme]!==undefined)delete out.rFonts[slot]
      }
    }
    out[p.localName]={...out[p.localName],...p.attributes}
    if(!p.attributes||!Object.keys(p.attributes).length)out[p.localName].val=true
  }
  return out
}
function textOf(node) {
  if(node.kind==='textValue')return node.value
  if(node.namespaceUri===W) {
    if(['instrText','delInstrText','delText','drawing','pict','pPr','rPr'].includes(node.localName))return ''
    if(node.localName==='tab')return '\t'
    if(node.localName==='sym'){const code=parseInt(attr(node,'char')||'',16);return Number.isFinite(code)&&code<=0x10ffff?String.fromCodePoint(code):'\ufffc'}
    if(node.localName==='noBreakHyphen')return '\u2011'
    if(node.localName==='softHyphen')return '\u00ad'
    if(node.localName==='cr')return '\r'
    if(node.localName==='br')return attr(node,'type')==='page'?'\f':attr(node,'type')==='column'?'\u000e':'\u000b'
  }
  return children(node).map(textOf).join('')
}
function fontFrom(properties) {
  const f=resolveRunStyle(properties)
  return {name:f.fontFamily||'',size:f.fontSizePt,bold:f.bold,italic:f.italic,color:`#${f.color||'000000'}`,
    underline:({none:'None',single:'Single',double:'Double',dotted:'Dotted',dash:'Dash',wave:'Wave'})[f.underline?.variant]||'None',
    strikeThrough:f.strike,doubleStrikeThrough:f.doubleStrike,subscript:f.verticalAlign==='subscript',superscript:f.verticalAlign==='superscript'}
}
export function commonFont(runs,fallback={}) {
  if(!runs.length)return fallback
  return Object.fromEntries(Object.keys(runs[0].font).map(key=>[key,runs.every(r=>r.font[key]===runs[0].font[key])?runs[0].font[key]:undefined]))
}
export function findTextRanges(paragraph,text) {
  const matches=[]
  for(let start=paragraph.text.indexOf(text);start!==-1;start=paragraph.text.indexOf(text,start+text.length)) {
    const end=start+text.length,runs=paragraph.runs.filter(r=>r.end>start&&r.start<end)
    matches.push({start,end,text,runs,font:commonFont(runs)})
  }
  return matches
}
// Business snapshot over the editor's canonical tree. The editor owns package
// parsing, paragraph traversal and style inheritance; this layer converts units
// and supplies the existing inspection/validation contract. It never writes bytes.
export function readEditorDocument(bytes) {
  const loaded=readOoxmlPackage(new Uint8Array(bytes))
  if(!loaded.ok)throw problem(`编辑器无法读取文档：${loaded.reason}`,422)
  const pkg=loaded.package,part=pkg.parts.get(pkg.mainDocumentPart),root=bodyStoryRoot(part)
  if(!root)throw problem('文档缺少正文',422)
  const cascade=buildStyleCascadeTable(stylesPartOf(pkg)?.root||null)
  const tocParagraphs=new Set(detectBodyTocs(part).flatMap(t=>[t.beginParagraphId,t.endParagraphId,...t.resultParagraphIds]))
  const parents=new Map()
  const walk=node=>{for(const c of children(node)){parents.set(c,node);walk(c)}}
  walk(root)
  const ancestor=(node,name)=>{for(let p=parents.get(node);p;p=parents.get(p))if(p.namespaceUri===W&&p.localName===name)return p}
  const cellCache=new Map()
  const paragraphs=storyParagraphs(root).map(node=>{
    const pr=propertyContainer(node,'paragraphProperties','pPr'),pPr=bag(pr)
    const inherited=cascadeParagraphFormatting(cascade,pr),effective=flatBag(inherited.paragraphProperties)
    // Nested tab definitions are read from the cascade's nodes, not flattened props.
    for(const props of inherited.paragraphPropertyNodes)if(child(props,'tabs'))effective.tabs=bag(props).tabs
    const runs=[]
    let offset=0
    const collect=n=>{
      if(n!==node&&n.namespaceUri===W&&['p','drawing','pict'].includes(n.localName))return
      if(n.namespaceUri===W&&n.localName==='r') {
        const rPr=propertyContainer(n,'runProperties','rPr'),text=textOf(n)
        const direct=[...authoredProperties(rPr,AUTHORABLE_RUN_PROPERTIES)]
        const runStyle=child(rPr,'rStyle');if(runStyle)direct.push({localName:'rStyle',attributes:attrs(runStyle)})
        if(text.length)runs.push({value:{rPr:bag(rPr)},font:fontFrom(cascadeRunProperties(inherited.runProperties,direct,cascade)),effectiveFonts:flatBag(cascadeRunProperties(inherited.runProperties,direct,cascade)).rFonts||{},text,start:offset,end:offset+text.length})
        offset+=text.length;return
      }
      for(const c of children(n))collect(c)
    }
    collect(node)
    const text=runs.map(r=>r.text).join(''),font=commonFont(runs,fontFrom(inherited.markRunProperties))
    const spacing=effective.spacing||{},ind=effective.ind||{}
    const points=value=>Number(value||0)/20
    const level=effective.outlineLvl?.val
    const outlineLevel=level!==undefined&&Number(level)>=0&&Number(level)<9?Number(level)+1:10
    const formatting={alignment:({center:'Centered',right:'Right',both:'Justified',distribute:'Justified'})[effective.jc?.val]||'Left',
      leftIndent:points(ind.left??ind.start),rightIndent:points(ind.right??ind.end),
      firstLineIndent:ind.hanging!==undefined?-points(ind.hanging):points(ind.firstLine),
      spaceBefore:points(spacing.before),spaceAfter:points(spacing.after),
      lineSpacing:points(spacing.line),outlineLevel}
    const inlinePictures=[]
    const pictures=n=>{if(n!==node&&n.localName==='p'&&n.namespaceUri===W)return;if(n.namespaceUri==='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'&&n.localName==='inline')inlinePictures.push(n);for(const c of children(n))pictures(c)}
    pictures(node)
    let nesting=0
    for(let n=parents.get(node);n;n=parents.get(n))if(n.namespaceUri===W&&n.localName==='tbl')nesting++
    const cell=ancestor(node,'tc')
    if(cell&&!cellCache.has(cell)) {
      const row=ancestor(cell,'tr'),table=ancestor(row,'tbl'),width=child(child(cell,'tcPr'),'tcW')
      cellCache.set(cell,{tc:cell,rowIndex:children(table).filter(n=>n.namespaceUri===W&&n.localName==='tr').indexOf(row),
        cellIndex:children(row).filter(n=>n.namespaceUri===W&&n.localName==='tc').indexOf(cell),
        parentRow:{isHeader:!!child(child(row,'trPr'),'tblHeader')&&on(attr(child(child(row,'trPr'),'tblHeader'),'val')??true)},
        width:attr(width,'type')==='dxa'?Number(attr(width,'w'))/20:0,text:textOf(cell)})
    }
    return {styleName:attr(child(children(stylesPartOf(pkg)?.root).find(s=>attr(s,'styleId')===inherited.styleId),'name'),'val')||inherited.styleId||'Normal',effectiveMarkFonts:flatBag(inherited.markRunProperties).rFonts||{},nodeId:node.id,generatedToc:tocParagraphs.has(node.id),text,runs,font,p:{pPr},effectivePPr:effective,styleId:inherited.styleId||'Normal',
      ...formatting,format:formatting,inlinePictures,tableNestingLevel:nesting,parentTableCell:cellCache.get(cell)}
  })
  return {paragraphs,text:paragraphs.map(p=>p.text).join('\r')}
}

