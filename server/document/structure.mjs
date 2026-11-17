import {inspectFooters} from './footer-inspection.mjs'
import {isCaptionLabel} from './caption-label.mjs'
import {readEditorDocument} from './editor-reader.mjs'
import { unpack, xml, W, A, M, V, digest } from './package.mjs'
import { paginationKeys } from './rules.mjs'
import { paragraphRole } from './paragraph-role.mjs'
import { inspectLayoutRisks } from './layout-risks.mjs'
import { captionFieldDetails } from './caption-fields.mjs'
import { databaseTableRecommendation, analysisTableRecommendation } from './table-recommendation.mjs'
function layoutDetails(p, defaultTabStop) {
  const effective = p.effectivePPr
  const spacing = effective?.spacing
  const rule = spacing?.lineRule || (spacing?.line != null ? 'auto' : 'unspecified')
  const raw = spacing?.line
  const tabs = value => (value?.tab || []).map(t => ({
    positionTwips: t.pos, alignment: t.val, leader: t.leader || 'none',
  }))
  return {
    lineSpacing: { rule, ...(raw != null ? { rawValue: raw,
      ...(rule === 'auto' ? { multiple: Number(raw) / 240 } : { points: Number(raw) / 20 }) } : {}),
      direct: { line: p.p.pPr?.spacing?.line ?? null, rule: p.p.pPr?.spacing?.lineRule ?? null } },
    tabStops: { direct: tabs(p.p.pPr?.tabs), effective: tabs(effective?.tabs), defaultIntervalTwips: defaultTabStop,
      note: '位置相对段落区域；clear 表示清除制表位。字符宽度、列表及环绕仍影响视觉起点，不能仅据此保证页面效果。' },
  }
}
export function sectionNodes(documentBody) {
  return Array.from(documentBody.getElementsByTagNameNS(W, 'sectPr')).filter(section => {
    let parent = section.parentNode
    while (parent && parent !== documentBody) {
      if (parent.namespaceURI === W && parent.localName === 'sectPrChange') return false
      parent = parent.parentNode
    }
    return parent === documentBody
  })
}

export async function inspect(bytes) {
  const parts = unpack(bytes)
  const body = readEditorDocument(bytes)
  const mainDocument = xml(parts['word/document.xml'])
  const documentBody = mainDocument.getElementsByTagNameNS(W, 'body')[0]
  const tabSetting = parts['word/settings.xml'] ? xml(parts['word/settings.xml']).getElementsByTagNameNS(W, 'defaultTabStop')[0] : undefined
  const defaultTabStop = tabSetting ? Number(tabSetting.getAttributeNS(W, 'val')) : null
  const sdkParagraphs = body.paragraphs
  const cellIds = new Map()
  const cells = []
  for (const [index, p] of sdkParagraphs.entries()) {
    const cell = p.parentTableCell
    if (!cell) continue
    if (!cellIds.has(cell.tc)) {
      cellIds.set(cell.tc, cells.length)
      cells.push({ index: cells.length, rowIndex: cell.rowIndex, cellIndex: cell.cellIndex,
        nestingLevel: p.tableNestingLevel, isHeader: cell.parentRow.isHeader,
        widthPoints: cell.width, paragraphIndices: [], text: cell.text.slice(0, 500) })
    }
    cells[cellIds.get(cell.tc)].paragraphIndices.push(index)
  }
  const paragraphs = sdkParagraphs.map((p, index) => ({ index, address: `body/${index}`, tableNestingLevel: p.tableNestingLevel,
    text: p.text.slice(0, 1000), textTruncated: p.text.length > 1000, style: p.styleId, styleName:p.styleName,
    role: paragraphRole(p), inlinePictureCount: p.inlinePictures.length,
    cell: p.parentTableCell ? { id: cellIds.get(p.parentTableCell.tc), rowIndex: p.parentTableCell.rowIndex, cellIndex: p.parentTableCell.cellIndex } : undefined,
    leadingWhitespace: { tabs: (p.text.match(/^\t+/)?.[0] || '').length, prefix: p.text.match(/^[\t \u3000]+/)?.[0] || '' },
    indentWarning: p.text.startsWith('\t') && (p.firstLineIndent !== 0 || p.leftIndent !== 0)
      ? '前导制表符与段落缩进同时存在，实际首行位置由制表位决定，需检查是否重复缩进；不要批量删除标题或列表的合法缩进' : undefined,
    heading: p.outlineLevel <= 9 ? p.outlineLevel : 0,
    font: { effectiveFonts:p.runs.length?p.runs.map(r=>r.effectiveFonts):[p.effectiveMarkFonts], paragraphMarkFonts:p.p.pPr?.rPr?.rFonts ? Object.fromEntries(['ascii','hAnsi','eastAsia'].map(key=>[key,p.p.pPr.rPr.rFonts[key]??null])) : null, name: p.font.name, size: p.font.size, bold: p.font.bold, italic: p.font.italic, color:p.font.color, underline:p.font.underline,
      asciiDirect: [...new Set(p.runs.map(run => run.value.rPr?.rFonts?.ascii ?? null))],
      hAnsiDirect: [...new Set(p.runs.map(run => run.value.rPr?.rFonts?.hAnsi ?? null))],
      eastAsiaDirect: [...new Set(p.runs.map(run => run.value.rPr?.rFonts?.eastAsia ?? null))],
      eastAsiaThemes: [...new Set(p.runs.map(run => run.value.rPr?.rFonts?.eastAsiaTheme ?? null))] },
    numbering: p.p?.pPr?.numPr ? { level: p.p.pPr.numPr.ilvl?.val, id: p.p.pPr.numPr.numId?.val } : undefined,
    ...layoutDetails(p, defaultTabStop),
    pagination: Object.fromEntries([...paginationKeys].map(key => [key, p.effectivePPr?.[key] ? ![false,'false',0,'0','off'].includes(p.effectivePPr[key].val) : null])),
    paragraph: p.format }))
  const captionFields = captionFieldDetails(mainDocument)
  const captions = paragraphs.filter(p => p.style === 'Caption' || isCaptionLabel(p.text)).map(p=>({index:p.index,kind:/^\s*图/u.test(p.text)?'figure':/^\s*表/u.test(p.text)?'table':'unknown',text:p.text,style:p.style,
    ...(captionFields.length===paragraphs.length?captionFields[p.index]:{numbering:'manual-or-unverified',sequenceIdentifiers:[]}),
    associationVerified:false}))
  const elements = (node, name) => Array.from(node?.getElementsByTagNameNS(W, name) || [])
  const attr = (node, name) => node?.getAttributeNS(W, name) || undefined
  const children = (node, name) => Array.from(node?.childNodes || []).filter(child => child.nodeType === 1 && child.namespaceURI === W && child.localName === name)
  const sectionElements = sectionNodes(documentBody)
  const sections = sectionElements.map((section, index) => {
    const size = children(section, 'pgSz')[0], margins = children(section, 'pgMar')[0]
    const refs = name => children(section, name).map(reference => ({ type: attr(reference, 'type') || '', relationshipId: reference.getAttribute('r:id') || reference.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') }))
    return { index, page: size ? { widthTwips: Number(attr(size, 'w')), heightTwips: Number(attr(size, 'h')), orientation: attr(size, 'orient') || 'portrait' } : undefined,
      margins: margins ? Object.fromEntries(['top','right','bottom','left','header','footer','gutter'].map(key => [key, Number(attr(margins, key) || 0)])) : undefined,
      headers: refs('headerReference'), footers: refs('footerReference') }
  })
  const pageNumbers=inspectFooters(parts,sectionElements)
  const tableElements = elements(documentBody, 'tbl')
  const blockParagraphs=[]
  const collect=node=>{if(node.namespaceURI===W&&node.localName==='p'){blockParagraphs.push(node);return}for(const child of Array.from(node.childNodes||[]))collect(child)}
  collect(documentBody)
  for(const caption of captions){
    const node=blockParagraphs[caption.index]
    if(!node||node.parentNode!==documentBody)continue
    let neighbor=caption.kind==='figure'?node.previousSibling:node.nextSibling
    while(neighbor&&neighbor.nodeType!==1)neighbor=caption.kind==='figure'?neighbor.previousSibling:neighbor.nextSibling
    caption.associationVerified=caption.kind==='table'?neighbor?.namespaceURI===W&&neighbor.localName==='tbl':caption.kind==='figure'&&neighbor?.namespaceURI===W&&neighbor.localName==='p'&&elements(neighbor,'t').every(t=>!t.textContent.trim())&&neighbor.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing','inline').length===1
  }
  const paragraphIndices=new Map()
  if(blockParagraphs.length===sdkParagraphs.length)blockParagraphs.forEach((p,i)=>{if(elements(p,'t').map(t=>t.textContent||'').join('')===sdkParagraphs[i].text)paragraphIndices.set(p,i)})
  const on = node => Boolean(node) && !['0','false','off','no'].includes((attr(node,'val') || 'true').toLowerCase())
  const tables = tableElements.map((table, index) => {
    const rows = children(table, 'tr')
    let availableWidthTwips=null
    if(table.parentNode===documentBody){
      let section
      for(let n=table.nextSibling;n;n=n.nextSibling){section=n.namespaceURI===W&&n.localName==='sectPr'?n:children(children(n,'pPr')[0],'sectPr')[0];if(section)break}
      const page=children(section,'pgSz')[0],margin=children(section,'pgMar')[0],cols=children(section,'cols')[0],indent=children(children(table,'tblPr')[0],'tblInd')[0]
      if(page&&margin&&Number(attr(cols,'num')||1)===1&&(!indent||attr(indent,'type')==='dxa'))availableWidthTwips=Number(attr(page,'w'))-Number(attr(margin,'left')||0)-Number(attr(margin,'right')||0)-Number(attr(margin,'gutter')||0)-Math.max(0,Number(attr(indent,'w')||0))
    }
    return { availableWidthTwips, index, indexSpace:'document.xml body descendant tables in document order', rows: rows.length, columns: Math.max(0, ...rows.map(row => children(row, 'tc').length)),
      gridWidthsTwips: children(children(table, 'tblGrid')[0], 'gridCol').map(column => Number(attr(column, 'w') || 0)),
      declaredWidth: (() => { const width=children(children(table,'tblPr')[0],'tblW')[0]; return {type:attr(width,'type')||'auto',twips:Number(attr(width,'w')||0)} })(),
      layoutRecommendation: databaseTableRecommendation(table,index,paragraphIndices) || analysisTableRecommendation(table,index,paragraphIndices),
      cellDetails: rows.flatMap((row,rowIndex)=>children(row,'tc').map((cell,cellIndex)=>{
        const pr=children(cell,'tcPr')[0]
        return {rowIndex,cellIndex,verticalAlignment:attr(children(pr,'vAlign')[0],'val')||null,
          gridSpan:Number(attr(children(pr,'gridSpan')[0],'val')||1),verticalMerge:children(pr,'vMerge').length?(attr(children(pr,'vMerge')[0],'val')||'continue'):null}
      })),
      rowDetails: rows.map((row,rowIndex)=>({rowIndex,header:on(children(children(row,'trPr')[0]||row,'tblHeader')[0]),cantSplit:on(children(children(row,'trPr')[0]||row,'cantSplit')[0]),cells:children(row,'tc').length})),
      text: elements(table, 't').map(text => text.textContent || '').join('').slice(0, 500) }
  })
  const topBlocks = Array.from(documentBody.childNodes).filter(node => node.nodeType === 1 && node.namespaceURI === W)
  const captionTargets = { figures: [], tables: [] }
  let tableIndex = 0, paragraphIndex = 0
  for (const block of topBlocks) {
    if (block.localName === 'tbl') captionTargets.tables.push({index:tableIndex++, text:elements(block, 't').map(t=>t.textContent||'').join('').slice(0,300)})
    if (block.localName === 'p') {
      if (block.getElementsByTagNameNS(W, 'drawing').length) captionTargets.figures.push({index:paragraphIndex,
        text:elements(block, 't').map(t=>t.textContent||'').join('').slice(0,300), floating:block.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing','anchor').length>0})
      paragraphIndex++
    }
  }
  const nestedParagraphs = elements(documentBody, 'p').filter(paragraph => {
    let parent = paragraph.parentNode
    while (parent && parent !== documentBody) { if (parent.namespaceURI === W && parent.localName === 'tbl') return true; parent = parent.parentNode }
    return false
  }).map(paragraph => ({ text: elements(paragraph, 't').map(text => text.textContent || '').join('').slice(0, 500),
    style: children(paragraph, 'pPr').flatMap(pPr => children(pPr, 'pStyle'))[0]?.getAttributeNS(W, 'val') || 'Normal' }))
  const count = (namespace, name, matcher) => Object.keys(parts).filter(path => matcher.test(path)).reduce((sum, path) =>
    sum + Array.from(xml(parts[path]).getElementsByTagNameNS(namespace, name)).length, 0)
  const relationshipIds = name => elements(documentBody, name).length
  const pagePartCount = Object.keys(parts).filter(path => /^word\/(header|footer)\d+\.xml$/.test(path)).length
  return { paragraphs, nestedParagraphs, tables, cells, sections, pageNumbers, ...inspectLayoutRisks(mainDocument), partCount: Object.keys(parts).length, hash: digest(bytes),
    elements: { images: Array.from(documentBody.getElementsByTagNameNS(A, 'blip')).length + Array.from(documentBody.getElementsByTagNameNS(V, 'imagedata')).length,
      equations: Array.from(documentBody.getElementsByTagNameNS(M, 'oMath')).length,
      lists: paragraphs.filter(p => p.numbering).length, hyperlinks: elements(documentBody, 'hyperlink').length,
      fields: elements(documentBody, 'fldSimple').length + elements(documentBody, 'instrText').length,
      footnotes: count(W, 'footnote', /^word\/footnotes\.xml$/), endnotes: count(W, 'endnote', /^word\/endnotes\.xml$/),
      headerFooterParts: pagePartCount, headerFooterReferences: relationshipIds('headerReference') + relationshipIds('footerReference') },
    captions, captionTargets, limitations: ['font.name 是西文字体；eastAsiaDirect/eastAsiaThemes 是各文本运行的中文直接字体/主题引用，null 表示未直接指定，需继承样式或默认值，不能据此推断最终中文字体', '标题仅依据 Word 样式和大纲级别识别，不自动猜测无样式标题', '原生目录通过已连接编辑器的update_native_toc执行；页脚支持受限PAGE域设置，复杂页脚仍需单独检查；题注对象范围以实际工具能力为准'] }
}

