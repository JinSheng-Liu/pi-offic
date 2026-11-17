import {formatTocEntries,tabularTextToTable} from './tab-layout.mjs'
import {applyCharacterStyle} from './character-style.mjs'
import {applyNamedStyle} from './named-style.mjs'
import {createFormatReplay} from './format-replay.mjs'
import {ensurePageNumbers} from './ensure-page-numbers.mjs'
import {fitTableWidth} from './table-width.mjs'
import {layoutCover} from './cover-layout.mjs'
import {convertCaptions} from './convert-captions.mjs'
import {readEditorDocument,findTextRanges} from './editor-reader.mjs'
import {formatEditorParagraphs} from './editor-format.mjs'
import { zipSync } from 'fflate'
import { XMLSerializer } from '@xmldom/xmldom'
import { unpack, xml, W, digest, canonical, compareProtected, problem } from './package.mjs'
import { inspect } from './structure.mjs'
import { isBodyCandidate } from './paragraph-role.mjs'
import { validateOperations, paginationKeys } from './rules.mjs'
import { formatFooterPageNumber } from './footer-page-number.mjs'
import { separateTocTitle } from './toc-layout.mjs'
import { flowObjects } from './object-flow.mjs'
import { formatTables } from './table-layout.mjs'
import { applySectionOperations } from './ooxml-sections.mjs'
import { numberEquations, renumberHeadings } from './numbering-operations.mjs'
// Character snapshots permit native run splitting while protecting every unselected
// character's direct formatting. Complex inline structures are deliberately refused.
function rangeSnapshot(paragraph) {
  const chars = []
  const children = node => Array.from(node.childNodes).filter(n => n.nodeType === 1)
  for (const run of children(paragraph)) {
    if (run.namespaceURI === W && run.localName === 'pPr') continue
    if (run.namespaceURI !== W || run.localName !== 'r' || Array.from(run.attributes).some(a => a.namespaceURI !== W || !['rsidR','rsidRPr','rsidDel'].includes(a.localName))) throw problem('该段包含复杂结构，暂不支持局部格式', 422)
    const attributes = JSON.stringify(Array.from(run.attributes).map(a => [a.namespaceURI, a.localName, a.value]).sort())
    const nodes = children(run)
    if (nodes.some(n => n.namespaceURI !== W || !['rPr','t','tab','br','cr'].includes(n.localName))) throw problem('该段包含域或复杂结构，暂不支持局部格式', 422)
    const props = nodes.find(n => n.localName === 'rPr')
    for (const node of nodes.filter(n => n.localName !== 'rPr')) {
      if (node.localName !== 't') {
        const type = node.getAttributeNS(W, 'type')
        const char = node.localName === 'tab' ? '\t' : node.localName === 'cr' ? '\r' : type === 'page' ? '\u000C' : type === 'column' ? '\u000E' : '\u000B'
        chars.push({char, attributes, node:canonical(node), props:props ? canonical(props) : '', protected:props ? canonical(props,true) : ''})
        continue
      }
      if (Array.from(node.attributes).some(a => a.namespaceURI !== 'http://www.w3.org/XML/1998/namespace' || a.localName !== 'space')) throw problem('文字节点含未知属性', 422)
      for (const char of (node.textContent || '').split('')) chars.push({ char, attributes, props: props ? canonical(props) : '', protected: props ? canonical(props, true) : '' })
    }
  }
  return chars
}

// Match the body API's container traversal without entering paragraph drawings/text boxes.
// Block-level alternate content needs editor branch selection and is deliberately refused here.
function editableParagraphNodes(document) {
  const result = []
  const walk = node => {
    if (node.namespaceURI === W && node.localName === 'p') { result.push(node); return }
    if (node.namespaceURI === 'http://schemas.openxmlformats.org/markup-compatibility/2006' && node.localName === 'AlternateContent') throw problem('分支内容暂不支持局部格式定位', 422)
    for (const child of Array.from(node.childNodes)) if (child.nodeType === 1) walk(child)
  }
  walk(document.getElementsByTagNameNS(W, 'body')[0])
  return result
}

async function formatText(bytes, op, signal) {
  signal?.throwIfAborted()
  const body = readEditorDocument(bytes)
  const p = body.paragraphs[op.indices[0]]
  if (!p) throw problem('段落索引超出范围')
  const matches = findTextRanges(p,op.text)
  if ((op.occurrence === undefined && matches.length !== 1) || !matches[op.occurrence ?? 0]) throw problem('文字范围未唯一匹配，请提供准确原文或 occurrence', 422)
  const range = matches[op.occurrence ?? 0]
  const parts = unpack(bytes), before = xml(parts['word/document.xml'])
  const paragraphs = editableParagraphNodes(before)
  const source = paragraphs[op.indices[0]]
  if (!source || paragraphs.length !== body.paragraphs.length) throw problem('复杂文档段落定位暂不支持局部格式', 422)
  const snapshot = rangeSnapshot(source)
  if (snapshot.map(c => c.char).join('') !== p.text) throw problem('文字范围与文档字符位置不一致', 422)
  const font = { ...op.font, ...(op.font.color ? { color: op.font.color.startsWith('#') ? op.font.color : `#${op.font.color}` } : {}) }
  let output = formatEditorParagraphs(bytes,[{indices:op.indices,op:{font,range:{start:range.start,end:range.end}}}],signal)
  const savedParts = unpack(output), after = xml(savedParts['word/document.xml'])
  const target = editableParagraphNodes(after)[op.indices[0]]
  // The editor may split runs without copying rsid metadata. Restore it per original
  // character span; refuse a merge across different metadata instead of guessing.
  let offset = 0
  for (const run of Array.from(target.childNodes).filter(n => n.namespaceURI === W && n.localName === 'r')) {
    const length = Array.from(run.childNodes).reduce((sum, n) => sum + (n.namespaceURI !== W ? 0 : n.localName === 't' ? (n.textContent || '').length : ['tab','br','cr'].includes(n.localName) ? 1 : 0), 0)
    const span = snapshot.slice(offset, offset + length)
    if (span.some(c => c.attributes !== span[0].attributes)) throw problem('局部格式合并了不同修订属性的文字', 422)
    if (span.length) {
      for (const name of ['rsidR','rsidRPr','rsidDel']) run.removeAttributeNS(W, name)
      for (const [ns, name, value] of JSON.parse(span[0].attributes)) run.setAttributeNS(ns, `w:${name}`, value)
    }
    offset += length
  }
  savedParts['word/document.xml'] = new TextEncoder().encode(new XMLSerializer().serializeToString(after))
  output = zipSync(savedParts)
  const saved = rangeSnapshot(target)
  if (saved.length !== snapshot.length || saved.some((c,i) => c.char !== snapshot[i].char || c.attributes !== snapshot[i].attributes || c.node !== snapshot[i].node || c.protected !== snapshot[i].protected || ((i < range.start || i >= range.end) && c.props !== snapshot[i].props))) throw problem('局部格式影响了非目标文字或属性', 422)
  const sourceProps = Array.from(source.childNodes).find(n => n.localName === 'pPr')
  const targetProps = Array.from(target.childNodes).find(n => n.localName === 'pPr')
  if ((sourceProps ? canonical(sourceProps) : '') !== (targetProps ? canonical(targetProps) : '')) throw problem('局部格式改变了段落属性', 422)
  target.parentNode.replaceChild(after.importNode(source, true), target)
  if (canonical(before.documentElement) !== canonical(after.documentElement)) throw problem('局部格式改变了非目标段落', 422)
  savedParts['word/document.xml'] = new TextEncoder().encode(new XMLSerializer().serializeToString(after))
  const validation = compareProtected(bytes, zipSync(savedParts))
  const reopened = readEditorDocument(output)
  const verified = findTextRanges(reopened.paragraphs[op.indices[0]],op.text)[op.occurrence ?? 0]
  for (const [key,value] of Object.entries(font)) if (verified?.font[key === 'latinName' ? 'name' : key] !== (key === 'name' ? font.latinName || value : value)) throw problem('局部字体重读验证失败', 422)
  if (font.name && verified.runs.some(r => r.value.rPr?.rFonts?.eastAsia !== font.name || r.value.rPr?.rFonts?.eastAsiaTheme)) throw problem('局部中文字体重读验证失败', 422)
  if (font.latinName) verifyLatinFont(verified.runs.map(r => r.value), font.latinName)
  signal?.throwIfAborted()
  validation.checks.push('局部文字格式重读验证', '未选中文字和格式保持')
  return {bytes:output, affected:1, validation}
}

function verifyLatinFont(holders, name) {
  for (const holder of holders) {
    const fonts = holder?.rPr?.rFonts
    if (fonts?.ascii !== name || fonts.hAnsi !== name || fonts.asciiTheme || fonts.hAnsiTheme) throw problem('西文字体保存验证失败', 422)
  }
}

export async function format(bytes, operations, signal) {
  validateOperations(operations)
  if(operations[0].scope==='tocEntries')return formatTocEntries(bytes,signal,formatParagraphs)
  if (operations[0].scope === 'tabularText') {
    const converted = tabularTextToTable(bytes, operations[0], signal)
    const data = await inspect(converted.bytes)
    const recommendation = data.tables[converted.tableIndex]?.layoutRecommendation
    if (!recommendation) return converted
    let result = converted
    // Structural conversion and its verified layout are committed together.
    for (let start = 0; start < recommendation.operations.length; start += 30) {
      result = await format(result.bytes, recommendation.operations.slice(start, start + 30), signal)
    }
    if (recommendation.widthOperation) result = await format(result.bytes, [recommendation.widthOperation], signal)
    const verified = await inspect(result.bytes)
    if (verified.tables[converted.tableIndex]?.layoutRecommendation?.kind !== recommendation.kind) throw problem('转换后的表格分类验证失败', 422)
    return { ...result, affected: converted.affected, tableIndex: converted.tableIndex,
      validation: { ...result.validation, checks: [...converted.validation.checks, ...result.validation.checks, '表格转换及分类排版原子提交'],
        limitations: result.validation.limitations } }
  }
  if(operations[0].scope==='namedStyle')return applyNamedStyle(bytes,operations[0],signal,formatParagraphs,editableParagraphNodes)
  if(operations[0].scope==='ensurePageNumbers')return ensurePageNumbers(bytes,signal)
  if(operations[0].scope==='fitTableWidth')return fitTableWidth(bytes,operations[0],signal)
  if(operations[0].scope==='coverLayout')return layoutCover(bytes,operations[0],signal)
  if (operations[0].scope==='convertCaptions') return convertCaptions(bytes,operations[0],signal)
  if (operations[0].scope==='footerPageNumber') return formatFooterPageNumber(bytes,operations[0],signal)
  if (operations[0].scope==='tocTitle') return separateTocTitle(bytes,operations[0],signal)
  if (operations[0].scope==='floatingObjects') return flowObjects(bytes,operations[0],signal)
  if (operations[0].scope === 'renumberHeadings') return renumberHeadings(bytes, signal)
  if (operations[0].scope === 'numberEquations') return numberEquations(bytes, operations[0], signal)
  if (!operations.some(op => ['characterStyle','text','tableRows','tableCells'].includes(op.scope))) return formatParagraphs(bytes, operations, signal)
  let result = {bytes, affected:0}, checks = new Set(), limitations = new Set()
  // Preserve operation order; the store commits only once after every stage succeeds.
  for (const op of operations) {
    const next = op.scope==='characterStyle' ? await applyCharacterStyle(result.bytes,op,signal,formatText,editableParagraphNodes) : ['tableRows','tableCells'].includes(op.scope) ? formatTables(result.bytes, op, signal) : op.scope === 'text' ? await formatText(result.bytes, op, signal) : await formatParagraphs(result.bytes, [op], signal)
    result = {bytes:next.bytes, affected:result.affected + next.affected}
    next.validation.checks.forEach(c => checks.add(c)); next.validation.limitations.forEach(c => limitations.add(c))
  }
  return {...result, validation:{status:'passed',checks:[...checks],limitations:[...limitations]}}
}

function ensureRequestedStyles(bytes, operations) {
  const requested = [...new Set(operations.map(op => op.style).filter(Boolean))]
  if (!requested.length) return bytes
  const parts = unpack(bytes)
  if (!parts['word/styles.xml']) throw problem('缺少样式定义部分，无法安全应用标题样式', 422)
  const document = xml(parts['word/styles.xml'])
  const styles = Array.from(document.getElementsByTagNameNS(W, 'style'))
  const missing = requested.filter(id => !styles.some(s => s.getAttributeNS(W, 'styleId') === id))
  if (!missing.length) return bytes
  const defaultStyle = styles.find(s => s.getAttributeNS(W, 'type') === 'paragraph' && ['1','true','on'].includes(s.getAttributeNS(W, 'default')))?.getAttributeNS(W, 'styleId')
  for (const id of missing) {
    const style = document.createElementNS(W, 'w:style')
    style.setAttributeNS(W, 'w:type', 'paragraph'); style.setAttributeNS(W, 'w:styleId', id)
    const add = (parent, key, value) => { const node = document.createElementNS(W, `w:${key}`); if (value !== undefined) node.setAttributeNS(W, 'w:val', String(value)); parent.appendChild(node); return node }
    add(style, 'name', id.replace(/Heading(\d)/, 'heading $1'))
    if (defaultStyle && defaultStyle !== id) add(style, 'basedOn', defaultStyle)
    if (/^Heading[1-9]$/.test(id)) {
      const pPr = add(style, 'pPr')
      add(pPr, 'keepNext'); add(pPr, 'keepLines'); add(pPr, 'outlineLvl', Number(id.slice(-1))-1)
    }
    document.documentElement.appendChild(style)
  }
  parts['word/styles.xml'] = new TextEncoder().encode(new XMLSerializer().serializeToString(document))
  return zipSync(parts)
}

async function formatParagraphs(bytes, operations, signal) {
  validateOperations(operations)
  const originalBytes = bytes
  bytes = ensureRequestedStyles(bytes, operations)
  operations = operations.map(op => ({
    ...op,
    ...(op.font ? { font: { ...op.font, ...(op.font.color ? { color: op.font.color.startsWith('#') ? op.font.color : `#${op.font.color}` } : {}) } } : {}),
    ...(op.paragraph ? { paragraph: { ...op.paragraph, ...(op.paragraph.alignment ? { alignment: ({ Center: 'Centered', Justify: 'Justified' })[op.paragraph.alignment] || op.paragraph.alignment } : {}) } } : {}),
  })).map(op => /^Heading[1-9]$/.test(op.style || '')
    ? {...op, paragraph:{outlineLevel:Number(op.style.slice(-1)), ...op.paragraph}} : op)
  unpack(bytes)
  const body = readEditorDocument(bytes)
  const text = body.text
  const navigationStyles = new Map()
  const targets = []
  const tabTargets = new Set()
  const originalTabNodes = operations.some(op=>op.paragraph?.tocPageTabPosition!==undefined||op.paragraph?.leadingTabPosition!==undefined) ? editableParagraphNodes(xml(unpack(bytes)['word/document.xml'])) : []
  const sectionOperations = operations.filter(op => op.scope === 'sections')
  for (const op of operations) {
    signal?.throwIfAborted()
    if (op.scope === 'sections') continue
    if (op.indices?.some(i => i >= body.paragraphs.length)) throw problem('段落索引超出范围')
    const indices = body.paragraphs.flatMap((p, i) => (op.scope === 'all' ||
      (op.scope === 'body' && isBodyCandidate(p)) ||
      (op.scope === 'tableText' && p.tableNestingLevel > 0) ||
      (op.scope === 'inlinePictures' && p.inlinePictures.length > 0 && !p.text.trim()) ||
      (op.scope === 'headings' && p.outlineLevel <= 9 && (!op.heading || p.outlineLevel === op.heading)) ||
      (op.scope === 'indices' && op.indices.includes(i))) ? [i] : [])
    if (!indices.length) throw problem('操作未匹配任何段落；请明确标题或段落范围', 422)
    for (const i of indices) {
      const p = body.paragraphs[i]
      if (op.style) {
        p.styleId = op.style
        // Direct body outline levels in imported files override style inheritance.
        if (/^Heading[1-9]$/.test(op.style) && op.paragraph?.outlineLevel === undefined) p.outlineLevel = Number(op.style.slice(-1))
      }
      for (const [key, value] of Object.entries(op.paragraph || {})) {
        if (key === 'tocPageTabPosition') {
          if (!/^[^\t]+\t[0-9IVXLCDMivxlcdm]+\s*$/.test(p.text) || p.tableNestingLevel || p.inlinePictures.length || (p.effectivePPr?.tabs?.tab||[]).some(t=>t.val!=='right'||t.leader!=='dot') || value<=0) throw problem('仅支持单Tab加页码、无制表位或已有右对齐点线制表位的目录条目',422)
          p.p.pPr ||= {}; p.p.pPr.tabs={tab:[{val:'right',leader:'dot',pos:Math.round(value*20)}]};tabTargets.add(i)
        }
        else if (key === 'leadingTabPosition') {
          if (!/^\t[^\t]/.test(p.text) || p.text.slice(1).includes('\t') || p.tableNestingLevel || p.inlinePictures.length || p.p.pPr?.numPr) throw problem('仅支持普通段落的单个前导Tab', 422)
          if (p.effectivePPr?.tabs?.tab?.length || p.leftIndent !== 0 || p.firstLineIndent !== 0 || value <= 0) throw problem('存在制表位或缩进，不能自动覆盖；请先检查', 422)
          p.p.pPr ||= {}; p.p.pPr.tabs = {tab:[{val:'left',pos:Math.round(value*20)}]}; tabTargets.add(i)
        }
        else if (key === 'outlineLevel') p[key] = value
      }
      if (op.paragraph?.outlineLevel <= 9 && !op.style) {
        const base = p.styleId
        const level = op.paragraph.outlineLevel
        const styleId = `PiNav${level}_${digest(new TextEncoder().encode(base)).slice(0, 12)}`
        navigationStyles.set(styleId, { base, level })
        p.styleId = styleId
      }
    }
    targets.push({ indices, op })
  }
  const hasSupplementalWrites = targets.some(({op}) => op.style || ['outlineLevel','leadingTabPosition','tocPageTabPosition'].some(key => op.paragraph?.[key] !== undefined))
  let output = bytes
  if (hasSupplementalWrites) {
    // The editor's public property vocabulary does not expose outlineLvl or nested
    // tab definitions. Supplement only these selected properties; never reserialize
    // the document through the old SDK.
    const parts=unpack(output),document=xml(parts['word/document.xml'])
    const nodes=editableParagraphNodes(document)
    for(const {indices,op} of targets)for(const index of indices) {
      const node=nodes[index],p=body.paragraphs[index]
      let pr=Array.from(node.childNodes).find(n=>n.namespaceURI===W&&n.localName==='pPr')
      if(!pr){pr=document.createElementNS(W,'w:pPr');node.insertBefore(pr,node.firstChild)}
      const set=(name,value)=>{
        let el=Array.from(pr.childNodes).find(n=>n.namespaceURI===W&&n.localName===name)
        if(!el){el=document.createElementNS(W,`w:${name}`);pr.appendChild(el)}
        el.setAttributeNS(W,'w:val',String(value));return el
      }
      if(op.style||op.paragraph?.outlineLevel<=9)set('pStyle',p.styleId)
      if(op.paragraph?.outlineLevel!==undefined)set('outlineLvl',p.outlineLevel-1)
      if(op.paragraph?.leadingTabPosition!==undefined||op.paragraph?.tocPageTabPosition!==undefined) {
        for(const old of Array.from(pr.childNodes).filter(n=>n.namespaceURI===W&&n.localName==='tabs'))pr.removeChild(old)
        const tabs=document.createElementNS(W,'w:tabs'),tab=document.createElementNS(W,'w:tab')
        for(const [key,value] of Object.entries(p.p.pPr.tabs.tab[0]))tab.setAttributeNS(W,`w:${key}`,String(value))
        tabs.appendChild(tab);pr.appendChild(tabs)
      }
    }
    parts['word/document.xml']=new TextEncoder().encode(new XMLSerializer().serializeToString(document))
    output=zipSync(parts)
  }
  if (targets.length) output = formatEditorParagraphs(output, targets, signal)
  if (navigationStyles.size) {
    const parts = unpack(output)
    const styles = xml(parts['word/styles.xml'])
    for (const [id, { base, level }] of navigationStyles) {
      if (Array.from(styles.getElementsByTagNameNS(W, 'style')).some(s => s.getAttributeNS(W, 'styleId') === id)) continue
      const style = styles.createElementNS(W, 'w:style')
      style.setAttributeNS(W, 'w:type', 'paragraph'); style.setAttributeNS(W, 'w:styleId', id)
      const add = (parent, name, value) => { const element = styles.createElementNS(W, `w:${name}`); if (value !== undefined) element.setAttributeNS(W, 'w:val', String(value)); parent.appendChild(element); return element }
      add(style, 'name', `导航标题 ${level}（保留原格式）`)
      add(style, 'basedOn', base)
      add(add(style, 'pPr'), 'outlineLvl', level - 1)
      styles.documentElement.appendChild(style)
    }
    parts['word/styles.xml'] = new TextEncoder().encode(new XMLSerializer().serializeToString(styles))
    output = zipSync(parts)
  }
  // Keep global tab protection: remove only newly added, verified target tabs
  // from a temporary comparison copy. The saved document retains the new stops.
  let comparisonOutput = output
  if (tabTargets.size) {
    const parts = unpack(output), document = xml(parts['word/document.xml'])
    const nodes = editableParagraphNodes(document)
    for (const index of tabTargets) {
      const pPr = Array.from(nodes[index].childNodes).find(n=>n.namespaceURI===W && n.localName==='pPr')
      const tabs = Array.from(pPr?.childNodes || []).filter(n=>n.namespaceURI===W && n.localName==='tabs')
      if (tabs.length !== 1) throw problem('目标制表位结构验证失败',422)
      const originalPr=Array.from(originalTabNodes[index].childNodes).find(n=>n.namespaceURI===W&&n.localName==='pPr')
      const originalTabs=Array.from(originalPr?.childNodes||[]).find(n=>n.namespaceURI===W&&n.localName==='tabs')
      if(originalTabs)pPr.replaceChild(document.importNode(originalTabs,true),tabs[0]);else pPr.removeChild(tabs[0])
    }
    parts['word/document.xml'] = new TextEncoder().encode(new XMLSerializer().serializeToString(document))
    comparisonOutput = zipSync(parts)
  }
  const validation = compareProtected(originalBytes, comparisonOutput)
  let sectionAffected = 0
  if (sectionOperations.length) {
    const result = applySectionOperations(output, sectionOperations)
    output = result.bytes
    sectionAffected = result.affected
    const verified = await inspect(output)
    const expectedSections = new Map()
    for (const op of sectionOperations) for (const index of op.sectionIndices) {
      const previous = expectedSections.get(index) || { page: {}, margins: {} }
      expectedSections.set(index, { page: { ...previous.page, ...op.page }, margins: { ...previous.margins, ...op.margins } })
    }
    for (const [index, expected] of expectedSections) {
      for (const kind of ['page', 'margins']) for (const [key, value] of Object.entries(expected[kind])) {
        if (verified.sections[index]?.[kind]?.[key] !== value) throw problem(`分节 ${index + 1} 的 ${key} 重读验证失败`, 422)
      }
    }
    validation.checks.push('目标分节页面属性重读验证')
  }
  signal?.throwIfAborted()
  const result = readEditorDocument(output)
  if (result.text !== text || result.paragraphs.length !== body.paragraphs.length) throw problem('正文内容验证失败', 422)
  // Later overlapping operations intentionally win.
  const expected = new Map()
  for (const { indices, op } of targets) for (const i of indices) {
    const prev = expected.get(i) || { font: {}, paragraph: {} }
    if (op.font?.name && !op.font.latinName) delete prev.font.latinName
    if (op.paragraph?.lineSpacingMultiple !== undefined) delete prev.paragraph.lineSpacing
    if (op.paragraph?.lineSpacing !== undefined) delete prev.paragraph.lineSpacingMultiple
    expected.set(i, { style: op.style ?? prev.style, font: { ...prev.font, ...op.font }, paragraph: { ...prev.paragraph, ...op.paragraph } })
  }
  for (const [i, e] of expected) {
    const p = result.paragraphs[i]
    if (e.font.latinName) verifyLatinFont([...p.runs.map(r => r.value), p.p.pPr], e.font.latinName)
    if (e.style && p.styleId !== e.style) throw problem('标题样式验证失败', 422)
    if (/^Heading[1-9]$/.test(e.style || '') && e.paragraph.outlineLevel === undefined && p.outlineLevel !== Number(e.style.slice(-1))) throw problem('标题大纲层级验证失败', 422)
    if (e.font.name) {
      const holders = [...p.runs.map(run => run.value), p.p.pPr]
      for (const holder of holders) {
        const fonts = holder?.rPr?.rFonts
        if (fonts?.eastAsia !== e.font.name || fonts.eastAsiaTheme) throw problem('中文字体保存验证失败', 422)
      }
    }
    for (const [key, val] of Object.entries(e.font)) if (p.font[key === 'latinName' ? 'name' : key] !== (key === 'name' ? e.font.latinName || val : val)) throw problem(`字体属性 ${key} 保存验证失败`, 422)
    for (const [key, val] of Object.entries(e.paragraph)) if ((['leadingTabPosition','tocPageTabPosition'].includes(key) ? (p.p.pPr?.tabs?.tab?.length === 1 && p.p.pPr.tabs.tab[0].val === (key==='leadingTabPosition'?'left':'right') ? p.p.pPr.tabs.tab[0].pos / 20 : undefined) : key === 'lineSpacingMultiple' ? (p.p.pPr?.spacing?.lineRule === 'auto' ? p.p.pPr.spacing.line / 240 : undefined) : paginationKeys.has(key) ? p.p.pPr?.[key]?.val : p[key]) !== (['leadingTabPosition','tocPageTabPosition'].includes(key) ? Math.round(val*20)/20 : key === 'lineSpacingMultiple' ? Math.round(val*240)/240 : val)) throw problem(`段落属性 ${key} 保存验证失败`, 422)
  }
  validation.checks.push('目标格式重读验证', '正文文本一致')
  return { bytes: output, validation, affected: expected.size + sectionAffected,
    ...(!hasSupplementalWrites&&!sectionOperations.length?{replay:createFormatReplay(originalBytes,output,targets)}:{}) }
}
