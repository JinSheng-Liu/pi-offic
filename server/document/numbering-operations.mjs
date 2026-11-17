import { zipSync } from 'fflate'
import { XMLSerializer } from '@xmldom/xmldom'
import { unpack, xml, W, M, canonical, problem } from './package.mjs'

const children = node => Array.from(node.childNodes).filter(child => child.nodeType === 1)
const direct = (node, namespace, name) => node ? children(node).find(child => child.namespaceURI === namespace && child.localName === name) : undefined
const bodyBlocks = document => children(document.getElementsByTagNameNS(W, 'body')[0])
  .filter(node => node.namespaceURI === W && ['p', 'tbl'].includes(node.localName))
const paragraphText = paragraph => Array.from(paragraph.getElementsByTagNameNS(W, 't')).map(node => node.textContent || '').join('')
const paragraphLevel = paragraph => {
  const pPr = direct(paragraph, W, 'pPr')
  const outline = direct(pPr, W, 'outlineLvl')?.getAttributeNS(W, 'val')
  if (outline !== undefined && outline !== null && Number(outline) <= 8) return Number(outline) + 1
  const style = direct(pPr, W, 'pStyle')?.getAttributeNS(W, 'val') || ''
  const match = /^Heading([1-9])$/i.exec(style)
  return match ? Number(match[1]) : 0
}

function replaceTextPrefix(paragraph, oldPrefix, newPrefix) {
  const runs = children(paragraph).filter(node => node.namespaceURI === W && node.localName === 'r')
  const textNodes = runs.flatMap(run => children(run).filter(node => node.namespaceURI === W && node.localName === 't'))
  const original = textNodes.map(node => node.textContent || '')
  let remaining = oldPrefix.length
  if (!textNodes.length || original.join('').slice(0, remaining) !== oldPrefix) throw problem('标题编号无法精确定位，未更改标题', 422)
  const updated = original.map((value, index) => {
    const consumed = Math.min(remaining, value.length)
    remaining -= consumed
    return (index === 0 ? newPrefix : '') + value.slice(consumed)
  })
  textNodes.forEach((node, index) => {
    node.textContent = updated[index]
    if (/^\s|\s$/u.test(updated[index])) node.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve')
  })
}

function headingPrefix(text) {
  const match = /^(\s*)(第\s*\d+\s*章|\d+(?:\.\d+){0,8})([.)、]?)(\s*)/u.exec(text)
  if (!match) return null
  return { full: match[0], leading: match[1], number: match[2], punctuation: match[3], trailing: match[4], chapter: match[2].startsWith('第') }
}

function expectedHeadingNumber(level, counters) {
  counters[level - 1] = (counters[level - 1] || 0) + 1
  counters.length = level
  for (let index = 0; index < level - 1; index++) counters[index] ||= 1
  return counters.slice(0, level).join('.')
}

function snapshotTargets(originalDocument, changedDocument, targets) {
  const encoder = new TextEncoder()
  const original = xml(encoder.encode(new XMLSerializer().serializeToString(originalDocument)))
  const restored = xml(encoder.encode(new XMLSerializer().serializeToString(changedDocument)))
  const targetSet = new Set(targets)
  const originals = bodyBlocks(original)
  const changed = bodyBlocks(restored)
  for (const index of targetSet) changed[index].parentNode.replaceChild(restored.importNode(originals[index], true), changed[index])
  if (canonical(original.documentElement) !== canonical(restored.documentElement)) throw problem('操作影响了目标以外的文档结构', 422)
}

function serializeChecked(bytes, parts, document, originalDocument, targetIndices, checks, affected) {
  snapshotTargets(originalDocument, document, targetIndices)
  const output = new TextEncoder().encode(new XMLSerializer().serializeToString(document))
  const reopened = xml(output)
  const reopenedBlocks = bodyBlocks(reopened)
  for (const index of targetIndices) {
    if (!reopenedBlocks[index]) throw problem('编号结果回读失败', 422)
  }
  const allParts = { ...parts, 'word/document.xml': output }
  for (const [name, data] of Object.entries(parts)) if (name !== 'word/document.xml' && allParts[name] !== data) throw problem('操作改动了非目标 DOCX 部件', 422)
  return { bytes: zipSync(allParts), affected, validation: { status: 'passed', checks, limitations: ['已验证 OOXML 编号域和目标范围；最终分页、域显示与视觉位置需在 Word/WPS 或编辑器中检查'] } }
}

export function renumberHeadings(bytes, signal) {
  signal?.throwIfAborted()
  const parts = unpack(bytes)
  const originalDocument = xml(parts['word/document.xml'])
  const document = xml(parts['word/document.xml'])
  const blocks = bodyBlocks(document)
  const counters = []
  const targets = []
  let changed = 0
  for (let index = 0; index < blocks.length; index++) {
    signal?.throwIfAborted()
    const paragraph = blocks[index]
    if (paragraph.localName !== 'p') continue
    const level = paragraphLevel(paragraph)
    if (!level) continue
    const prefix = headingPrefix(paragraphText(paragraph))
    if (!prefix) continue
    if (direct(direct(paragraph, W, 'pPr'), W, 'numPr')) throw problem('检测到使用 Word 原生列表编号的标题；需使用编辑器列表编号能力调整，未混合重排', 422)
    const number = prefix.chapter ? String((counters[0] || 0) + 1) : expectedHeadingNumber(level, counters)
    if (prefix.chapter) { counters[0] = Number(number); counters.length = 1 }
    const next = `${prefix.leading}${prefix.chapter ? `第${number}章` : number}${prefix.punctuation}${prefix.trailing}`
    if (next === prefix.full) continue
    const allRuns = children(paragraph).filter(node => node.namespaceURI === W && node.localName === 'r')
    if (allRuns.length !== children(paragraph).filter(node => node.namespaceURI === W && node.localName !== 'pPr').length ||
        allRuns.some(run => children(run).some(node => !['rPr', 't'].includes(node.localName) || node.namespaceURI !== W)))
      throw problem('标题包含域、书签、超链接或复杂内联对象；为保护内容已停止编号重排', 422)
    replaceTextPrefix(paragraph, prefix.full, next)
    targets.push(index)
    changed++
  }
  if (!changed) throw problem('未发现需要校正且可确认的手工编号标题', 422)
  return serializeChecked(bytes, parts, document, originalDocument, targets, ['标题层级识别与手工编号连续性校验', '非标题正文和标题后续文字保持'], changed)
}

const numberFormats = {
  arabic: { instruction: 'ARABIC', render: value => String(value) },
  upperRoman: { instruction: 'ROMAN', render: value => roman(value) },
  lowerRoman: { instruction: 'roman', render: value => roman(value).toLowerCase() },
  upperLetter: { instruction: 'ALPHABETIC', render: value => letters(value) },
  lowerLetter: { instruction: 'alphabetic', render: value => letters(value).toLowerCase() },
}

function roman(value) {
  if (value > 3999) return String(value)
  const pairs = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']]
  let output = ''
  for (const [number, symbol] of pairs) while (value >= number) { output += symbol; value -= number }
  return output
}

function letters(value) {
  let output = ''
  while (value > 0) { value--; output = String.fromCharCode(65 + value % 26) + output; value = Math.floor(value / 26) }
  return output || '0'
}

function wordRun(document, name, value) {
  const run = document.createElementNS(W, 'w:r')
  const node = document.createElementNS(W, `w:${name}`)
  if (name === 'instrText') node.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve')
  if (name !== 'tab') node.appendChild(document.createTextNode(value))
  run.appendChild(node)
  return run
}

function fieldRuns(document, index, instruction, prefix, suffix, result) {
  const fieldChar = type => {
    const run = document.createElementNS(W, 'w:r')
    const node = document.createElementNS(W, 'w:fldChar')
    node.setAttributeNS(W, 'w:fldCharType', type)
    run.appendChild(node)
    return run
  }
  return [
    wordRun(document, 't', prefix), fieldChar('begin'),
    wordRun(document, 'instrText', ` SEQ PiEquation${instruction ? ` \\* ${instruction}` : ''}${index === 1 ? ' \\r 1' : ''} `),
    fieldChar('separate'), wordRun(document, 't', result), fieldChar('end'), wordRun(document, 't', suffix),
  ]
}

function sectionWidthTwips(paragraph, body) {
  let section = direct(direct(paragraph, W, 'pPr'), W, 'sectPr')
  if (!section) {
    const blocks = bodyBlocks(body.ownerDocument)
    const position = blocks.indexOf(paragraph)
    for (let index = position + 1; index < blocks.length && !section; index++) {
      section = direct(direct(blocks[index], W, 'pPr'), W, 'sectPr')
    }
    section ||= direct(body, W, 'sectPr')
  }
  const page = direct(section, W, 'pgSz')
  const margins = direct(section, W, 'pgMar')
  if (!page || !margins) return null
  const width = Number(page.getAttributeNS(W, 'w'))
  const left = Number(margins.getAttributeNS(W, 'left'))
  const right = Number(margins.getAttributeNS(W, 'right'))
  const gutter = Number(margins.getAttributeNS(W, 'gutter') || 0)
  const usable = width - left - right - gutter
  return Number.isInteger(usable) && usable > 0 ? usable : null
}

export function numberEquations(bytes, options = {}, signal) {
  const { format = 'arabic', prefix = '(', suffix = ')' } = options
  if (!numberFormats[format] || typeof prefix !== 'string' || typeof suffix !== 'string' || prefix.length > 20 || suffix.length > 20 || /[\r\n\u0000-\u001f]/u.test(prefix + suffix)) throw problem('公式编号格式无效')
  signal?.throwIfAborted()
  const parts = unpack(bytes)
  const originalDocument = xml(parts['word/document.xml'])
  const document = xml(parts['word/document.xml'])
  const body = document.getElementsByTagNameNS(W, 'body')[0]
  const blocks = bodyBlocks(document)
  const formulas = []
  for (let index = 0; index < blocks.length; index++) {
    const paragraph = blocks[index]
    if (paragraph.localName !== 'p') continue
    const equations = Array.from(paragraph.getElementsByTagNameNS(M, 'oMathPara'))
    if (!equations.length) continue
    if (equations.length !== 1 || equations[0].parentNode !== paragraph) throw problem('仅支持每段一个独立显示公式；复杂嵌套公式未修改', 422)
    const content = children(paragraph).filter(node => !(node.namespaceURI === W && node.localName === 'pPr'))
    if (content.length !== 1 || content[0] !== equations[0]) throw problem('公式段含有混排文字、手工编号或其他内联对象；未添加重复编号', 422)
    if (paragraph.getElementsByTagNameNS(W, 'instrText').length || paragraph.getElementsByTagNameNS(W, 'fldSimple').length) throw problem('公式段已有域，无法确认现有编号状态', 422)
    const width = sectionWidthTwips(paragraph, body)
    if (!width) throw problem('无法可靠读取公式所在分节的页宽和边距', 422)
    formulas.push({ paragraph, index, width })
  }
  if (!formulas.length) throw problem('未找到独立显示公式段落；行内公式不会自动编号', 422)
  if (Array.from(document.getElementsByTagNameNS(W, 'instrText')).some(node => /\bSEQ\s+PiEquation\b/i.test(node.textContent || '')) ||
      Array.from(document.getElementsByTagNameNS(W, 'fldSimple')).some(node => /\bSEQ\s+PiEquation\b/i.test(node.getAttributeNS(W, 'instr') || '')))
    throw problem('文档已有 PiEquation 序列域；请先更新或核查现有公式编号', 422)
  for (let index = 0; index < formulas.length; index++) {
    signal?.throwIfAborted()
    const { paragraph, width } = formulas[index]
    const pPr = direct(paragraph, W, 'pPr') || (() => { const node = document.createElementNS(W, 'w:pPr'); paragraph.insertBefore(node, paragraph.firstChild); return node })()
    const tabs = direct(pPr, W, 'tabs') || document.createElementNS(W, 'w:tabs')
    if (!tabs.parentNode) {
      const order = ['pStyle','keepNext','keepLines','pageBreakBefore','framePr','widowControl','numPr','suppressLineNumbers','pBdr','shd','tabs']
      const next = children(pPr).find(node => !order.includes(node.localName))
      pPr.insertBefore(tabs, next || null)
    }
    const tab = document.createElementNS(W, 'w:tab')
    tab.setAttributeNS(W, 'w:val', 'right')
    tab.setAttributeNS(W, 'w:pos', String(width))
    tabs.appendChild(tab)
    paragraph.appendChild(wordRun(document, 'tab', ''))
    const number = index + 1
    for (const run of fieldRuns(document, number, numberFormats[format].instruction, prefix, suffix, numberFormats[format].render(number))) {
      paragraph.appendChild(run)
    }
  }
  const targetIndices = formulas.map(formula => formula.index)
  return serializeChecked(bytes, parts, document, originalDocument, targetIndices, ['独立显示公式识别', 'Word SEQ 公式编号域与右对齐制表位回读', '公式本体及非目标内容保持'], formulas.length)
}
