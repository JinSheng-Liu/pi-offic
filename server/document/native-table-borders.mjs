import { randomBytes } from 'node:crypto'
import { zipSync } from 'fflate'
import { XMLSerializer } from '@xmldom/xmldom'
import { unpack, xml, W, problem } from './package.mjs'

const elements = node => Array.from(node?.childNodes || []).filter(n => n.nodeType === 1)
const encode = doc => new TextEncoder().encode(new XMLSerializer().serializeToString(doc))
const NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

export function prepareThreeLineTable(bytes, tableIndex) {
  const parts = unpack(bytes)
  const doc = xml(parts['word/document.xml'])
  const body = doc.getElementsByTagNameNS(W, 'body')[0]
  const tables = elements(body).filter(n => n.namespaceURI === W && n.localName === 'tbl')
  const table = tables[tableIndex]
  if (!table) throw problem('三线表目标不存在或不是正文顶层表格', 422)
  const paragraph = table.getElementsByTagNameNS(W, 'p')[0]
  if (!paragraph) throw problem('三线表缺少可定位的首个单元格段落', 422)
  let paraId = paragraph.getAttributeNS('http://schemas.microsoft.com/office/word/2010/wordml', 'paraId')
  if (!paraId) {
    paraId = randomBytes(4).toString('hex').toUpperCase()
    paragraph.setAttributeNS('http://schemas.microsoft.com/office/word/2010/wordml', 'w14:paraId', paraId)
    parts['word/document.xml'] = encode(doc)
  }
  return { bytes: zipSync(parts), paraId }
}

function child(parent, name) {
  return elements(parent).find(node => node.namespaceURI === W && node.localName === name)
}

function ensure(parent, name) {
  let node = child(parent, name)
  if (!node) {
    node = parent.ownerDocument.createElementNS(NS, `w:${name}`)
    parent.appendChild(node)
  }
  return node
}

function setBorder(parent, name, val) {
  const node = ensure(parent, name)
  node.setAttributeNS(W, 'w:val', val)
  if (val === 'single') {
    node.setAttributeNS(W, 'w:sz', '8')
    node.setAttributeNS(W, 'w:color', '000000')
  } else {
    node.removeAttributeNS(W, 'sz')
    node.removeAttributeNS(W, 'color')
  }
}

/**
 * Structural fallback for editor builds that cannot select a table by anchor.
 * It is deliberately limited to top-level tables and is validated after writing.
 */
export function applyThreeLineTableOOXML(bytes, tableIndex) {
  const parts = unpack(bytes)
  const doc = xml(parts['word/document.xml'])
  const body = doc.getElementsByTagNameNS(W, 'body')[0]
  const tables = elements(body).filter(node => node.namespaceURI === W && node.localName === 'tbl')
  const table = tables[tableIndex]
  if (!table) throw problem('三线表目标不存在或不是正文顶层表格', 422)
  const props = child(table, 'tblPr') || table.insertBefore(doc.createElementNS(NS, 'w:tblPr'), table.firstChild)
  const borders = ensure(props, 'tblBorders')
  for (const name of ['left', 'right', 'insideV', 'insideH']) setBorder(borders, name, 'nil')
  setBorder(borders, 'top', 'single')
  setBorder(borders, 'bottom', 'single')
  const rows = elements(table).filter(node => node.namespaceURI === W && node.localName === 'tr')
  if (rows.length < 2) throw problem('三线表至少需要表头和一行数据', 422)
  for (const cell of elements(rows[0]).filter(node => node.namespaceURI === W && node.localName === 'tc')) {
    const cellProps = child(cell, 'tcPr') || cell.insertBefore(doc.createElementNS(NS, 'w:tcPr'), cell.firstChild)
    const cellBorders = ensure(cellProps, 'tcBorders')
    setBorder(cellBorders, 'bottom', 'single')
  }
  parts['word/document.xml'] = encode(doc)
  return zipSync(parts)
}

function border(table, name) {
  const props = elements(table).find(n => n.namespaceURI === W && n.localName === 'tblPr')
  const borders = elements(props).find(n => n.namespaceURI === W && n.localName === 'tblBorders')
  return elements(borders).find(n => n.namespaceURI === W && n.localName === name)
}

function value(node, name) { return node?.getAttributeNS(W, name) || node?.getAttribute(`w:${name}`) || '' }

export function validateThreeLineTable(prepared, output, tableIndex) {
  const doc = xml(unpack(output)['word/document.xml'])
  const body = doc.getElementsByTagNameNS(W, 'body')[0]
  const tables = elements(body).filter(n => n.namespaceURI === W && n.localName === 'tbl')
  const table = tables[tableIndex]
  if (!table) throw problem('三线表结果中目标表格不存在', 422)
  const rows = elements(table).filter(n => n.namespaceURI === W && n.localName === 'tr')
  if (rows.length < 2) throw problem('三线表至少需要表头和一行数据', 422)
  const required = ['top', 'bottom'].every(name => value(border(table, name), 'val') === 'single')
  const forbidden = ['left', 'right', 'insideH', 'insideV'].every(name => ['nil', ''].includes(value(border(table, name), 'val')))
  const headerProps = elements(rows[0]).filter(n => n.namespaceURI === W && n.localName === 'tc')
    .flatMap(cell => elements(cell).filter(n => n.namespaceURI === W && n.localName === 'tcPr'))
  const headerBorders = headerProps.flatMap(props => elements(props).filter(n => n.namespaceURI === W && n.localName === 'tcBorders'))
  const headerBottom = headerBorders.length > 0 && headerBorders.every(b => value(elements(b).find(n => n.localName === 'bottom'), 'val') === 'single')
  if (!required || !forbidden || !headerBottom) throw problem('三线表边框回读失败：需要顶线、表头下线和底线，且无竖线及其他横线', 422)
  return { status: 'passed', checks: ['原生编辑器设置三线表边框', '顶线、表头下线和底线回读通过', '竖线及内部横线已清除'], limitations: ['页面线宽视觉效果仍需编辑器页面验收'] }
}
