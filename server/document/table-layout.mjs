import { executeAutomationBatch } from './editor-automation.mjs'
import { zipSync } from 'fflate'
import { XMLSerializer } from '@xmldom/xmldom'
import { unpack, xml, W, compareProtected, problem } from './package.mjs'

const children = (node, name) => Array.from(node?.childNodes || []).filter(n => n.namespaceURI === W && n.localName === name)
const enabled = node => Boolean(node) && !['0', 'false', 'off', 'no'].includes((node.getAttributeNS(W, 'val') || 'true').toLowerCase())
const encode = document => new TextEncoder().encode(new XMLSerializer().serializeToString(document))
const rowOrder = ['cnfStyle', 'divId', 'gridBefore', 'gridAfter', 'wBefore', 'wAfter', 'cantSplit', 'trHeight', 'tblHeader', 'tblCellSpacing', 'jc', 'hidden', 'ins', 'del', 'trPrChange']

function locate(document, index) {
  const body = document.getElementsByTagNameNS(W, 'body')[0]
  const tables = Array.from(body.getElementsByTagNameNS(W, 'tbl'))
  const table = tables[index]
  if (!table) throw problem('表格索引超出范围', 422)
  for (let node = table; node && node !== body; node = node.parentNode) {
    if (['AlternateContent', 'ins', 'del'].includes(node.localName)) throw problem('分支或修订中的表格暂不支持布局修改', 422)
  }
  const rows = children(table, 'tr')
  if (!rows.length || rows.length !== table.getElementsByTagNameNS(W, 'tr').length) throw problem('嵌套表格或包装行需单独定位，不能按当前行索引修改', 422)
  return { tables, rows }
}

export function formatTables(bytes, op, signal) {
  signal?.throwIfAborted()
  const originalParts = unpack(bytes)
  const original = xml(originalParts['word/document.xml'])
  const { tables, rows } = locate(original, op.tableIndex)
  const targets = []
  let headerRowCount
  if (op.scope === 'tableRows') {
    const selected = new Set(op.rowIndices)
    if (op.rowIndices.some(index => !rows[index])) throw problem('表格行索引超出范围', 422)
    if (op.table.repeatHeader !== undefined) {
      const flags = rows.map((row, index) => selected.has(index) ? op.table.repeatHeader : enabled(children(children(row, 'trPr')[0], 'tblHeader')[0]))
      headerRowCount = flags.filter(Boolean).length
      if (flags.some((flag, index) => flag && index >= headerRowCount)) throw problem('重复表头必须是从首行开始的连续行', 422)
    }
    for (const rowIndex of selected) {
      for (const [property, key, value] of [['repeatHeader', 'tblHeader', op.table.repeatHeader], ['cantSplit', 'cantSplit', op.table.cantSplit], ['autoHeight', 'trHeight', 0]]) {
        if (op.table[property] === undefined || property === 'autoHeight' && !op.table.autoHeight) continue
        targets.push({ rowIndex, group: 'trPr', key, value })
      }
    }
  } else {
    const seen = new Set()
    for (const { rowIndex, cellIndex } of op.cellTargets) {
      const key = `${rowIndex}:${cellIndex}`
      if (seen.has(key)) continue
      seen.add(key)
      const row = rows[rowIndex], cell = children(row, 'tc')[cellIndex]
      if (!cell) throw problem('单元格物理索引超出范围', 422)
      if (children(row, 'tc').length !== row.getElementsByTagNameNS(W, 'tc').length) throw problem('包装或嵌套单元格需单独定位', 422)
      targets.push({ rowIndex, cellIndex, group: 'tcPr', key: 'vAlign', value: op.table.verticalAlignment.toLowerCase() })
    }
  }
  const owner = (rowList, target) => target.group === 'trPr' ? rowList[target.rowIndex] : children(rowList[target.rowIndex], 'tc')[target.cellIndex]
  for (const target of targets) {
    const group = children(owner(rows, target), target.group)[0]
    if (children(group, target.key).length > 1) throw problem('存在重复表格属性，不能安全修改', 422)
  }

  let output = bytes
  if (op.scope === 'tableCells' || headerRowCount !== undefined) {
    output = executeAutomationBatch(bytes, host => {
      const read = operation => {
        const result = host.execute({ operations: [operation] })
        if (!result.ok) throw problem('无法读取原生表格定位信息', 422)
        return result.results[0].value
      }
      const document = read({ op: 'getDocument' }).handle
      const body = read({ op: 'getBody', document }).handle
      const handles = read({ op: 'getTables', scope: { body } }).handles
      if (handles.length !== tables.length) throw problem('原生表格索引与文档结构不一致', 422)
      const table = handles[op.tableIndex]
      if (headerRowCount !== undefined) return [{ op: 'updateTable', table, mutation: { kind: 'properties', headerRowCount } }]
      return targets.map(target => ({ op: 'updateTableCell',
        cell: read({ op: 'getTableCell', table, rowIndex: target.rowIndex, cellIndex: target.cellIndex }).handle,
        properties: { verticalAlignment: op.table.verticalAlignment } }))
    }).bytes
  }
  const outputParts = unpack(output), document = xml(outputParts['word/document.xml'])
  const outputRows = locate(document, op.tableIndex).rows
  // Only these two row properties lack an equivalent public Automation write.
  for (const target of targets.filter(target => ['cantSplit', 'trHeight'].includes(target.key))) {
    const row = owner(outputRows, target)
    let group = children(row, 'trPr')[0]
    if (!group) { group = document.createElementNS(W, 'w:trPr'); row.insertBefore(group, row.firstChild) }
    const old = children(group, target.key)[0]
    const node = document.createElementNS(W, `w:${target.key}`)
    node.setAttributeNS(W, 'w:val', String(target.value))
    if (target.key === 'trHeight') node.setAttributeNS(W, 'w:hRule', 'auto')
    if (old) group.replaceChild(node, old)
    else group.insertBefore(node, Array.from(group.childNodes).find(n => n.namespaceURI === W && rowOrder.indexOf(n.localName) > rowOrder.indexOf(target.key)) || null)
  }
  outputParts['word/document.xml'] = encode(document)
  const resultBytes = zipSync(outputParts)

  // Verify the actual output, then restore only allowed properties in a comparison copy.
  for (const target of targets) {
    const row = owner(outputRows, target), group = children(row, target.group)[0]
    const nodes = children(group, target.key), actual = nodes[0]
    const valid = ['tblHeader', 'cantSplit'].includes(target.key) ? enabled(actual) === target.value :
      actual?.getAttributeNS(W, 'val') === String(target.value) && (target.key !== 'trHeight' || actual.getAttributeNS(W, 'hRule') === 'auto')
    if (nodes.length > 1 || !valid) throw problem('目标表格属性回读失败', 422)
    const originalGroup = children(owner(rows, target), target.group)[0]
    const previous = children(originalGroup, target.key)[0]
    if (actual) group.removeChild(actual)
    if (previous) {
      if (!group) throw problem('表格属性容器缺失', 422)
      const siblings = Array.from(originalGroup.childNodes)
      const nextNames = siblings.slice(siblings.indexOf(previous) + 1).map(n => n.localName)
      group.insertBefore(document.importNode(previous, true), Array.from(group.childNodes).find(n => nextNames.includes(n.localName)) || null)
    }
    if (group && !originalGroup && !group.childNodes.length) row.removeChild(group)
  }
  compareProtected(bytes, zipSync({ ...outputParts, 'word/document.xml': encode(document) }), { formatting: false })
  signal?.throwIfAborted()
  return { bytes: resultBytes, affected: targets.length, validation: { status: 'passed',
    checks: ['原生表格属性执行及回读', '仅补充自动行高与禁止跨页', '非目标内容与属性保持'],
    limitations: ['超高表格行仍可能跨页；页面视觉未验证'] } }
}
