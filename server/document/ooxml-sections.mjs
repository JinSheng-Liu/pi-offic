import { executeAutomationBatch } from './editor-automation.mjs'
import { zipSync } from 'fflate'
import { XMLSerializer } from '@xmldom/xmldom'
import { unpack, xml, W, problem } from './package.mjs'
import { sectionNodes } from './structure.mjs'
import { validatePage, validateMargins } from './rules.mjs'
export function applySectionOperations(bytes, operations) {
  const parts = unpack(bytes)
  const document = xml(parts['word/document.xml'])
  const body = document.getElementsByTagNameNS(W, 'body')[0]
  const sections = sectionNodes(body)
  const child = (node, name) => Array.from(node.childNodes).find(n => n.namespaceURI === W && n.localName === name)
  const attributes = (node, keys) => Object.fromEntries(keys.map(key => [key, Number(node?.getAttributeNS(W, key) || 0)]))
  const planned = sections.map(section => ({
    size: attributes(child(section, 'pgSz'), ['w', 'h']),
    margins: attributes(child(section, 'pgMar'), ['top', 'bottom', 'left', 'right', 'gutter']),
    hasSize: Boolean(child(section, 'pgSz')),
    hasMargins: Boolean(child(section, 'pgMar')),
  }))
  for (const operation of operations) {
    if (operation.page) validatePage(operation.page)
    if (operation.margins) validateMargins(operation.margins)
    const indices = operation.sectionIndices
    if (!Array.isArray(indices) || !indices.length || indices.some(index => !Number.isInteger(index) || index < 0 || index >= sections.length)) throw problem('分节索引无效或超出范围')
    for (const index of new Set(indices)) {
      const plan = planned[index]
      if (operation.page) {
        plan.hasSize = true
        if (operation.page.widthTwips !== undefined) plan.size.w = operation.page.widthTwips
        if (operation.page.heightTwips !== undefined) plan.size.h = operation.page.heightTwips
      }
      if (operation.margins) {
        plan.hasMargins = true
        Object.assign(plan.margins, operation.margins)
      }
    }
  }
  for (const { size, margins, hasSize, hasMargins } of planned) {
    if (!hasSize || !hasMargins) continue
    if (margins.left + margins.right + margins.gutter >= size.w || margins.top + margins.bottom >= size.h) {
      throw problem('页边距总和必须小于纸张尺寸，正文区域不能为空')
    }
  }
  {
    const native = executeAutomationBatch(bytes, host => {
      const read = operation => {
        const result = host.execute({ operations: [operation] })
        if (!result.ok) throw problem('无法读取编辑器分节', 422)
        return result.results[0].value
      }
      const handle = read({ op: 'getDocument' }).handle
      const handles = read({ op: 'getSections', document: handle }).handles
      const marginKeys = {
        top: 'topMargin', right: 'rightMargin', bottom: 'bottomMargin',
        left: 'leftMargin', header: 'headerDistance', footer: 'footerDistance',
      }
      return operations.flatMap(operation => {
        const setup = {}
        if (operation.page?.widthTwips !== undefined) setup.pageWidth = operation.page.widthTwips / 20
        if (operation.page?.heightTwips !== undefined) setup.pageHeight = operation.page.heightTwips / 20
        if (operation.page?.orientation !== undefined) setup.orientation = operation.page.orientation
        for (const [key, value] of Object.entries(operation.margins || {})) {
          if (key !== 'gutter') setup[marginKeys[key]] = value / 20
        }
        if (!Object.keys(setup).length) return []
        return [...new Set(operation.sectionIndices)].map(index => ({
          op: 'setPageSetup', section: handles[index], setup,
        }))
      })
    })
    // Only the unsupported gutter property is supplemented after native edits.
    const outputParts = unpack(native.bytes)
    const outputDocument = xml(outputParts['word/document.xml'])
    const outputSections = sectionNodes(outputDocument.getElementsByTagNameNS(W, 'body')[0])
    let patchedGutter = false
    for (const operation of operations) {
      if (operation.margins?.gutter === undefined) continue
      for (const index of new Set(operation.sectionIndices)) {
        const section = outputSections[index]
        if (!section) throw problem('原生编辑器保存后的分节数量不一致', 422)
        let margin = Array.from(section.childNodes).find(node => node.namespaceURI === W && node.localName === 'pgMar')
        if (!margin) {
          margin = outputDocument.createElementNS(W, 'w:pgMar')
          const next = Array.from(section.childNodes).find(node => node.namespaceURI === W &&
            ['paperSrc', 'pgBorders', 'lnNumType', 'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings', 'sectPrChange'].includes(node.localName))
          section.insertBefore(margin, next || null)
        }
        margin.setAttributeNS(W, 'w:gutter', String(operation.margins.gutter))
        patchedGutter = true
      }
    }
    if (patchedGutter) outputParts['word/document.xml'] = new TextEncoder().encode(new XMLSerializer().serializeToString(outputDocument))
    return {
      bytes: patchedGutter ? zipSync(outputParts) : native.bytes,
      affected: operations.reduce((sum, op) => sum + new Set(op.sectionIndices).size, 0),
    }
  }
}

