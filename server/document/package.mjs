import { createHash } from 'node:crypto'
import { AsyncLocalStorage } from 'node:async_hooks'
import { unzipSync } from 'fflate'
import { DOMParser } from '@xmldom/xmldom'
export const MAX_BYTES = 20 * 1024 * 1024
export const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
export const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
export const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
export const V = 'urn:schemas-microsoft-com:vml'
export const digest = bytes => createHash('sha256').update(bytes).digest('hex')
export function problem(message, status = 400) { return Object.assign(new Error(message), { status }) }

const packageCache = new AsyncLocalStorage()
export function withPackageCache(operation) {
  return packageCache.run(new Map(), operation)
}
const copyParts = parts => Object.fromEntries(Object.entries(parts).map(([name, bytes]) => [name, bytes.slice()]))

export function unpack(bytes) {
  if (!bytes.length || bytes.length > MAX_BYTES) throw problem('DOCX 文件需在 20 MB 以内')
  const cache = packageCache.getStore()
  const key = cache ? digest(bytes) : undefined
  if (cache?.has(key)) return copyParts(cache.get(key))
  let total = 0, count = 0
  let parts
  try {
    parts = unzipSync(new Uint8Array(bytes), { filter(entry) {
      total += entry.originalSize
      if (++count > 3000 || total > 100 * 1024 * 1024 || entry.originalSize > 20 * 1024 * 1024)
        throw problem('DOCX 解压规模超过限制')
      if (entry.name.includes('..') || entry.name.startsWith('/')) throw problem('无效的 DOCX 包路径')
      return true
    } })
  } catch (error) { throw problem(error.status ? error.message : '不是有效的 DOCX ZIP 文件') }
  for (const name of ['[Content_Types].xml', '_rels/.rels', 'word/document.xml'])
    if (!parts[name]) throw problem(`DOCX 缺少必要部分：${name}`)
  if (Object.keys(parts).some(n => /vbaProject|_xmlsignatures/i.test(n))) throw problem('暂不接受宏或数字签名文档')
  for (const [name, data] of Object.entries(parts)) if (/\.(xml|rels)$/.test(name)) xml(data)
  // Retain at most two small packages within one job; callers may mutate their
  // own parts, so never expose the cache's arrays. Failed validation is not cached.
  if (cache && total <= 8 * 1024 * 1024) {
    if (cache.size >= 2) cache.delete(cache.keys().next().value)
    cache.set(key, copyParts(parts))
  }
  return parts
}

export function xml(bytes) {
  const source = new TextDecoder().decode(bytes)
  if (/<!DOCTYPE|<!ENTITY/i.test(source)) throw problem('文档包含不支持的 XML 实体声明')
  try {
    return new DOMParser({ onError(level, message) { if (level !== 'warning') throw new Error(message) } })
      .parseFromString(source, 'application/xml')
  } catch { throw problem('文档包含无效 XML') }
}

// Ignore only explicitly supported formatting properties. Preserve content, numbering,
// section properties, relationships, unknown elements and tracked revisions in the comparison.
const mutable = new Set(['rFonts', 'sz', 'szCs', 'b', 'bCs', 'i', 'iCs', 'u', 'color', 'strike', 'dstrike', 'vertAlign', 'jc', 'spacing', 'ind', 'pStyle', 'outlineLvl', 'keepNext', 'keepLines', 'widowControl', 'pageBreakBefore'])
const booleanProperties = new Set(['b','bCs','i','iCs','strike','dstrike','caps','smallCaps','vanish','webHidden','noProof','rtl','cs','emboss','imprint','outline','shadow','snapToGrid','keepNext','keepLines','widowControl','pageBreakBefore','kinsoku','wordWrap','overflowPunct','topLinePunct','autoSpaceDE','autoSpaceDN','bidi','adjustRightInd','contextualSpacing','suppressLineNumbers','suppressAutoHyphens','mirrorIndents','rtlGutter'])
export function canonical(node, formatting = false) {
  if (node.nodeType === 3 || node.nodeType === 4) {
    const parent = node.parentNode
    const contentText = (parent?.namespaceURI === W && ['t','delText','instrText','delInstrText'].includes(parent.localName)) ||
      ([A, M].includes(parent?.namespaceURI) && parent.localName === 't')
    return contentText || node.data.trim() ? node.data : ''
  }
  if (node.nodeType !== 1) return ''
  if (formatting && node.namespaceURI === W && mutable.has(node.localName) &&
      ['rPr', 'pPr'].includes(node.parentNode?.localName)) return ''
  const attributes = Array.from(node.attributes).filter(a => a.namespaceURI !== 'http://www.w3.org/2000/xmlns/')
    .map(a => [a.namespaceURI || '', a.localName,
      a.namespaceURI === 'http://schemas.openxmlformats.org/markup-compatibility/2006' && a.localName === 'Ignorable'
        ? [...new Set(a.value.trim().split(/\s+/).filter(Boolean).map(prefix=>node.lookupNamespaceURI(prefix)||`unresolved:${prefix}`))].sort().join(' ') :
      ((node.namespaceURI === W && node.localName === 'hyperlink' && a.namespaceURI === W && a.localName === 'history') ||
       (node.namespaceURI === 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape' && node.localName === 'bodyPr' && !a.namespaceURI && ['anchorCtr','fromWordArt','rtlCol','compatLnSpc','forceAA','spcFirstLastPara'].includes(a.localName)) ||
       (node.namespaceURI === W && booleanProperties.has(node.localName) && a.namespaceURI === W && a.localName === 'val') ||
       (node.namespaceURI === 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing' && node.localName === 'anchor' && !a.namespaceURI && ['allowOverlap','behindDoc','layoutInCell','locked','simplePos'].includes(a.localName)) ||
       (node.namespaceURI === A && ['graphicFrameLocks','picLocks'].includes(node.localName) && !a.namespaceURI && ['noChangeAspect','noChangeArrowheads','noCrop','noEditPoints','noGrp','noMove','noResize','noRot','noSelect','noAdjustHandles','noChangeShapeType','noDrilldown'].includes(a.localName)) ||
       (node.namespaceURI === 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape' && node.localName === 'cNvSpPr' && !a.namespaceURI && a.localName === 'txBox') ||
       (node.namespaceURI === 'http://schemas.microsoft.com/office/drawing/2010/main' && node.localName === 'useLocalDpi' && !a.namespaceURI && a.localName === 'val'))
        ? ({'0':'false','off':'false','1':'true','on':'true'}[a.value] || a.value) : (node.namespaceURI === V && node.localName === 'shape' && a.namespaceURI === 'urn:schemas-microsoft-com:office:office' && a.localName === 'gfxdata' ? a.value.replace(/[ \t\r\n]/g, '') : a.value)]).sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b)))
  const children = Array.from(node.childNodes).map(n => canonical(n, formatting)).filter(Boolean)
  if (formatting && node.namespaceURI === W && ['rPr', 'pPr'].includes(node.localName) && !children.length && !attributes.length) return ''
  // Hash each subtree once. Embedding serialized child strings recursively caused
  // exponential JSON escaping for nested drawings and exhausted V8 string limits.
  return createHash('sha256').update(JSON.stringify([node.namespaceURI || '', node.localName, attributes, children])).digest('hex')
}

export function compareProtected(before, after, { formatting = true } = {}) {
  const a = unpack(before), b = unpack(after)
  // ZIP directory markers are optional; only actual files are document parts.
  const missing = Object.keys(a).filter(name => !name.endsWith('/') && !b[name])
  if (missing.length) throw problem(`保存会丢失文档部分：${missing.join(', ')}`, 422)
  for (const name of Object.keys(a)) {
    if (name.endsWith('/')) continue
    if (digest(a[name]) === digest(b[name])) continue
    if (!formatting && name === 'word/document.xml') continue
    if (formatting && name === 'word/document.xml') {
      if (canonical(xml(a[name]).documentElement, true) !== canonical(xml(b[name]).documentElement, true))
        throw problem('格式修改改变了受保护的正文结构，已取消保存', 422)
    } else if (formatting && name === 'word/styles.xml') {
      // Built-in heading style materialization may append definitions. Existing styles must survive.
      const oldStyles = Array.from(xml(a[name]).getElementsByTagNameNS(W, 'style'))
      const newStyles = Array.from(xml(b[name]).getElementsByTagNameNS(W, 'style'))
      for (const style of oldStyles) {
        const next = newStyles.find(s => s.getAttributeNS(W, 'styleId') === style.getAttributeNS(W, 'styleId'))
        if (!next || canonical(style) !== canonical(next)) throw problem('原有样式定义发生非目标变化', 422)
      }
    } else if (formatting && digest(a[name]) !== digest(b[name])) {
      // Serializers may normalize XML prefixes/attribute order; compare semantics independently.
      if (!/\.(xml|rels)$/.test(name) || canonical(xml(a[name]).documentElement) !== canonical(xml(b[name]).documentElement))
        throw problem(`非目标部分发生变化：${name}`, 422)
    }
  }
  return { status: 'passed', checks: ['DOCX 包可读', '原有部分保留', ...(formatting ? ['正文结构保持', '非目标部分保持'] : [])],
    limitations: ['未验证 Microsoft Word 排版效果', ...(formatting ? [] : ['手动编辑仅验证包完整性，不验证内容一致性'])] }
}

