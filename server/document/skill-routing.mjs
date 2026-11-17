import {documentCoverage} from './coverage.mjs'
// Suggestions are evidence for skill selection, never authority to change formatting.
export function recommendSkills(data) {
  const text = data.paragraphs.slice(0, 160).map(p => p.text).join('\n')
  const candidates = [
    ['docx-academic', /毕业论文|学位论文|本科毕业|硕士论文|博士论文/u],
    ['docx-contract', /合同|协议书|甲方|乙方/u],
    ['docx-report', /研究报告|工作报告|工作总结|调研报告/u],
    ['docx-technical', /操作手册|使用说明|技术方案|接口文档/u],
    ['docx-official', /发文字号|主送机关|版记|党政机关/u],
  ].flatMap(([name, pattern]) => {
    const evidence = text.match(pattern)?.[0]
    return evidence ? [{ name, evidence }] : []
  })
  return { candidates, uncertain: candidates.length !== 1,
    objectSkills: [data.tables.length && 'docx-objects', data.elements.images && 'docx-objects',
      data.elements.fields && 'docx-toc-sections'].filter((s, i, a) => s && a.indexOf(s) === i),
    note: '仅为前160段文字线索；可能混合类型或误命中，不能据此自动套用字号和样式。' }
}

export function projectInspection(data, { offset = 0, limit = 40, indices, view = 'paragraphs', detail = indices ? 'formatting' : 'structure' } = {}) {
  if (!['structure', 'formatting'].includes(detail)) throw new Error('未知详细程度')
  if (!Number.isInteger(offset) || offset < 0 || !Number.isInteger(limit) || limit < 1 || limit > 80) throw new Error('分页参数无效')
  const collections = { paragraphs: data.paragraphs, tables: data.tables, cells: data.cells || [], floatingObjects:data.floatingObjects || [], sections: data.sections,
    captions: data.captions, figures: data.captionTargets.figures, captionTables: data.captionTargets.tables }
  if (view !== 'summary' && !Object.hasOwn(collections, view)) throw new Error('未知结构视图')
  if (indices && (view !== 'paragraphs' || !Array.isArray(indices) || indices.length > 80 || indices.some(i => !Number.isInteger(i) || i < 0 || i >= data.paragraphs.length))) throw new Error('段落索引无效')
  const source = view === 'summary' ? [] : collections[view]
  const selected = indices ? [...new Set(indices)].map(i => data.paragraphs[i]) : source.slice(offset, offset + limit)
  const compact = Boolean(indices) || detail === 'formatting'
  const context = compact ? {} : { pageNumbers:data.pageNumbers, partCount:data.partCount, elements:data.elements, layoutRisks:data.layoutRisks,
    totals: Object.fromEntries(Object.entries(collections).map(([name, values]) => [name, values.length])),
    recommendations: recommendSkills(data), coverage: documentCoverage(data) }
  return { hash: data.hash, ...context, view, detail,
    items: view === 'paragraphs' && detail === 'structure' ? selected.map(p => ({
      index: p.index, text: p.text.slice(0, 160), textTruncated: Boolean(p.textTruncated || p.text.length > 160),
      style: p.style, styleName:p.styleName, heading: p.heading, role: p.role,
      ...(p.tableNestingLevel ? { cell: p.cell, tableNestingLevel: p.tableNestingLevel } : {}),
      ...(p.inlinePictureCount ? { inlinePictureCount: p.inlinePictureCount } : {}),
      ...(p.leadingWhitespace?.prefix ? { leadingWhitespace: p.leadingWhitespace } : {}),
      ...(p.numbering ? { numbering: p.numbering } : {}),
    })) : view === 'floatingObjects' && detail === 'structure' ? selected.map(o=>({objectIndex:o.objectIndex,readOnly:o.readOnly,text:o.text,flowCapability:o.flowCapability,textBoxCount:o.textBoxCount})) : selected,
    readingHint: detail === 'structure' ? '结构摘要不是格式验收；修改前或验收时用 indices 定点读取完整格式。textTruncated=true 时不得用摘要进行局部文本匹配。' : undefined,
    pagination: { offset: indices ? null : offset, returned: selected.length, total: source.length,
      nextOffset: !indices && offset + limit < source.length ? offset + limit : null,
      complete: view === 'summary' ? false : !indices && offset === 0 && selected.length === source.length },
    ...(compact ? {} : { limitations: data.limitations }) }
}
