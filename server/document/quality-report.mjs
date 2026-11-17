import {documentCoverage} from './coverage.mjs'
export function qualityReport(data, validation, { contentVerified = true } = {}) {
  const risks = data.layoutRisks?.warnings || []
  const limitations = [...new Set([...(validation.limitations || []),
    '未使用 Microsoft Word 进行页面渲染；分页、重叠、留白及实际字体呈现未验证。',
    ...risks.map(risk => risk.message),
    ...(contentVerified ? [] : ['当前操作未验证正文内容保持。']),
  ])]
  return { documentCoverage:documentCoverage(data), status: limitations.length ? 'partial' : 'passed', checks: validation.checks,
    limitations, coverage: { package: 'verified', targetedFormatting: validation.checks.includes('目标格式重读验证') ? 'verified' : 'not-applicable',
      floatingObjects: data.layoutRisks?.floatingCount > 0 ? 'not-fully-processed' : 'no-structural-risk-detected', visualLayout: 'not-verified', content: contentVerified ? 'verified' : 'not-verified' },
    layoutRisks: data.layoutRisks || {warnings:[],visualVerified:false} }
}
