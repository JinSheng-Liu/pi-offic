import {isCaptionLabel} from './caption-label.mjs'
// Structural exclusions only. Normal style alone cannot identify a cover or body.
export function paragraphRole(p) {
  if (p.tableNestingLevel > 0) return 'table'
  if (p.inlinePictures.length) return 'inline-picture'
  if (p.generatedToc || /^TOC\d*$/i.test(p.styleId || '')) return 'toc'
  if (p.styleId === 'Caption' || isCaptionLabel(p.text)) return 'caption'
  if (/^(Title|Subtitle)$/i.test(p.styleId || '')) return 'title'
  if (p.outlineLevel <= 9) return 'heading'
  if (p.p?.pPr?.numPr) return 'list'
  if (!p.text.trim()) return 'empty'
  return 'unclassified-text'
}

export function isBodyCandidate(p) {
  return paragraphRole(p) === 'unclassified-text'
}
