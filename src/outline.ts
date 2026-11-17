export type OutlineItem = {
  text: string
  level: number
  blockId: string
}

const STANDALONE_HEADINGS = /^(摘要|abstract|目录|目\s+录|参考文献|致谢|附录|前言|绪论|结论)$/i
const CHAPTER_HEADING = /^(第[一二三四五六七八九十百千万零〇\d]+[章节篇部]|[一二三四五六七八九十百]+、)/
const SECTION_HEADING = /^(第[一二三四五六七八九十百千万零〇\d]+节|（[一二三四五六七八九十百]+）|\([一二三四五六七八九十百]+\))/
const NUMERIC_HEADING = /^(\d+(?:\.\d+){1,3})(?:\s+|(?=[A-Za-z\u4e00-\u9fa5]))/

export function recognizeHeading(raw: string): number | null {
  const text = String(raw || '').trim()
  if (!text || text.length > 60) return null
  if (/[。！？；!?;]/.test(text)) return null
  if (/^\d{4}\s*年|\d+(?:\.\d+)?%$/.test(text)) return null

  if (STANDALONE_HEADINGS.test(text)) return 0
  if (CHAPTER_HEADING.test(text)) {
    return /^第[一二三四五六七八九十百千万零〇\d]+节/.test(text) ? 1 : 0
  }
  if (SECTION_HEADING.test(text)) return 1

  const numeric = text.match(NUMERIC_HEADING)
  if (numeric) {
    const depth = numeric[1]!.split('.').length - 1
    return Math.min(Math.max(depth, 1), 3)
  }
  return null
}

export function visibleOutline(items: OutlineItem[], collapsed: Set<string>): OutlineItem[] {
  const result: OutlineItem[] = []
  let hiddenAboveLevel: number | null = null

  for (const item of items) {
    if (hiddenAboveLevel !== null) {
      if (item.level > hiddenAboveLevel) continue
      hiddenAboveLevel = null
    }
    result.push(item)
    if (collapsed.has(item.blockId)) {
      hiddenAboveLevel = item.level
    }
  }
  return result
}
