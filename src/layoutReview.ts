export type LayoutFragment = {
  kind: string
  paragraphId?: string
  box?: { y?: number; height?: number }
  lines?: { spans?: { text?: string }[] }[]
}

export type LayoutPage = {
  index: number
  fragments?: LayoutFragment[]
  contentBox?: { height?: number }
}

export type LayoutFinding = {
  paragraphId: string
  page: number
  message: string
}

function fragmentText(fragment: LayoutFragment): string {
  return (fragment.lines || [])
    .flatMap(line => (line.spans || []).map(span => span.text || ''))
    .join('')
    .trim()
}

export function reviewLayout(pages: LayoutPage[], headingIds: Set<string>): LayoutFinding[] {
  const findings: LayoutFinding[] = []
  const seenHeadings = new Set<string>()

  for (let pageIdx = 0; pageIdx < pages.length; pageIdx++) {
    const page = pages[pageIdx]!
    const paragraphs = (page.fragments || []).filter(f => f.kind === 'paragraph' && f.paragraphId)
    if (!paragraphs.length) continue

    for (let i = 0; i < paragraphs.length; i++) {
      const frag = paragraphs[i]!
      const id = frag.paragraphId!
      if (!headingIds.has(id)) continue

      const lines = frag.lines || []
      if (lines.length > 1 && !seenHeadings.has(`tail:${id}`)) {
        const lastLineText = (lines.at(-1)?.spans || []).map(s => s.text || '').join('').trim()
        if (lastLineText.length > 0 && lastLineText.length <= 3) {
          seenHeadings.add(`tail:${id}`)
          findings.push({
            paragraphId: id,
            page: page.index + 1,
            message: `标题“${fragmentText(frag).slice(0, 18)}”最后一行仅 ${lastLineText.length} 个字，建议检查断行`,
          })
        }
      }

      if (i === paragraphs.length - 1 && pageIdx + 1 < pages.length) {
        const nextParagraphs = (pages[pageIdx + 1]?.fragments || []).filter(f => f.kind === 'paragraph' && f.paragraphId)
        if (nextParagraphs.length > 0 && nextParagraphs[0]!.paragraphId !== id) {
          findings.push({
            paragraphId: id,
            page: page.index + 1,
            message: `标题“${fragmentText(frag).slice(0, 18)}”位于页尾，与后续正文分处两页`,
          })
        }
      }
    }

    if (pageIdx + 1 < pages.length) {
      const contentHeight = page.contentBox?.height || 0
      const last = paragraphs.at(-1)!
      const usedBottom = (last.box?.y || 0) + (last.box?.height || 0)
      if (contentHeight > 0 && contentHeight - usedBottom > contentHeight * 0.35) {
        findings.push({
          paragraphId: last.paragraphId!,
          page: page.index + 1,
          message: `第 ${page.index + 1} 页底部存在较大留白，请确认是否为有意分页`,
        })
      }
    }
  }

  return findings
}
