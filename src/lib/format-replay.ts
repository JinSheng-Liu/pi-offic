import { compileFormatOps } from '../../shared/editor-format-ops.mjs'
import { replayFingerprint } from '../../shared/replay-fingerprint.mjs'
import type { FormatReplay } from '../types'
import { bodyStoryRoot, storyParagraphs, paragraphTextOf } from '@docx-editor.dev/core/store'

export type ReplayLocation = { blockId: string; text: string }

const compileFormatOpsWithIds = compileFormatOps as unknown as (
  part: unknown,
  indices: number[],
  op: Record<string, unknown>,
  paragraphIds: string[],
) => unknown[]

const BATCH_SIZE = 20

async function sha256Hex(input: string): Promise<string> {
  const bytes = new TextEncoder().encode(input)
  const digest = await crypto.subtle.digest('SHA-256', bytes)
  return Array.from(new Uint8Array(digest), b => b.toString(16).padStart(2, '0')).join('')
}

function nextPaintFrame(): Promise<void> {
  if (typeof requestAnimationFrame === 'function') {
    // A promise resolved inside rAF resumes before paint. Resume in a fresh
    // task instead, so layout, painting and queued input can run between writes.
    return new Promise(resolve => requestAnimationFrame(() => setTimeout(resolve, 0)))
  }
  return Promise.resolve()
}

export async function applyFormatReplay(
  editor: any,
  replay: FormatReplay,
  onBatch?: (done: number, total: number, index?: number, location?: ReplayLocation) => void | Promise<void>,
): Promise<void> {
  if (!editor || replay?.kind !== 'format-v1') {
    throw new Error('无效的格式回放数据')
  }
  if (
    (typeof editor.getMode === 'function' && editor.getMode() !== 'edit') ||
    (typeof editor.surface?.editingMode === 'function' && editor.surface.editingMode() !== 'edit')
  ) {
    throw new Error('编辑器处于只读模式，拒绝应用格式回放')
  }
  const session = editor.surface?.session
  if (!session || typeof session.part !== 'function' || typeof editor.surface.applyAutomationOps !== 'function') {
    throw new Error('编辑器会话不可用')
  }

  const initialPart = session.part()
  const currentBefore = await sha256Hex(replayFingerprint(initialPart))
  if (currentBefore !== replay.before) {
    throw new Error('编辑器文档指纹与回放起点不一致')
  }
  // Formatting preserves paragraph order and identities. Keep only IDs here;
  // properties must still be read from the latest part after each write.
  const root = bodyStoryRoot(initialPart)
  if (!root) throw new Error('编辑器正文不可用')
  const paragraphIds = storyParagraphs(root).map(node => node.id)

  const queue: { index: number; op: Record<string, unknown> }[] = []
  for (const target of replay.targets || []) {
    for (const index of target.indices || []) {
      queue.push({ index, op: target.op })
    }
  }

  const total = queue.length
  let done = 0

  while (done < total) {
    if (
      (typeof editor.getMode === 'function' && editor.getMode() !== 'edit') ||
      (typeof editor.surface?.editingMode === 'function' && editor.surface.editingMode() !== 'edit')
    ) {
      throw new Error('编辑器在回放期间切换为只读')
    }
    const slice = queue.slice(done, done + BATCH_SIZE)
    // Adjacent identical operations share a compilation. Different operations
    // must see preceding writes, including overlapping paragraph properties.
    const groups: { indices: number[]; op: Record<string, unknown>; key: string }[] = []
    for (const item of slice) {
      const key = JSON.stringify(item.op)
      const previous = groups.at(-1)
      if (!previous || previous.key !== key) groups.push({ indices: [item.index], op: item.op, key })
      else previous.indices.push(item.index)
    }
    for (const group of groups) {
      const batchOps = compileFormatOpsWithIds(session.part(), group.indices, group.op, paragraphIds)
      if (batchOps.length > 0) {
        const applied = editor.surface.applyAutomationOps(() => batchOps)
        if (!applied || applied.committed === false || applied.rejected === true) {
          throw new Error('编辑器拒绝应用格式回放批次')
        }
      }
    }
    done += slice.length
    await nextPaintFrame()
    const index = slice[0]?.index
    const blockId = index === undefined ? undefined : paragraphIds[index]
    if (onBatch) await onBatch(done, total, index, blockId ? {
      blockId, text: (paragraphTextOf(initialPart, blockId) || '').trim().slice(0, 60),
    } : undefined)
  }

  const currentAfter = await sha256Hex(replayFingerprint(session.part()))
  if (currentAfter !== replay.after) {
    throw new Error('编辑器文档回放终点指纹校验未通过')
  }
}
