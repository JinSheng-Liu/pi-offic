export type HistoryTask = {
  id: string
  summary: string
  startRevision: number
  startedAt: string
  stepLabel?: string
  state?: string
  finishedAt?: string
}

export type Validation = {
  operationReceipt?: {
    inputRevision: number
    kind: string
    operations: Record<string, unknown>[]
    affected: number
    batchChecks: string[]
    requirementsVerified: boolean
    visualVerified: boolean
  }
  status: string
  checks: string[]
  limitations: string[]
  coverage?: Record<string, string>
  layoutRisks?: {
    floatingCount: number
    floatingTextBoxCount: number
    noWrapCount: number
    tocInstructions: number
    warnings: { code: string; count: number; message: string }[]
    visualVerified: boolean
  }
}

export type Version = {
  task?: HistoryTask
  revision: number
  size: number
  time: string
  summary: string
  validation: Validation
}

export type HistoryGroup = {
  id: string
  task?: HistoryTask
  versions: Version[]
}

export type DocumentMeta = {
  id: string
  name: string
  revision: number
  updatedAt: string
  versions: Version[]
  savedVersions?: { revision: number; summary: string; kind: string; time: string; taskId?: string }[]
  pendingWork?: { taskId: string; startRevision: number; revision: number; summary: string } | null
  tasks?: Record<string, HistoryTask>
}

export type ChatStep = {
  id: string
  label: string
  state: string
}

export type ChatLine = {
  role: 'user' | 'assistant'
  text: string
  status?: string
  steps?: ChatStep[]
}

export type SectionInfo = {
  index: number
  page?: { widthTwips: number; heightTwips: number; orientation: string }
  margins?: Record<string, number>
}

export type FormatReplayTarget = {
  indices: number[]
  op: Record<string, unknown>
}

export type FormatReplay = {
  kind: 'format-v1'
  before: string
  after: string
  targets: FormatReplayTarget[]
}
