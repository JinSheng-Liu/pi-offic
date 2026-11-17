import {
  computed,
  onMounted,
  onUnmounted,
  reactive,
  ref,
  shallowRef,
  watch,
} from 'vue'
import { createDocumentRefresh } from '@docx-editor.dev/vue'
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor'
import type { DocxEditorRef } from '@docx-editor.dev/vue'
import type {
  Validation,
  DocumentMeta,
  ChatLine,
  SectionInfo,
  FormatReplay,
} from '../types'
import { api, consumeEvents } from '../lib/api'
import { createEditorEventHandler } from '../lib/editor-events'
import { groupHistory } from '../lib/history'
import { applyFormatReplay } from '../lib/format-replay'
import type { ReplayLocation } from '../lib/format-replay'

const LF = String.fromCharCode(10)

export function useStudioWorkspace() {
  const navigationReady = ref(false)
  onMounted(() => {
    navigationReady.value = true
  })

  const editorRef = ref<DocxEditorRef | null>(null)
  const inputRef = ref<HTMLInputElement | null>(null)
  const documentData = shallowRef<ArrayBuffer | 'blank'>('blank')
  const documentKey = ref(0)
  const documents = ref<DocumentMeta[]>([])
  const documentId = ref('')
  const fileName = ref('未打开文档')
  const revision = ref(0)
  const previewRevision = ref(0)
  const dirty = ref(false)
  const saving = ref(false)
  const aiBusy = ref(false)
  const documentBusy = ref(false)
  const refreshingAfterTask = ref(false)
  let editGeneration = 0
  const modelConfigured = ref(false)
  const backendReady = ref(false)
  const assistantOpen = ref(true)
  const rightTab = ref<'chat' | 'history'>('chat')
  const recentOpen = ref(false)
  const settingsOpen = ref(false)
  const pageSettingsOpen = ref(false)
  const sections = ref<SectionInfo[]>([])
  const sectionSelection = ref('all')
  const pagePreset = ref('A4')
  const pageWidthMm = ref(210)
  const pageHeightMm = ref(297)
  const pageOrientation = ref<'portrait' | 'landscape'>('portrait')
  const pageMarginsMm = ref({
    top: 25.4,
    right: 25.4,
    bottom: 25.4,
    left: 25.4,
    gutter: 0,
  })
  const modelName = ref('gpt-4o-mini')
  const modelBaseUrl = ref('https://api.openai.com/v1')
  const modelApiKey = ref('')
  const settingsSaving = ref(false)
  const errorMessage = ref('')
  const closeDialog = ref<HTMLDialogElement | null>(null)
  const previewBehind = computed(
    () => Boolean(documentId.value) && previewRevision.value !== revision.value,
  )
  const closingBlocked = computed(
    () =>
      aiBusy.value ||
      saving.value ||
      documentBusy.value ||
      refreshingAfterTask.value ||
      previewBehind.value,
  )
  const chatInput = ref('')
  const activeTool = ref('')
  const followChanges = ref(true)
  const changeLocation = ref('')
  let latestChangeLocation: ReplayLocation | undefined
  let lastFollowTime = 0

  function pauseFollowing() {
    if (aiBusy.value) followChanges.value = false
  }

  function locateChange(location?: ReplayLocation) {
    if (location) {
      latestChangeLocation = location
      changeLocation.value = location.text
    }
    if (!followChanges.value) return
    if (location && performance.now() - lastFollowTime < 250) return
    lastFollowTime = performance.now()
    const editor = editorRef.value?.getEditor()
    if (!editor || !latestChangeLocation) return
    editor.scrollToBlock(latestChangeLocation.blockId, { block: 'centerIfNeeded' })
  }

  function resumeFollowing() {
    followChanges.value = true
    locateChange()
  }

  async function syncPreview(id: string, targetRevision: number) {
    latestChangeLocation = undefined
    const editor = editorRef.value?.getEditor()
    if (!editor || !('surface' in editor) || !('mountGeneration' in editor))
      throw new Error('编辑器挂载实例尚未就绪')
    const refresh = createDocumentRefresh(editor as DocxEditorInstance)
    const submission = await refresh.capture()
    try {
      const response = await fetch(
        `/api/documents/${id}/file?revision=${targetRevision}`,
      )
      if (!response.ok) throw new Error('无法读取更新后的文档')
      const result = await refresh.applyUpdate({
        submission,
        sequence: targetRevision,
        bytes: await response.arrayBuffer(),
      })
      if (!result.ok) throw new Error(`预览同步失败：${result.code}`)
      previewRevision.value = targetRevision
    } finally {
      refresh.finish(submission)
    }
  }
  const messagesByDocument = ref<Record<string, ChatLine[]>>({})
  const formatValidation = ref<Validation | null>(null)
  const abortController = shallowRef<AbortController | null>(null)
  let stopEditorChange: (() => void) | null = null

  watch(
    editorRef,
    (handle) => {
      stopEditorChange?.()
      stopEditorChange =
        handle?.getEditor()?.on('change', (change) => {
          if (!change.source) markDirty()
        }) || null
    },
    { flush: 'post' },
  )

  onUnmounted(() => stopEditorChange?.())

  const current = computed(() =>
    documents.value.find((d) => d.id === documentId.value),
  )
  const versions = computed(() =>
    [...(current.value?.versions || [])].reverse(),
  )
  const pendingWork = computed(() => current.value?.pendingWork)
  const historyGroups = computed(() => {
    const meta = current.value
    if (!meta?.savedVersions)
      return groupHistory(meta?.versions || [], meta?.tasks || {})
    return [...meta.savedVersions].reverse().flatMap((saved) => {
      const version = meta.versions.find((v) => v.revision === saved.revision)
      return version
        ? [
            {
              id: `saved-${saved.revision}`,
              versions: [
                {
                  ...version,
                  summary:
                    saved.summary +
                    (saved.kind === 'partial' ? '（部分结果）' : ''),
                },
              ],
            },
          ]
        : []
    })
  })

  async function savePartialResult() {
    if (closingBlocked.value || dirty.value || !pendingWork.value) return
    documentBusy.value = true
    try {
      await api(`/api/documents/${documentId.value}/publish`, {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ revision: revision.value }),
      })
      await refreshList()
    } catch (error) {
      errorMessage.value =
        error instanceof Error ? error.message : '部分结果保存失败'
    } finally {
      documentBusy.value = false
    }
  }
  const chatLines = computed(
    () => messagesByDocument.value[documentId.value] || [],
  )
  const sendBlockedReason = computed(() => {
    if (aiBusy.value) return '正在处理，可先输入下一条需求'
    if (documentBusy.value || refreshingAfterTask.value || previewBehind.value)
      return '文档正在同步预览，请稍后发送'
    if (saving.value) return '文档正在保存，请稍后发送'
    if (!documentId.value) return '请先打开或新建文档'
    if (!backendReady.value)
      return '后端能力未更新，请重启 npm run dev 后刷新页面'
    if (!modelConfigured.value) return '尚未配置模型，可先输入需求'
    return ''
  })
  const saveLabel = computed(() => {
    if (saving.value) return '正在保存'
    if (previewBehind.value)
      return `同步预览 · v${previewRevision.value} → v${revision.value}`
    if (dirty.value) return '有未保存更改'
    if (revision.value) return `已保存 · v${revision.value}`
    return '本地工作台'
  })

  async function fetchDocument(id: string, preserveConversation = false) {
    latestChangeLocation = undefined
    const meta = await api<DocumentMeta>(`/api/documents/${id}`)
    const [file, checked, history] = await Promise.all([
      fetch(`/api/documents/${id}/file?revision=${meta.revision}`),
      api<{ validation: Validation; sections: SectionInfo[] }>(
        `/api/documents/${id}/validate?revision=${meta.revision}`,
      ),
      api<{ messages: ChatLine[] }>(`/api/documents/${id}/chat`),
    ])
    if (!file.ok) throw new Error('无法读取当前版本')
    const loadedBytes = await file.arrayBuffer()
    documentId.value = meta.id
    revision.value = meta.revision
    previewRevision.value = meta.revision
    fileName.value = meta.name
    documentData.value = loadedBytes
    documentKey.value++
    documents.value = documents.value.some((d) => d.id === meta.id)
      ? documents.value.map((d) => (d.id === meta.id ? meta : d))
      : [meta, ...documents.value]
    dirty.value = false
    formatValidation.value = checked.validation
    sections.value = checked.sections || []
    if (!preserveConversation) messagesByDocument.value[id] = history.messages
    errorMessage.value = ''
  }

  async function refreshMetadataOnly(id: string) {
    const [meta, checked] = await Promise.all([
      api<DocumentMeta>(`/api/documents/${id}`),
      api<{ validation: Validation; sections: SectionInfo[] }>(
        `/api/documents/${id}/validate`,
      ),
    ])
    revision.value = meta.revision
    if (previewRevision.value !== meta.revision)
      await syncPreview(id, meta.revision)
    documents.value = documents.value.some((d) => d.id === meta.id)
      ? documents.value.map((d) => (d.id === meta.id ? meta : d))
      : [meta, ...documents.value]
    formatValidation.value = checked.validation
    sections.value = checked.sections || []
  }

  async function refreshList() {
    documents.value = (
      await api<{ documents: DocumentMeta[] }>('/api/documents')
    ).documents
  }

  async function createDocument(name = '未命名文档.docx') {
    if (!canLeaveDocument()) return
    documentBusy.value = true
    try {
      const meta = await api<DocumentMeta>(
        `/api/documents?name=${encodeURIComponent(name)}`,
        { method: 'POST' },
      )
      await refreshList()
      await fetchDocument(meta.id)
    } catch (error) {
      errorMessage.value =
        error instanceof Error ? error.message : '新建文档失败'
    } finally {
      documentBusy.value = false
    }
  }

  async function openLocalFile(event: Event) {
    const input = event.target as HTMLInputElement
    const file = input.files?.[0]
    input.value = ''
    if (file) await uploadDocument(file)
  }

  async function uploadDocument(file: File) {
    if (!canLeaveDocument()) return
    if (
      !/\.docx$/i.test(file.name) ||
      file.size > 20 * 1024 * 1024 ||
      file.size === 0
    ) {
      errorMessage.value = '请选择非空且不超过 20 MB 的 DOCX 文档'
      return
    }
    documentBusy.value = true
    try {
      errorMessage.value = ''
      const meta = await api<DocumentMeta>(
        `/api/documents?name=${encodeURIComponent(file.name)}`,
        {
          method: 'POST',
          headers: {
            'content-type':
              'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          },
          body: await file.arrayBuffer(),
        },
      )
      await refreshList()
      await fetchDocument(meta.id)
    } catch (error) {
      errorMessage.value =
        error instanceof Error ? error.message : '文档打开失败'
    } finally {
      documentBusy.value = false
    }
  }

  function markDirty() {
    if (documentId.value && !aiBusy.value && !refreshingAfterTask.value) {
      dirty.value = true
      editGeneration++
    }
  }

  function canLeaveDocument() {
    if (closingBlocked.value) return false
    return (
      !dirty.value || window.confirm('当前文档有未保存更改。放弃更改并继续？')
    )
  }

  function closeDocument() {
    if (!documentId.value || closingBlocked.value) return
    if (dirty.value) {
      closeDialog.value?.showModal()
      return
    }
    finishClose()
  }

  function finishClose() {
    if (closingBlocked.value) return
    closeDialog.value?.close()
    documentId.value = ''
    revision.value = 0
    previewRevision.value = 0
    dirty.value = false
    documentData.value = 'blank'
    documentKey.value++
    fileName.value = '未打开文档'
    sections.value = []
    formatValidation.value = null
    chatInput.value = ''
    activeTool.value = ''
    errorMessage.value = ''
    pageSettingsOpen.value = false
    recentOpen.value = false
  }

  async function saveAndClose() {
    await saveFile()
    if (!dirty.value && !errorMessage.value) finishClose()
  }

  async function discardChanges() {
    if (
      !dirty.value ||
      closingBlocked.value ||
      !window.confirm('放弃当前未保存的修改，恢复到最近保存的版本？')
    )
      return
    documentBusy.value = true
    try {
      await fetchDocument(documentId.value, true)
    } catch (error) {
      errorMessage.value =
        error instanceof Error ? error.message : '恢复文档失败，未关闭文档'
    } finally {
      documentBusy.value = false
    }
  }

  async function switchDocument(id: string) {
    if (id === documentId.value || !canLeaveDocument()) return
    documentBusy.value = true
    try {
      await fetchDocument(id)
    } catch (error) {
      errorMessage.value =
        error instanceof Error ? error.message : '文档切换失败'
    } finally {
      documentBusy.value = false
    }
  }

  function guardUnload(event: BeforeUnloadEvent) {
    if (dirty.value || saving.value) {
      event.preventDefault()
      event.returnValue = ''
    }
  }
  window.addEventListener('beforeunload', guardUnload)
  onUnmounted(() => window.removeEventListener('beforeunload', guardUnload))

  async function saveFile() {
    if (
      !documentId.value ||
      !editorRef.value ||
      saving.value ||
      previewBehind.value
    )
      return false
    const generation = editGeneration
    saving.value = true
    errorMessage.value = ''
    try {
      const bytes = await editorRef.value.save()
      if (!bytes) throw new Error('编辑器没有返回文档内容')
      const meta = await api<DocumentMeta>(
        `/api/documents/${documentId.value}/file`,
        {
          method: 'PUT',
          headers: {
            'content-type':
              'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'x-document-revision': String(revision.value),
          },
          body: bytes,
        },
      )
      revision.value = meta.revision
      previewRevision.value = meta.revision
      documents.value = documents.value.map((d) =>
        d.id === meta.id ? meta : d,
      )
      dirty.value = editGeneration !== generation
      formatValidation.value = meta.versions.at(-1)?.validation || null
      const checked = await api<{ sections: SectionInfo[] }>(
        `/api/documents/${documentId.value}/validate?revision=${meta.revision}`,
      )
      sections.value = checked.sections
      return !dirty.value
    } catch (error) {
      errorMessage.value =
        error instanceof Error ? error.message : '文档保存失败'
      return false
    } finally {
      saving.value = false
    }
  }

  async function exportFile() {
    if (!editorRef.value || previewBehind.value) return
    try {
      const buffer = await editorRef.value.save()
      if (!buffer) throw new Error('编辑器没有返回文档内容')
      const url = URL.createObjectURL(
        new Blob([buffer], {
          type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        }),
      )
      const link = document.createElement('a')
      link.href = url
      link.download = fileName.value
      link.click()
      window.setTimeout(() => URL.revokeObjectURL(url), 1000)
    } catch (error) {
      errorMessage.value =
        error instanceof Error ? error.message : 'DOCX 导出失败'
    }
  }

  async function restoreVersion(target: number) {
    if (!canLeaveDocument()) return
    documentBusy.value = true
    try {
      const meta = await api<DocumentMeta>(
        `/api/documents/${documentId.value}/restore`,
        {
          method: 'POST',
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ revision: revision.value, target }),
        },
      )
      documents.value = documents.value.map((d) =>
        d.id === meta.id ? meta : d,
      )
      await fetchDocument(meta.id)
    } catch (error) {
      errorMessage.value =
        error instanceof Error ? error.message : '版本恢复失败'
    } finally {
      documentBusy.value = false
    }
  }

  function twipsToMm(value: number) {
    return Math.round(((value * 25.4) / 1440) * 10) / 10
  }
  function mmToTwips(value: number) {
    return Math.round((value / 25.4) * 1440)
  }

  function openPageSettings() {
    if (!sections.value.length || dirty.value || closingBlocked.value) return
    sectionSelection.value = sections.value.length > 1 ? 'all' : '0'
    loadSectionSettings(0)
    pageSettingsOpen.value = true
  }

  watch(sectionSelection, (value) => {
    if (pageSettingsOpen.value)
      loadSectionSettings(value === 'all' ? 0 : Number(value))
  })

  function loadSectionSettings(index: number) {
    const section = sections.value[index]
    if (!section) return
    pageOrientation.value =
      section.page?.orientation === 'landscape' ? 'landscape' : 'portrait'
    pageWidthMm.value = twipsToMm(section.page?.widthTwips || 11906)
    pageHeightMm.value = twipsToMm(section.page?.heightTwips || 16838)
    pagePreset.value =
      Math.abs(pageWidthMm.value - 210) < 1 &&
      Math.abs(pageHeightMm.value - 297) < 1
        ? 'A4'
        : 'custom'
    pageMarginsMm.value = {
      top: twipsToMm(section.margins?.top || 0),
      right: twipsToMm(section.margins?.right || 0),
      bottom: twipsToMm(section.margins?.bottom || 0),
      left: twipsToMm(section.margins?.left || 0),
      gutter: twipsToMm(section.margins?.gutter || 0),
    }
  }

  function setPagePreset(value: string) {
    pagePreset.value = value
    const sizes: Record<string, [number, number]> = {
      A4: [210, 297],
      Letter: [215.9, 279.4],
      Legal: [215.9, 355.6],
    }
    if (sizes[value]) [pageWidthMm.value, pageHeightMm.value] = sizes[value]
  }

  async function savePageSettings() {
    if (
      !documentId.value ||
      !sections.value.length ||
      dirty.value ||
      closingBlocked.value
    )
      return
    documentBusy.value = true
    const sectionIndices =
      sectionSelection.value === 'all'
        ? sections.value.map((section) => section.index)
        : [Number(sectionSelection.value)]
    errorMessage.value = ''
    try {
      await api(`/api/documents/${documentId.value}/format`, {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({
          revision: revision.value,
          operations: [
            {
              scope: 'sections',
              sectionIndices,
              page: {
                widthTwips: mmToTwips(
                  pageOrientation.value === 'landscape'
                    ? Math.max(pageWidthMm.value, pageHeightMm.value)
                    : Math.min(pageWidthMm.value, pageHeightMm.value),
                ),
                heightTwips: mmToTwips(
                  pageOrientation.value === 'landscape'
                    ? Math.min(pageWidthMm.value, pageHeightMm.value)
                    : Math.max(pageWidthMm.value, pageHeightMm.value),
                ),
                orientation: pageOrientation.value,
              },
              margins: Object.fromEntries(
                Object.entries(pageMarginsMm.value).map(([key, value]) => [
                  key,
                  mmToTwips(value),
                ]),
              ),
            },
          ],
        }),
      })
      pageSettingsOpen.value = false
      await refreshList()
      await fetchDocument(documentId.value)
    } catch (error) {
      errorMessage.value =
        error instanceof Error ? error.message : '页面设置未能保存'
    } finally {
      documentBusy.value = false
    }
  }

  function onComposerEnter(event: KeyboardEvent) {
    if (event.isComposing) return
    event.preventDefault()
    void sendMessage()
  }

  async function sendMessage() {
    const prompt = chatInput.value.trim()
    if (!prompt || sendBlockedReason.value) return
    errorMessage.value = ''
    aiBusy.value = true
    followChanges.value = true
    changeLocation.value = ''
    latestChangeLocation = undefined
    if (dirty.value && !(await saveFile())) {
      aiBusy.value = false
      return
    }
    const targetDocId = documentId.value
    messagesByDocument.value[targetDocId] ||= []
    const lines = messagesByDocument.value[targetDocId]!
    lines.push({ role: 'user', text: prompt })
    const reply = reactive<ChatLine>({
      role: 'assistant',
      text: '',
      status: '正在等待模型响应',
      steps: [],
    })
    lines.push(reply)
    chatInput.value = ''
    activeTool.value = ''
    const controller = new AbortController()
    abortController.value = controller
    let needsFullSnapshot = false

    try {
      const response = await fetch(`/api/documents/${targetDocId}/chat`, {
        method: 'POST',
        signal: controller.signal,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ message: prompt, revision: revision.value, editorCapabilities: { nativeToc: true } }),
      })

      let completed = false
      const handleEvent = async (event: string, payload: any) => {
        if (event === 'editor_request') {
          const url = `/api/documents/${targetDocId}/editor-request?token=${encodeURIComponent(payload.token)}`
          try {
            if (targetDocId !== documentId.value || dirty.value || payload.revision !== revision.value) throw new Error('文档状态已变化，取消原生目录操作')
            const editor = editorRef.value?.getEditor()
            if (!editor || !('surface' in editor)) throw new Error('编辑器尚未就绪')
            const input = await fetch(url, { signal: controller.signal })
            if (!input.ok) throw new Error('目录请求已失效')
            const refresh = createDocumentRefresh(editor as DocxEditorInstance)
            const submission = await refresh.capture()
            try {
              const applied = await refresh.applyUpdate({ submission, sequence: payload.revision, bytes: await input.arrayBuffer() })
              if (!applied.ok) throw new Error('目录准备文档加载失败')
            } finally { refresh.finish(submission) }
            const currentEditor = editorRef.value?.getEditor()
            if (!currentEditor) throw new Error('刷新后编辑器尚未就绪')
            const selected = currentEditor.exec({ type: 'setSelection', anchor: { paraId: payload.paraId } })
            if (!selected.ok) throw new Error(`无法定位原生操作位置：${JSON.stringify(selected)}`)
            if (payload.mode === 'threeLineTable') {
              const tableSelection = { type: 'selectTableRegion' as const, region: 'table' as const }
              if (!currentEditor.can(tableSelection).ok) throw new Error('当前位置无法选择目标表格')
              if (!currentEditor.exec(tableSelection).ok) throw new Error('目标表格选择失败')
              const clear = { type: 'setTableBorders' as const, scope: 'none' as const, target: 'all' as const }
              if (!currentEditor.can(clear).ok || !currentEditor.exec(clear).ok) throw new Error('清除表格边框失败')
              const line = (scope: 'top' | 'bottom') => ({
                type: 'setTableBorders' as const,
                scope,
                spec: { style: 'single' as const, size: 8, color: { kind: 'hex' as const, value: '000000' } },
              })
              for (const scope of ['top', 'bottom'] as const) {
                const command = line(scope)
                if (!currentEditor.can(command).ok || !currentEditor.exec(command).ok) throw new Error(`设置表格${scope === 'top' ? '顶线' : '底线'}失败`)
              }
              const headerSelection = { type: 'selectTableRegion' as const, region: 'row' as const }
              if (!currentEditor.can(headerSelection).ok || !currentEditor.exec(headerSelection).ok) throw new Error('选择表头行失败')
              const headerLine = line('bottom')
              if (!currentEditor.can(headerLine).ok || !currentEditor.exec(headerLine).ok) throw new Error('设置表头下线失败')
            } else {
              const command = payload.mode === 'refresh'
                ? { type: 'refreshToc' as const, mode: 'entire' as const }
                : { type: 'insertToc' as const }
              if (!currentEditor.can(command).ok) throw new Error('当前位置无法执行原生目录命令')
              const result = currentEditor.exec(command)
              if (!result.ok) throw new Error('原生目录命令执行失败')
            }
            const bytes = await currentEditor.save()
            controller.signal.throwIfAborted()
            const accepted = await fetch(url, { method: 'PUT', body: bytes, signal: controller.signal })
            if (!accepted.ok) throw new Error('目录结果回传失败')
            needsFullSnapshot = true
            if (!completed) reply.status = '原生操作已完成，正在校验并保存'
          } catch (error) {
            needsFullSnapshot = true
            if (!controller.signal.aborted) {
              await fetch(url, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ error: error instanceof Error ? error.message : '目录执行失败' }) }).catch(() => {})
              await syncPreview(targetDocId, revision.value)
            }
          }
          return
        }
        if (event === 'status') {
          if (payload.message) reply.status = payload.message
          else if (payload.state === 'running') reply.status = '正在处理文档'
        }
        if (event === 'text') {
          reply.text += payload.delta
          reply.status = '正在生成回复'
        }
        if (event === 'tool') {
          const step = reply.steps!.find((item) => item.id === payload.id)
          if (step) step.state = payload.state
          else
            reply.steps!.push({
              id: payload.id,
              label: payload.label || payload.name,
              state: payload.state,
            })
          activeTool.value =
            payload.state === 'running' ? payload.label || payload.name : ''
          reply.status =
            payload.state === 'running'
              ? payload.label || payload.name
              : '正在等待模型继续处理'
        }
        if (event === 'plan' && Array.isArray(payload.steps)) {
          for (const planStep of payload.steps) {
            const id = `plan-${payload.planIndex ?? 0}-${planStep.index}`
            const existing = reply.steps!.find((item) => item.id === id)
            const mappedState =
              planStep.state === 'committed'
                ? 'complete'
                : planStep.state === 'failed'
                  ? 'error'
                  : planStep.state
            if (existing) existing.state = mappedState
            else
              reply.steps!.push({
                id,
                label: planStep.label,
                state: mappedState,
              })
          }
        }
        if (event === 'checklist' && Array.isArray(payload.items)) {
          const stateLabels: Record<string, string> = {
            pending: '待处理',
            awaitingVerification: '待核验',
            needsDecision: '需要确认',
            addressed: '已处理（非视觉验收）',
            blocked: '受阻',
            notApplicable: '不适用',
          }
          for (const item of payload.items) {
            const id = `requirement-${item.id}`
            const state = ['pending', 'awaitingVerification'].includes(
              item.state,
            )
              ? 'pending'
              : ['blocked', 'needsDecision'].includes(item.state)
                ? 'error'
                : 'complete'
            const label = `${item.label} · ${stateLabels[item.state] || item.state}${item.reason ? `：${item.reason}` : ''}${item.question ? `；需确认：${item.question}` : ''}`
            const existing = reply.steps!.find((s) => s.id === id)
            if (existing) {
              existing.state = state
              existing.label = label
            } else {
              reply.steps!.push({ id, label, state })
            }
          }
        }
        if (
          event === 'document' &&
          payload.id === targetDocId &&
          payload.revision > revision.value
        ) {
          revision.value = payload.revision
          const editor = editorRef.value?.getEditor()
          if (
            !needsFullSnapshot &&
            payload.replay?.kind === 'format-v1' &&
            payload.previousRevision === previewRevision.value &&
            editor
          ) {
            try {
              await applyFormatReplay(
                editor,
                payload.replay as FormatReplay,
                (done, total, _index, location) => {
                  locateChange(location)
                  if (!completed) reply.status = `正在原位更新格式 (${done}/${total})…`
                },
              )
              previewRevision.value = payload.revision
            } catch {
              needsFullSnapshot = true
              await syncPreview(targetDocId, revision.value)
              // A successful full sync establishes a new replay baseline.
              // Keep later ordinary formatting incremental in the same task.
              needsFullSnapshot = false
            }
          } else {
            needsFullSnapshot = true
            try {
              await syncPreview(targetDocId, revision.value)
              // Full synchronization completed; future format-v1 commits can
              // safely replay from this preview revision again.
              needsFullSnapshot = false
            } catch {
              errorMessage.value =
                '新版本已提交，但预览刷新失败；任务结束后将再次同步'
            }
          }
        }
        if (event === 'task_result' && payload.state !== 'complete') {
          const saved = (payload.committedRevisions || []).length
          const suffix = `${saved ? '本次修改已保存，可在修改历史中查看本次任务。' : '本轮未提交修改。'}${payload.guidance || ''}`
          if (suffix.trim()) {
            reply.text += `${LF}${LF}${suffix}`
          }
        }
        if (event === 'done') {
          completed = true
          reply.status =
            payload.state === 'cancelled'
              ? '已停止'
              : payload.state === 'partial'
                ? '部分完成 · 查看任务清单'
                : '对话结束'
        }
        if (event === 'error') {
          throw new Error(payload.message)
        }
      }
      await consumeEvents(response, createEditorEventHandler(handleEvent))
      if (!completed) throw new Error('连接提前结束，未收到任务完成确认')
    } catch (error) {
      if (controller.signal.aborted) {
        reply.status = '已停止'
      } else {
        reply.status = '未完成'
        errorMessage.value =
          error instanceof Error ? error.message : 'AI 请求失败'
        reply.text ||= errorMessage.value
      }
    } finally {
      refreshingAfterTask.value = true
      aiBusy.value = false
      if (!dirty.value && targetDocId) {
        const taskError = errorMessage.value
        try {
          if (needsFullSnapshot || previewRevision.value !== revision.value) {
            await syncPreview(targetDocId, revision.value)
          } else {
            await refreshMetadataOnly(targetDocId)
          }
          await refreshList()
          if (taskError) errorMessage.value = taskError
        } catch (error) {
          errorMessage.value =
            error instanceof Error ? error.message : '任务结束后同步失败'
        }
      }
      refreshingAfterTask.value = false
      activeTool.value = ''
      abortController.value = null
    }
  }

  async function saveModelSettings() {
    settingsSaving.value = true
    errorMessage.value = ''
    try {
      const settings = await api<{
        configured: boolean
        model: string
        baseUrl: string
      }>('/api/settings/model', {
        method: 'PUT',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({
          apiKey: modelApiKey.value,
          model: modelName.value,
          baseUrl: modelBaseUrl.value,
        }),
      })
      modelConfigured.value = settings.configured
      modelName.value = settings.model
      modelBaseUrl.value = settings.baseUrl
      modelApiKey.value = ''
      settingsOpen.value = false
    } catch (error) {
      errorMessage.value =
        error instanceof Error ? error.message : '模型设置保存失败'
    } finally {
      settingsSaving.value = false
    }
  }

  async function stopMessage() {
    try {
      await api(`/api/documents/${documentId.value}/cancel`, {
        method: 'POST',
      })
    } catch (error) {
      errorMessage.value =
        error instanceof Error ? error.message : '停止请求失败'
    }
  }

  function formatTime(value: string) {
    return new Date(value).toLocaleString('zh-CN', {
      month: 'numeric',
      day: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
    })
  }

  onMounted(async () => {
    try {
      const [health, settings] = await Promise.all([
        api<{ modelConfigured: boolean; capabilities?: string[] }>(
          '/api/health',
        ),
        api<{ configured: boolean; model: string; baseUrl: string }>(
          '/api/settings/model',
        ),
      ])
      backendReady.value = [
        'east-asian-fonts',
        'text-range-format',
        'pagination-flags',
        'mixed-text-ranges',
        'explicit-line-spacing',
        'latin-font-slot',
        'document-commit-events',
      ].every((capability) => health.capabilities?.includes(capability))
      if (!backendReady.value) {
        errorMessage.value =
          '当前后端尚未加载新版排版工具，请保存文档后重启 npm run dev，再刷新页面。'
      }
      modelConfigured.value = health.modelConfigured
      modelName.value = settings.model
      modelBaseUrl.value = settings.baseUrl
      await refreshList()
    } catch (error) {
      errorMessage.value =
        error instanceof Error ? error.message : '本地服务连接失败'
    }
  })

  const tocNotice = ref('')

  function runTocAction(action: 'insert' | 'entire' | 'pageNumbers') {
    if (!documentId.value || closingBlocked.value) return
    const editor = editorRef.value?.getEditor()
    if (!editor) {
      errorMessage.value = '文档尚未加载完成，请稍后再试。'
      return
    }
    tocNotice.value = ''
    errorMessage.value = ''
    try {
      if (action === 'insert' && !editor.getOutline().some(item => item.level < 3)) {
        errorMessage.value = '请先为章节设置标题 1～3 样式，再插入目录。左侧规则识别结果不会自动写入标题样式。'
        return
      }
      const command = action === 'insert'
        ? { type: 'insertToc' as const }
        : { type: 'refreshToc' as const, mode: action }
      if (!editor.can(command).ok) {
        errorMessage.value = action === 'insert'
          ? '当前位置不能插入目录，请先将光标放在正文中的目录插入位置。'
          : '请先点击文档中的正式目录，再选择更新。手工输入的目录文字不能作为目录域更新。'
        return
      }
      const result = editor.exec(command)
      if (!result.ok) throw new Error('编辑器未能完成目录操作，请检查当前位置后重试。')
      tocNotice.value = action === 'insert'
        ? '已插入目录（标题 1～3）。请检查结果并保存文档。'
        : '目录更新命令已执行，请检查标题、页码和分页结果后保存。'
    } catch (error) {
      errorMessage.value = error instanceof Error ? error.message : '目录操作失败'
    }
  }

  return {
    tocNotice,
    runTocAction,
    navigationReady,
    editorRef,
    inputRef,
    documentData,
    documentKey,
    documents,
    documentId,
    fileName,
    revision,
    dirty,
    saving,
    aiBusy,
    documentBusy,
    refreshingAfterTask,
    modelConfigured,
    assistantOpen,
    rightTab,
    recentOpen,
    settingsOpen,
    pageSettingsOpen,
    sections,
    sectionSelection,
    pagePreset,
    pageWidthMm,
    pageHeightMm,
    pageOrientation,
    pageMarginsMm,
    modelName,
    modelBaseUrl,
    modelApiKey,
    settingsSaving,
    errorMessage,
    closeDialog,
    closingBlocked,
    chatInput,
    activeTool,
    followChanges,
    changeLocation,
    pauseFollowing,
    resumeFollowing,
    formatValidation,
    versions,
    historyGroups,
    pendingWork,
    savePartialResult,
    chatLines,
    sendBlockedReason,
    saveLabel,
    createDocument,
    openLocalFile,
    uploadDocument,
    closeDocument,
    finishClose,
    saveAndClose,
    discardChanges,
    switchDocument,
    saveFile,
    exportFile,
    restoreVersion,
    openPageSettings,
    setPagePreset,
    savePageSettings,
    onComposerEnter,
    sendMessage,
    saveModelSettings,
    stopMessage,
    formatTime,
  }
}
