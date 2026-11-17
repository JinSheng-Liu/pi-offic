<script setup lang="ts">
import { ref, watch } from 'vue'
import { ScanSearch, TriangleAlert } from '@lucide/vue'
import { recognizeHeading } from '../outline'
import { reviewLayout, type LayoutFinding } from '../layoutReview'

const props = defineProps<{
  editor: any
  blocked?: boolean
}>()

const checked = ref(false)
const pageCount = ref(0)
const findings = ref<LayoutFinding[]>([])
const fontMeasurement = ref('')
const fontSubstitutions = ref<string[]>([])

watch(
  () => props.editor,
  (editor, _, onCleanup) => {
    checked.value = false
    findings.value = []
    if (!editor || typeof editor.on !== 'function') return
    const off = editor.on('change', () => {
      checked.value = false
      findings.value = []
    })
    onCleanup(() => off?.())
  },
  { immediate: true },
)

function runReview() {
  const editor = props.editor
  if (!editor) return
  const layout = editor.surface?.layout?.()
  const pages = layout?.pages || []
  pageCount.value = pages.length
  fontMeasurement.value = layout?.fontMeasurement || 'fixed'
  fontSubstitutions.value = Array.isArray(layout?.fontSubstitutions) ? layout.fontSubstitutions : []

  const headingIds = new Set<string>()
  const native = typeof editor.getOutline === 'function' ? editor.getOutline() : []
  for (const item of native) {
    if (item?.blockId) headingIds.add(item.blockId)
  }
  for (const page of pages) {
    for (const frag of page.fragments || []) {
      if (frag.kind !== 'paragraph' || !frag.paragraphId) continue
      const text = (frag.lines || []).flatMap((l: any) => (l.spans || []).map((s: any) => s.text || '')).join('').trim()
      if (recognizeHeading(text) !== null) {
        headingIds.add(frag.paragraphId)
      }
    }
  }

  findings.value = reviewLayout(pages, headingIds)
  checked.value = true
}

function jumpTo(paragraphId: string) {
  props.editor?.scrollToBlock?.(paragraphId)
}
</script>

<template>
  <section class="layout-review" aria-label="页面排版检查">
    <div class="layout-review-header">
      <div>
        <strong>当前页面审阅</strong>
        <p>只读检测孤立标题、短尾行与页底留白，不修改文档。</p>
      </div>
      <button class="button button-quiet" type="button" :disabled="blocked || !editor" @click="runReview">
        <ScanSearch :size="14" />
        <span>检查当前页面</span>
      </button>
    </div>

    <div v-if="checked" class="layout-review-results" role="status">
      <p class="layout-meta">
        共 {{ pageCount }} 页 · 字宽测量：{{ fontMeasurement === 'fixed' ? '估算字宽（非 Word 精确渲染）' : fontMeasurement }}
        <template v-if="fontSubstitutions.length"> · 替代字体：{{ fontSubstitutions.join('、') }}</template>
      </p>
      <p v-if="!findings.length" class="layout-clean">未发现跨页孤立标题或异常短尾行提示。</p>
      <ul v-else class="layout-finding-list">
        <li v-for="(item, idx) in findings" :key="`${item.paragraphId}-${idx}`">
          <TriangleAlert :size="13" />
          <button type="button" class="layout-finding-link" @click="jumpTo(item.paragraphId)">
            <span>第 {{ item.page }} 页：{{ item.message }}</span>
          </button>
        </li>
      </ul>
    </div>
  </section>
</template>
