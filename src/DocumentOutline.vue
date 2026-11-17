<script setup lang="ts">
import { computed, nextTick, onMounted, onUnmounted, ref, watch } from 'vue'
import { useDocxEditor, useEditorEvent, useEditorState } from '@docx-editor.dev/vue'
import { recognizeHeading, visibleOutline, type OutlineItem } from './outline'

defineProps<{ fileName: string }>()
const editor = useDocxEditor()
const loading = useEditorState(state => state.isLoading)
const items = ref<OutlineItem[]>([])
const source = ref<'native' | 'rules'>('native')
const collapsed = ref(new Set<string>())
const activeId = ref('')
const notice = ref('')
const visible = computed(() => visibleOutline(items.value, collapsed.value))
const minimumLevel = computed(() => Math.min(...items.value.map(item => item.level), 0))
const parents = computed(() => new Set(items.value.filter((item, index) => (items.value[index + 1]?.level ?? -1) > item.level).map(item => item.blockId)))
let frame = 0
let refreshTimer: ReturnType<typeof setTimeout> | undefined
let mounted = false
const offsets = new Map<string, number>()
let scrollElement: HTMLElement | null = null
let headingElements: HTMLElement[] = []
let headingIndexDirty = true
let contentObserver: MutationObserver | undefined

function invalidateHeadingIndex() {
  headingIndexDirty = true
  scheduleActive()
}

function bindEditorScroll() {
  if (typeof document === 'undefined') return
  const next = document.querySelector<HTMLElement>('.editor-frame [data-testid="docx-editor-scroll"]')
  if (next === scrollElement) return
  contentObserver?.disconnect()
  scrollElement?.removeEventListener('scroll', scheduleActive)
  scrollElement = next
  scrollElement?.addEventListener('scroll', scheduleActive, { passive: true })
  headingIndexDirty = true
  offsets.clear()
  if (scrollElement && typeof MutationObserver !== 'undefined') {
    // Virtual pages can mount/unmount while scrolling. Rebuild only when their
    // DOM changes; a scroll event alone never walks all document paragraphs.
    contentObserver = new MutationObserver(invalidateHeadingIndex)
    contentObserver.observe(scrollElement, { childList: true, subtree: true })
  }
}

function updateActive() {
  frame = 0
  const scroll = scrollElement
  if (!scroll) return
  const top = scroll.getBoundingClientRect().top
  if (headingIndexDirty) {
    const ids = new Set(items.value.map(item => item.blockId))
    headingElements = Array.from(scroll.querySelectorAll<HTMLElement>('[data-paragraph-id]'))
      .filter(element => ids.has(element.dataset.paragraphId!))
    headingIndexDirty = false
  }
  const measured = new Set<string>()
  for (const element of headingElements) {
    const id = element.dataset.paragraphId!
    if (measured.has(id)) continue
    measured.add(id)
    offsets.set(id, element.getBoundingClientRect().top - top + scroll.scrollTop)
  }
  const candidates = items.value.filter(item => offsets.has(item.blockId))
  const current = candidates.filter(item => offsets.get(item.blockId)! <= scroll.scrollTop + 100).at(-1)
  activeId.value = current?.blockId || candidates[0]?.blockId || ''
}

function scheduleActive() {
  if (mounted && !frame) frame = requestAnimationFrame(updateActive)
}

function refreshOutline() {
  bindEditorScroll()
  if (!editor.value || loading.value) return
  notice.value = ''
  const native = editor.value.getOutline()
  if (native.length) {
    source.value = 'native'
    items.value = native.map(item => ({ ...item }))
  } else {
    source.value = 'rules'
    const paragraphs = editor.value.query({ type: 'paragraphs' })
    const candidates = paragraphs.flatMap((paragraph, index) => {
      const level = recognizeHeading(paragraph.text)
      return level === null ? [] : [{ text: paragraph.text.trim(), level, index, start: paragraph.text.length - paragraph.text.trimStart().length }]
    })
    const matches = candidates.length ? editor.value.findMatches(candidates.map(item => item.text), { matchCase: true }) : []
    items.value = candidates.flatMap((item, index) => {
      const match = matches[index]?.find(hit => (!hit.scope || hit.scope.kind === 'body') && hit.paragraphIndex === item.index && hit.start === item.start)
      return match ? [{ text: item.text, level: item.level, blockId: match.blockId }] : []
    })
  }
  offsets.clear()
  headingIndexDirty = true
  const ids = new Set(items.value.map(item => item.blockId))
  collapsed.value = new Set([...collapsed.value].filter(id => ids.has(id)))
  void nextTick(scheduleActive)
}

function scheduleRefresh() {
  clearTimeout(refreshTimer)
  refreshTimer = setTimeout(refreshOutline, 120)
}

function toggle(id: string) {
  const next = new Set(collapsed.value)
  if (next.has(id)) next.delete(id)
  else next.add(id)
  collapsed.value = next
}

function navigate(item: OutlineItem) {
  if (loading.value) return
  if (editor.value?.scrollToBlock(item.blockId)) {
    activeId.value = item.blockId
    notice.value = ''
  } else {
    notice.value = '文档位置已变化，正在更新大纲，请重新点击。'
    scheduleRefresh()
  }
}

watch([editor, loading], refreshOutline, { immediate: true })

useEditorEvent('change', scheduleRefresh)
useEditorEvent('selectionChange', scheduleActive)

onMounted(() => {
  mounted = true
  bindEditorScroll()
  window.addEventListener('resize', scheduleActive)
  refreshOutline()
})

onUnmounted(() => {
  mounted = false
  clearTimeout(refreshTimer)
  cancelAnimationFrame(frame)
  scrollElement?.removeEventListener('scroll', scheduleActive)
  contentObserver?.disconnect()
  headingElements = []
  scrollElement = null
  window.removeEventListener('resize', scheduleActive)
})
</script>

<template>
  <nav class="document-outline" aria-label="文档大纲">
    <header>
      <h2>文档大纲</h2>
      <p class="outline-file" :title="fileName">{{ fileName }}</p>
    </header>
    <p v-if="items.length" class="outline-source">
      {{ source === 'native' ? '文档原生标题' : '智能识别 · 规则识别，未修改文档' }}
    </p>
    <p v-if="loading && !items.length" class="outline-empty" role="status">正在读取文档结构…</p>
    <p v-else-if="!items.length" class="outline-empty">
      暂无可识别的标题。设置标题样式后会自动显示；普通正文不会按字体大小或页数生成目录。
    </p>
    <ul v-else class="outline-list">
      <li
        v-for="item in visible"
        :key="item.blockId"
        :style="{ paddingLeft: `${Math.max(0, item.level - minimumLevel) * 14}px` }"
      >
        <button
          v-if="parents.has(item.blockId)"
          class="outline-toggle"
          type="button"
          :aria-label="`${collapsed.has(item.blockId) ? '展开' : '折叠'}${item.text}`"
          :aria-expanded="!collapsed.has(item.blockId)"
          @click="toggle(item.blockId)"
        >
          {{ collapsed.has(item.blockId) ? '▸' : '▾' }}
        </button>
        <span v-else class="outline-toggle-space" />
        <button
          class="outline-link"
          type="button"
          :disabled="loading"
          :aria-current="activeId === item.blockId ? 'location' : undefined"
          :title="item.text"
          @mousedown.prevent
          @click="navigate(item)"
        >
          {{ item.text }}
        </button>
      </li>
    </ul>
    <p v-if="notice" class="outline-empty" role="status">{{ notice }}</p>
  </nav>
</template>

<style scoped>
.document-outline {
  height: 100%;
  display: flex;
  flex-direction: column;
  color: var(--studio-text);
  background: var(--studio-surface);
}
header {
  padding: 18px 16px 12px;
  border-bottom: 1px solid var(--studio-border);
}
h2 {
  margin: 0 0 6px;
  font-size: 13px;
  font-weight: 700;
  letter-spacing: 0.02em;
  text-transform: uppercase;
  color: var(--studio-text);
}
.outline-file {
  margin: 0;
  font-size: 12px;
  color: var(--studio-muted);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.outline-source {
  margin: 10px 16px 4px;
  padding: 3px 8px;
  font-size: 11px;
  color: var(--studio-muted);
  background: var(--studio-surface);
  border: 1px solid var(--studio-border);
  border-radius: 6px;
  align-self: flex-start;
}
.outline-list {
  flex: 1;
  overflow: auto;
  list-style: none;
  padding: 8px;
  margin: 0;
}
.outline-list li {
  display: flex;
  align-items: flex-start;
  margin: 2px 0;
}
.outline-link {
  flex: 1;
  min-width: 0;
  text-align: left;
  border: 0;
  border-left: 2px solid transparent;
  background: transparent;
  border-radius: 0 8px 8px 0;
  padding: 6px 8px;
  font: inherit;
  font-size: 12.5px;
  line-height: 1.5;
  cursor: pointer;
  overflow-wrap: anywhere;
  color: var(--studio-text);
  transition: background 0.12s;
}
.outline-link:hover {
  background: var(--studio-accent-soft);
}
.outline-link[aria-current] {
  color: var(--studio-accent);
  background: var(--studio-accent-soft);
  border-left-color: var(--studio-accent);
  font-weight: 600;
}
.outline-toggle,
.outline-toggle-space {
  width: 20px;
  flex: 0 0 20px;
  padding: 7px 0;
  border: 0;
  background: transparent;
  color: var(--studio-muted);
  cursor: pointer;
}
.outline-empty {
  padding: 14px 16px;
  margin: 0;
  font-size: 12px;
  color: var(--studio-muted);
  line-height: 1.7;
}
</style>
