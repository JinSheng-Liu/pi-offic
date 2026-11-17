<script setup lang="ts">
import { computed, ref } from 'vue'
import {
  ArrowRight,
  CloudUpload,
  FileText,
  History,
  Moon,
  Sun,
  LoaderCircle,
  Settings,
  X,
} from '@lucide/vue'
import type { DocumentMeta } from '../types'
import StudioBrand from './StudioBrand.vue'

const props = defineProps<{
  documents: DocumentMeta[]
  busy: boolean
  error: string
  theme: 'light' | 'dark'
}>()

const emit = defineEmits<{
  toggleTheme: []
  upload: []
  drop: [file: File]
  open: [id: string]
  recent: []
  settings: []
  dismissError: []
}>()

const dragging = ref(false)
const recentDocuments = computed(() => props.documents.slice(0, 4))

function receiveDrop(event: DragEvent) {
  dragging.value = false
  if (props.busy) return

  const file = event.dataTransfer?.files[0]
  if (file) emit('drop', file)
}

function formatDate(value: string) {
  return new Date(value).toLocaleString('zh-CN', {
    month: 'numeric',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  })
}
</script>

<template>
  <div class="studio-home" :aria-busy="busy">
    <header class="home-header">
      <a
        class="home-brand"
        href="#"
        aria-label="文序 DOCX 排版工作台首页"
        @click.prevent
      >
        <StudioBrand />
      </a>

      <div class="home-actions">
        <span class="home-local-status">
          <span class="home-status-mark" aria-hidden="true" />
          本地工作区
        </span>
        <span class="home-header-divider" aria-hidden="true" />
        <button
          class="icon-button home-theme-button"
          type="button"
          :aria-label="theme === 'light' ? '切换深色界面' : '切换浅色界面'"
          :title="theme === 'light' ? '切换深色界面' : '切换浅色界面'"
          @click="emit('toggleTheme')"
        >
          <Moon v-if="theme === 'light'" :size="17" />
          <Sun v-else :size="17" />
        </button>
        <button
          class="button button-quiet home-settings"
          type="button"
          aria-label="模型设置"
          title="模型设置"
          :disabled="busy"
          @click="emit('settings')"
        >
          <Settings :size="16" />
          <span>设置</span>
        </button>
      </div>
    </header>

    <main class="home-content">
      <div v-if="error" class="home-error" role="alert">
        <span>{{ error }}</span>
        <button
          type="button"
          aria-label="关闭错误提示"
          @click="emit('dismissError')"
        >
          <X :size="16" />
        </button>
      </div>

      <section class="home-start" aria-labelledby="home-title">
        <div class="home-intro">
          <span class="home-eyebrow">文档排版工作台</span>
          <h1 id="home-title">从打开文档开始，<br />把格式理清楚。</h1>
          <p>
            在同一工作区检查文档结构、调整格式并查看修改结果。你可以自己编辑，也可以描述需求交给助手处理。
          </p>
          <div class="home-capabilities" aria-label="工作区包含">
            <span>文档结构</span>
            <i aria-hidden="true" />
            <span>格式调整</span>
            <i aria-hidden="true" />
            <span>结果检查</span>
          </div>
        </div>

        <section class="home-open-panel" aria-labelledby="home-open-title">
          <div
            class="home-upload"
            :class="{ dragging, 'is-busy': busy }"
            @dragover.prevent="dragging = !busy"
            @dragleave.self="dragging = false"
            @drop.prevent="receiveDrop"
          >
            <div class="home-upload-topline">
              <span>新建排版任务</span>
              <span class="home-file-type">DOCX</span>
            </div>

            <div class="home-upload-body">
              <div class="home-upload-icon-wrap" aria-hidden="true">
                <LoaderCircle
                  v-if="busy"
                  class="spin"
                  :size="25"
                  :stroke-width="1.8"
                />
                <CloudUpload v-else :size="25" :stroke-width="1.8" />
              </div>
              <div class="home-upload-copy">
                <h2 id="home-open-title">
                  {{
                    busy
                      ? '正在准备文档'
                      : dragging
                        ? '松开鼠标，打开文档'
                        : '打开一份 Word 文档'
                  }}
                </h2>
                <p>
                  {{
                    busy
                      ? '文档准备好后会进入编辑工作区。'
                      : '拖放 DOCX 文件到这里，或从电脑中选择。'
                  }}
                </p>
              </div>
              <button
                class="home-upload-button"
                type="button"
                :disabled="busy"
                @click="emit('upload')"
              >
                <LoaderCircle v-if="busy" class="spin" :size="16" />
                <CloudUpload v-else :size="16" />
                {{ busy ? '正在打开' : '选择文档' }}
                <ArrowRight v-if="!busy" :size="15" />
              </button>
            </div>

            <p class="home-local-note">
              <span class="home-lock-mark" aria-hidden="true" />
              文件在本机工作区打开；原文件不会被覆盖。
            </p>
          </div>
        </section>
      </section>

      <section class="home-recent-panel" aria-labelledby="home-recent-title">
        <header class="home-section-heading">
          <div>
            <span class="home-section-kicker">继续处理</span>
            <h2 id="home-recent-title">最近文档</h2>
          </div>
          <button
            v-if="documents.length"
            class="home-view-all"
            type="button"
            :disabled="busy"
            @click="emit('recent')"
          >
            查看全部 <ArrowRight :size="15" />
          </button>
        </header>

        <div v-if="recentDocuments.length" class="home-document-list">
          <button
            v-for="document in recentDocuments"
            :key="document.id"
            class="home-document-row"
            type="button"
            :disabled="busy"
            @click="emit('open', document.id)"
          >
            <span class="home-file-mark" aria-hidden="true">
              <FileText :size="18" />
            </span>
            <span class="home-document-copy">
              <strong :title="document.name">{{ document.name }}</strong>
              <time :datetime="document.updatedAt">
                最近打开 {{ formatDate(document.updatedAt) }}
              </time>
            </span>
            <span class="home-document-action" aria-hidden="true">
              打开 <ArrowRight :size="14" />
            </span>
          </button>
        </div>

        <div v-else class="home-recent-empty">
          <span class="home-empty-mark" aria-hidden="true">
            <History :size="18" />
          </span>
          <div>
            <strong>这里会保留你最近打开的文档</strong>
            <p>打开一份 DOCX 后，就能从这里继续上次的工作。</p>
          </div>
        </div>
      </section>

      <footer class="home-footer">
        <span>文序 <i>·</i> DOCX 排版工作台</span>
        <span>适用于 DOCX 文档</span>
      </footer>
    </main>
  </div>
</template>

<style scoped>
.studio-home {
  min-height: 100dvh;
  color: var(--studio-text);
  background:
    linear-gradient(180deg, var(--studio-surface) 0, var(--studio-surface) 72px, transparent 72px),
    var(--studio-canvas);
}

.home-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  min-height: 72px;
  padding: 0 max(32px, calc((100vw - 1280px) / 2));
  border-bottom: 1px solid var(--studio-border);
  background: var(--studio-surface);
}

.home-brand {
  display: inline-flex;
  color: var(--studio-text);
  text-decoration: none;
}

.home-actions {
  display: flex;
  align-items: center;
  gap: 12px;
}

.home-local-status {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  color: var(--studio-muted);
  font-size: 12px;
}

.home-status-mark {
  width: 7px;
  height: 7px;
  border-radius: 50%;
  background: var(--studio-success);
  box-shadow: 0 0 0 3px var(--studio-success-soft);
}

.home-header-divider {
  width: 1px;
  height: 22px;
  margin: 0 2px;
  background: var(--studio-border);
}

.home-theme-button {
  color: var(--studio-muted);
}

.home-settings {
  display: inline-flex;
  align-items: center;
  gap: 7px;
  min-height: 36px;
  padding: 0 11px;
  border: 1px solid var(--studio-border);
  border-radius: 7px;
  color: var(--studio-text);
  background: var(--studio-surface);
  font-size: 12px;
}

.home-content {
  width: min(100% - 64px, 1180px);
  margin: 0 auto;
  padding: 64px 0 0;
}

.home-start {
  display: grid;
  grid-template-columns: minmax(0, 0.92fr) minmax(430px, 1.08fr);
  align-items: center;
  gap: clamp(44px, 7vw, 104px);
  padding: 4px 0 58px;
}

.home-intro {
  max-width: 480px;
}

.home-eyebrow,
.home-section-kicker {
  color: var(--studio-accent);
  font-size: 12px;
  font-weight: 650;
  letter-spacing: 0.04em;
}

.home-intro h1 {
  margin: 20px 0 16px;
  color: var(--studio-text);
  font-size: clamp(34px, 4.2vw, 50px);
  font-weight: 650;
  letter-spacing: -0.055em;
  line-height: 1.28;
}

.home-intro > p {
  max-width: 390px;
  margin: 0;
  color: var(--studio-muted);
  font-size: 14px;
  line-height: 1.9;
}

.home-capabilities {
  display: flex;
  align-items: center;
  flex-wrap: wrap;
  gap: 11px;
  margin-top: 27px;
  color: var(--studio-muted);
  font-size: 11px;
}

.home-capabilities i {
  width: 3px;
  height: 3px;
  border-radius: 50%;
  background: var(--studio-border-strong);
}

.home-open-panel {
  min-width: 0;
}

.home-upload {
  padding: 19px 21px 16px;
  border: 1px solid var(--studio-border);
  border-radius: 12px;
  background: var(--studio-surface);
  box-shadow: 0 16px 40px color-mix(in srgb, var(--studio-text) 5%, transparent);
  transition:
    border-color 160ms ease,
    box-shadow 160ms ease,
    transform 160ms ease;
}

.home-upload.dragging {
  border-color: var(--studio-accent);
  background: var(--studio-accent-soft);
  box-shadow: 0 0 0 4px var(--studio-focus-ring);
  transform: translateY(-2px);
}

.home-upload.dragging * {
  pointer-events: none;
}

.home-upload.is-busy {
  opacity: 0.76;
}

.home-upload-topline {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding-bottom: 16px;
  border-bottom: 1px solid var(--studio-border);
  color: var(--studio-muted);
  font-size: 11px;
  font-weight: 600;
}

.home-file-type {
  padding: 4px 7px;
  border-radius: 4px;
  color: var(--studio-accent);
  background: var(--studio-accent-soft);
  font-size: 9px;
  letter-spacing: 0.08em;
}

.home-upload-body {
  display: grid;
  grid-template-columns: 46px minmax(0, 1fr) auto;
  align-items: center;
  gap: 15px;
  min-height: 150px;
  padding: 22px 0;
}

.home-upload-icon-wrap {
  display: grid;
  place-items: center;
  width: 44px;
  height: 48px;
  border: 1px solid var(--studio-border);
  border-radius: 8px;
  color: var(--studio-accent);
  background: var(--studio-canvas);
}

.home-upload-copy {
  min-width: 0;
}

.home-upload-copy h2 {
  margin: 0;
  color: var(--studio-text);
  font-size: 15px;
  font-weight: 620;
  line-height: 1.5;
}

.home-upload-copy p {
  margin: 6px 0 0;
  color: var(--studio-muted);
  font-size: 11px;
  line-height: 1.6;
}

.home-upload-button {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 8px;
  min-height: 40px;
  padding: 0 13px;
  border: 0;
  border-radius: 7px;
  color: var(--studio-on-accent);
  background: var(--studio-accent);
  font-size: 12px;
  font-weight: 600;
  white-space: nowrap;
  transition:
    background-color 140ms ease,
    transform 140ms ease;
}

.home-upload-button:hover:not(:disabled) {
  background: var(--studio-accent-hover);
  transform: translateY(-1px);
}

.home-upload-button:active:not(:disabled) {
  transform: translateY(0) scale(0.98);
}

.home-upload-button:disabled {
  opacity: 0.75;
}

.home-local-note {
  display: flex;
  align-items: center;
  gap: 8px;
  margin: 0;
  padding-top: 13px;
  border-top: 1px solid var(--studio-border);
  color: var(--studio-muted);
  font-size: 10px;
}

.home-lock-mark {
  width: 7px;
  height: 7px;
  border: 1px solid var(--studio-success);
  border-radius: 2px;
}

.home-recent-panel {
  padding-top: 25px;
  border-top: 1px solid var(--studio-border);
}

.home-section-heading {
  display: flex;
  align-items: flex-end;
  justify-content: space-between;
  gap: 20px;
  margin-bottom: 11px;
}

.home-section-kicker {
  font-size: 10px;
}

.home-section-heading h2 {
  margin: 4px 0 0;
  color: var(--studio-text);
  font-size: 18px;
  font-weight: 630;
}

.home-view-all {
  display: inline-flex;
  align-items: center;
  gap: 5px;
  min-height: 32px;
  padding: 0 2px 0 8px;
  border: 0;
  color: var(--studio-accent);
  background: transparent;
  font-size: 11px;
}

.home-document-list {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  column-gap: 24px;
}

.home-document-row {
  display: grid;
  grid-template-columns: 38px minmax(0, 1fr) auto;
  align-items: center;
  gap: 12px;
  min-width: 0;
  min-height: 67px;
  padding: 9px 8px;
  border: 0;
  border-bottom: 1px solid var(--studio-border);
  color: var(--studio-text);
  background: transparent;
  text-align: left;
  transition: background-color 140ms ease;
}

.home-document-row:hover:not(:disabled),
.home-document-row:focus-visible {
  background: var(--studio-subtle);
}

.home-document-row:disabled {
  opacity: 0.62;
}

.home-file-mark {
  display: grid;
  place-items: center;
  width: 34px;
  height: 38px;
  border: 1px solid var(--studio-border);
  border-radius: 6px;
  color: var(--studio-accent);
  background: var(--studio-surface);
}

.home-document-copy {
  display: flex;
  flex-direction: column;
  min-width: 0;
  gap: 5px;
}

.home-document-copy strong {
  overflow: hidden;
  color: var(--studio-text);
  font-size: 12px;
  font-weight: 570;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.home-document-copy time {
  color: var(--studio-muted);
  font-size: 10px;
}

.home-document-action {
  display: inline-flex;
  align-items: center;
  gap: 4px;
  color: var(--studio-muted);
  font-size: 10px;
  opacity: 0;
  transition: opacity 140ms ease;
}

.home-document-row:hover .home-document-action,
.home-document-row:focus-visible .home-document-action {
  opacity: 1;
}

.home-recent-empty {
  display: flex;
  align-items: center;
  gap: 13px;
  min-height: 80px;
  padding: 13px 0;
  border-bottom: 1px solid var(--studio-border);
}

.home-empty-mark {
  display: grid;
  place-items: center;
  width: 36px;
  height: 36px;
  flex: 0 0 36px;
  border: 1px solid var(--studio-border);
  border-radius: 7px;
  color: var(--studio-muted);
  background: var(--studio-surface);
}

.home-recent-empty strong {
  color: var(--studio-text);
  font-size: 12px;
  font-weight: 570;
}

.home-recent-empty p {
  margin: 5px 0 0;
  color: var(--studio-muted);
  font-size: 11px;
}

.home-footer {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 16px;
  margin-top: 37px;
  padding: 14px 0 18px;
  border-top: 1px solid var(--studio-border);
  color: var(--studio-muted);
  font-size: 10px;
}

.home-footer i {
  padding: 0 3px;
  color: var(--studio-accent);
  font-style: normal;
}

.home-error {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding: 12px 14px;
  margin-bottom: 24px;
  border: 1px solid var(--studio-danger-border);
  border-radius: 8px;
  color: var(--studio-danger);
  background: var(--studio-danger-soft);
  font-size: 12px;
}

.home-error button {
  display: grid;
  place-items: center;
  flex: 0 0 auto;
  padding: 4px;
  border: 0;
  color: inherit;
  background: transparent;
}

@media (max-width: 900px) {
  .home-header {
    padding-inline: 24px;
  }

  .home-content {
    width: min(100% - 48px, 680px);
    padding-top: 48px;
  }

  .home-start {
    grid-template-columns: 1fr;
    gap: 30px;
    padding-bottom: 42px;
  }

  .home-intro {
    max-width: 560px;
  }

  .home-intro h1 {
    margin-top: 13px;
  }

  .home-capabilities {
    margin-top: 17px;
  }
}

@media (max-width: 560px) {
  .home-header {
    min-height: 62px;
    padding-inline: 16px;
  }

  .home-actions {
    gap: 7px;
  }

  .home-local-status,
  .home-header-divider {
    display: none;
  }

  .home-settings {
    width: 34px;
    min-height: 34px;
    justify-content: center;
    padding: 0;
  }

  .home-settings span {
    display: none;
  }

  .home-content {
    width: min(100% - 32px, 440px);
    padding-top: 35px;
  }

  .home-start {
    gap: 25px;
    padding-bottom: 34px;
  }

  .home-intro h1 {
    margin: 12px 0 11px;
    font-size: clamp(29px, 8vw, 36px);
    letter-spacing: -0.045em;
  }

  .home-intro > p {
    font-size: 12px;
    line-height: 1.8;
  }

  .home-capabilities {
    gap: 8px;
    margin-top: 15px;
    font-size: 10px;
  }

  .home-upload {
    padding: 15px 14px 13px;
    border-radius: 9px;
  }

  .home-upload-body {
    grid-template-columns: 40px minmax(0, 1fr);
    gap: 12px;
    min-height: 0;
    padding: 20px 0;
  }

  .home-upload-icon-wrap {
    width: 40px;
    height: 42px;
  }

  .home-upload-copy h2 {
    font-size: 13px;
  }

  .home-upload-copy p {
    font-size: 10px;
  }

  .home-upload-button {
    grid-column: 1 / -1;
    width: 100%;
    min-height: 40px;
  }

  .home-local-note {
    align-items: flex-start;
    font-size: 9px;
    line-height: 1.5;
  }

  .home-lock-mark {
    flex: 0 0 7px;
    margin-top: 2px;
  }

  .home-recent-panel {
    padding-top: 20px;
  }

  .home-document-list {
    grid-template-columns: 1fr;
  }

  .home-document-row {
    min-height: 62px;
  }

  .home-document-action {
    display: none;
  }

  .home-footer {
    margin-top: 28px;
    font-size: 9px;
  }
}

@media (prefers-reduced-motion: reduce) {
  .home-upload,
  .home-upload-button,
  .home-document-action,
  .home-document-row {
    transition: none;
  }
}
</style>
