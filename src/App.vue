<script setup lang="ts">
import { DocxEditor } from '@docx-editor.dev/vue'
import DocumentOutline from './DocumentOutline.vue'
import StudioHome from './components/StudioHome.vue'
import StudioBrand from './components/StudioBrand.vue'
import ChatMessages from './components/ChatMessages.vue'
import VersionHistory from './components/VersionHistory.vue'
import LayoutReview from './components/LayoutReview.vue'
import WorkspaceDialog from './components/WorkspaceDialog.vue'
import { ref, watch } from 'vue'
import baseZhCN from '@docx-editor.dev/i18n/zh-CN'
import {
  ArrowDownToLine,
  MoreHorizontal,
  Moon,
  Sun,
  PanelLeft,
  Bot,
  Check,
  FileText,
  FolderOpen,
  History,
  LoaderCircle,
  MessageSquareText,
  Plus,
  Send,
  Settings,
  SlidersHorizontal,
  Square,
  TriangleAlert,
  Undo2,
  X,
} from '@lucide/vue'
import '@docx-editor.dev/vue/styles.css'
import { blockBusyEditorInput } from './lib/editor-interaction'
import { useStudioWorkspace } from './composables/useStudioWorkspace'

const zhCN = {
  ...baseZhCN,
  editor: {
    ...baseZhCN.editor,
    fontSubstitutionNotice:
      '当前预览无法使用以下原字体：{fonts}；已使用可用字体替代显示。这不表示全文已改为这些字体。',
  },
}

const {
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
} = useStudioWorkspace()

const marginLabels: Record<string, string> = {
  top: '上边距',
  right: '右边距',
  bottom: '下边距',
  left: '左边距',
  gutter: '装订线',
}

const outlineOpen = ref(true)
const mobilePanel = ref<'editor' | 'outline' | 'assistant'>('editor')
const theme = ref<'light' | 'dark'>(
  localStorage.getItem('pi-docx-theme') === 'dark' ? 'dark' : 'light',
)

watch(
  theme,
  (value) => {
    document.documentElement.dataset.theme = value
    localStorage.setItem('pi-docx-theme', value)
  },
  { immediate: true },
)

function guardEditorInput(event: Event) {
  if (
    event.type === 'keydown' &&
    ['PageDown', 'PageUp', 'ArrowDown', 'ArrowUp', 'Home', 'End'].includes(
      (event as KeyboardEvent).key,
    )
  )
    pauseFollowing()
  if (aiBusy.value || refreshingAfterTask.value) blockBusyEditorInput(event)
}

function showMobileAssistant() {
  mobilePanel.value = 'assistant'
  assistantOpen.value = true
}

function createFromRecent() {
  void createDocument()
  recentOpen.value = false
}

function openFromRecent(id: string) {
  void switchDocument(id)
  recentOpen.value = false
}

function toggleTheme() {
  theme.value = theme.value === 'dark' ? 'light' : 'dark'
}
</script>

<template>
  <main :class="documentId ? 'studio-shell' : 'home-shell'">
    <input
      ref="inputRef"
      class="visually-hidden"
      aria-label="选择 DOCX 文件"
      type="file"
      accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      @change="openLocalFile"
    />
    <StudioHome
      v-if="!documentId"
      :documents="documents"
      :busy="documentBusy"
      :error="errorMessage"
      :theme="theme"
      @toggle-theme="toggleTheme"
      @upload="inputRef?.click()"
      @drop="uploadDocument"
      @open="switchDocument"
      @recent="recentOpen = true"
      @settings="settingsOpen = true"
      @dismiss-error="errorMessage = ''"
    />
    <template v-else>
      <header class="topbar">
        <a
          class="brand"
          href="#"
          aria-label="文序 DOCX 排版工作台首页"
          @click.prevent="closeDocument"
        >
          <StudioBrand />
        </a>
        <div class="document-heading">
          <span class="document-name">{{ fileName }}</span>
          <span class="save-state">
            <span class="state-dot" :class="{ dirty }" />
            {{ saveLabel }}
          </span>
        </div>
        <div class="top-actions">
          <button
            class="button button-quiet open-document"
            type="button"
            title="打开 DOCX"
            aria-label="打开 DOCX"
            :disabled="closingBlocked"
            @click="inputRef?.click()"
          >
            <FolderOpen :size="15" />
            <span>打开 DOCX</span>
          </button>
          <button
            class="button button-primary"
            type="button"
            :disabled="!dirty || closingBlocked"
            @click="saveFile"
          >
            <LoaderCircle v-if="saving" :size="15" class="spin" />
            <Check v-else :size="15" />
            <span>{{ saving ? '保存中' : '保存' }}</span>
          </button>
          <details class="document-menu">
            <summary
              class="icon-button"
              aria-label="更多文档操作"
              title="更多文档操作"
            >
              <MoreHorizontal :size="18" />
            </summary>
            <div
              class="document-menu-list"
              @click="
                ($event.currentTarget as HTMLElement)
                  .closest('details')
                  ?.removeAttribute('open')
              "
            >
              <button class="button" type="button" @click="recentOpen = true">
                <History :size="15" />最近文档
              </button>
              <button
                class="button"
                type="button"
                :disabled="closingBlocked"
                @click="createDocument()"
              >
                <Plus :size="15" />新建文档
              </button>
              <button
                class="button"
                type="button"
                :disabled="!sections.length || dirty || closingBlocked"
                @click="openPageSettings"
              >
                <SlidersHorizontal :size="15" />页面设置
              </button>
              <button
                class="button"
                type="button"
                :disabled="!documentId || closingBlocked"
                @click="exportFile"
              >
                <ArrowDownToLine :size="15" />导出副本
              </button>
              <button
                v-if="dirty"
                class="button"
                type="button"
                :disabled="closingBlocked"
                @click="discardChanges"
              >
                <Undo2 :size="15" />放弃修改
              </button>
              <button
                class="button"
                type="button"
                :disabled="!documentId || closingBlocked"
                :title="
                  aiBusy
                    ? '请先停止处理，等待任务结束后关闭'
                    : '关闭当前文档，保留已保存文件'
                "
                @click="closeDocument"
              >
                <X :size="15" />关闭文档
              </button>
            </div>
          </details>
          <span class="header-divider" />
          <button
            class="icon-button"
            type="button"
            :aria-label="theme === 'light' ? '切换深色界面' : '切换浅色界面'"
            :title="theme === 'light' ? '切换深色界面' : '切换浅色界面'"
            @click="toggleTheme"
          >
            <Moon v-if="theme === 'light'" :size="16" />
            <Sun v-else :size="16" />
          </button>
        </div>
      </header>

      <p v-if="tocNotice" class="panel-description" role="status">
        {{ tocNotice }}
      </p>
      <div class="workspace-toolbar">
        <details class="document-menu">
          <summary
            class="button button-text"
            aria-label="目录操作"
            @mousedown.prevent
          >
            目录
          </summary>
          <div class="document-menu-list">
            <button
              class="button"
              type="button"
              :disabled="closingBlocked"
              @mousedown.prevent
              @click="runTocAction('insert')"
            >
              插入目录（标题 1～3）
            </button>
            <button
              class="button"
              type="button"
              :disabled="closingBlocked"
              @mousedown.prevent
              @click="runTocAction('entire')"
            >
              更新整个目录
            </button>
            <button
              class="button"
              type="button"
              :disabled="closingBlocked"
              @mousedown.prevent
              @click="runTocAction('pageNumbers')"
            >
              仅更新页码
            </button>
          </div>
        </details>
        <button
          class="button button-text outline-toggle-button"
          type="button"
          :aria-pressed="outlineOpen"
          @click="outlineOpen = !outlineOpen"
        >
          <PanelLeft :size="15" />文档大纲
        </button>
        <span class="workspace-context">{{
          documentId ? '排版工作台' : '文档启动台'
        }}</span>
        <span v-if="aiBusy" class="task-label" role="status">
          {{
            changeLocation
              ? '正在更新：' + changeLocation
              : activeTool || '正在处理文档'
          }}
          <button
            v-if="!followChanges"
            class="button button-text"
            type="button"
            @click="resumeFollowing"
          >
            跟随修改
          </button>
          <span v-else>跟随修改中 · 滚动可暂停</span>
          <button class="button button-text" type="button" @click="stopMessage">
            <Square :size="12" />停止处理
          </button>
        </span>
        <button
          v-if="!assistantOpen"
          class="button button-text"
          type="button"
          @click="assistantOpen = true"
        >
          <MessageSquareText :size="15" />展开文档处理
        </button>
      </div>

      <nav class="mobile-panels" aria-label="工作区面板">
        <button
          type="button"
          :aria-pressed="mobilePanel === 'outline'"
          @click="mobilePanel = 'outline'"
        >
          大纲
        </button>
        <button
          type="button"
          :aria-pressed="mobilePanel === 'editor'"
          @click="mobilePanel = 'editor'"
        >
          文档
        </button>
        <button
          type="button"
          :aria-pressed="mobilePanel === 'assistant'"
          @click="showMobileAssistant"
        >
          助手
        </button>
      </nav>

      <div v-if="errorMessage" class="workspace-error" role="alert">
        <TriangleAlert :size="15" />
        <span>{{ errorMessage }}</span>
        <button
          class="icon-button"
          type="button"
          aria-label="关闭错误提示"
          @click="errorMessage = ''"
        >
          <X :size="14" />
        </button>
      </div>

      <section
        class="workspace"
        :class="{
          'assistant-closed': !assistantOpen,
          'outline-hidden': !outlineOpen,
        }"
        :data-panel="mobilePanel"
      >
        <aside class="rail">
          <div
            v-show="documentId"
            id="document-navigation"
            class="docx-editor"
          />
          <div v-if="!documentId" class="closed-outline">
            <h2>文档大纲</h2>
            <p>打开文档后，章节结构会显示在这里。</p>
          </div>
        </aside>

        <section class="editor-stage" aria-label="文档编辑区">
          <div
            v-if="documentId"
            class="editor-frame"
            :aria-busy="aiBusy || refreshingAfterTask"
            @wheel.passive="pauseFollowing"
            @touchstart.passive="pauseFollowing"
            @pointerdown.capture="
              (event) => {
                pauseFollowing()
                guardEditorInput(event)
              }
            "
            @beforeinput.capture="guardEditorInput"
            @keydown.capture="guardEditorInput"
            @paste.capture="guardEditorInput"
            @cut.capture="guardEditorInput"
            @drop.capture="guardEditorInput"
            @click.capture="guardEditorInput"
            @dblclick.capture="guardEditorInput"
            @compositionstart.capture="guardEditorInput"
          >
            <div v-if="documentBusy" class="editor-loading" role="status">
              <div class="loading-lines"><i /><i /><i /></div>
              <p>正在同步文档…</p>
            </div>
            <DocxEditor
              v-if="!documentBusy"
              :key="documentKey"
              ref="editorRef"
              :document="documentData"
              :title="fileName"
              :rulers="false"
              :navigation="false"
              :color-mode="theme"
              :i18n="zhCN"
              :mode="documentBusy ? 'view' : 'edit'"
              locale="zh-CN"
            >
              <template #titleBarLeft>
                <Teleport v-if="navigationReady" to="#document-navigation">
                  <DocumentOutline :file-name="fileName" />
                </Teleport>
              </template>
            </DocxEditor>
          </div>
          <div v-else class="document-empty">
            <div class="empty-document-icon"><FileText :size="30" /></div>
            <h2>未打开文档</h2>
            <p>打开一份文档，开始阅读、排版与编辑。</p>
            <div>
              <button
                class="button button-primary"
                type="button"
                @click="inputRef?.click()"
              >
                打开 DOCX
              </button>
              <button
                class="button button-quiet"
                type="button"
                @click="createDocument()"
              >
                新建文档
              </button>
              <button
                class="button button-quiet"
                type="button"
                @click="recentOpen = true"
              >
                最近文档
              </button>
            </div>
          </div>
        </section>

        <aside
          v-if="assistantOpen"
        class="assistant-panel"
          aria-label="文档处理面板"
        >
          <header class="assistant-header">
            <span class="assistant-mark"><Bot :size="16" /></span>
            <div class="assistant-title">
              <strong>文档处理</strong>
              <span>
                <i :class="{ ready: modelConfigured }" />
                {{ modelConfigured ? '处理服务已就绪' : '需要配置处理服务' }}
              </span>
            </div>
            <button
              class="icon-button"
              type="button"
              aria-label="模型设置"
              title="模型设置"
              @click="settingsOpen = true"
            >
              <Settings :size="15" />
            </button>
            <button
              class="icon-button close-assistant"
              type="button"
              aria-label="收起助手"
              title="收起助手"
              @click="assistantOpen = false"
            >
              <MessageSquareText :size="15" />
            </button>
          </header>

          <div class="side-tabs" role="tablist" aria-label="右侧面板">
            <button
              id="chat-tab"
              type="button"
              role="tab"
              :aria-selected="rightTab === 'chat'"
              aria-controls="chat-panel"
              @click="rightTab = 'chat'"
            >
              对话助手
            </button>
            <button
              id="history-tab"
              type="button"
              role="tab"
              :aria-selected="rightTab === 'history'"
              aria-controls="history-panel"
              @click="rightTab = 'history'"
            >
              修改历史 <span>{{ historyGroups.length }}</span>
            </button>
          </div>

          <section
            v-show="rightTab === 'history'"
            id="history-panel"
            role="tabpanel"
            aria-labelledby="history-tab"
            class="history-container"
          >
            <div v-if="pendingWork" class="panel-description">
              <p>本次处理未完成，修改保留在工作副本中，尚未形成结果版本。</p>
              <button
                class="button button-text"
                :disabled="closingBlocked || dirty"
                @click="savePartialResult"
              >
                保存部分结果
              </button>
              <button
                class="button button-text"
                :disabled="closingBlocked || dirty"
                @click="restoreVersion(pendingWork.startRevision)"
              >
                放弃本次修改
              </button>
              <p>继续处理可在对话中发送补充要求。</p>
            </div>
            <VersionHistory
              :groups="historyGroups"
              :revision="revision"
              :validation="formatValidation"
              :blocked="closingBlocked || dirty"
              @restore="restoreVersion"
            />
            <LayoutReview
              v-if="documentId"
              :editor="editorRef?.getEditor()"
              :blocked="closingBlocked"
            />
          </section>

          <ChatMessages
            v-show="rightTab === 'chat'"
            id="chat-panel"
            role="tabpanel"
            aria-labelledby="chat-tab"
            :lines="chatLines"
            :busy="aiBusy"
            :has-document="Boolean(documentId)"
          />

          <form
            v-show="rightTab === 'chat'"
            class="composer"
            @submit.prevent="sendMessage"
          >
            <label class="composer-label" for="chat-input">文档处理需求</label>
            <textarea
              id="chat-input"
              v-model="chatInput"
              aria-label="输入文档处理需求"
              rows="3"
              placeholder="描述希望调整的文档格式…"
              @keydown.enter.exact="onComposerEnter"
            />
            <div class="composer-footer">
              <span
                class="composer-hint"
                role="status"
                :title="sendBlockedReason"
              >
                {{ sendBlockedReason || `基于当前文档 v${revision}` }}
              </span>
              <button
                v-if="!modelConfigured"
                class="composer-action"
                type="button"
                @click="settingsOpen = true"
              >
                配置模型
              </button>
              <button
                v-else-if="dirty && !aiBusy && !documentBusy"
                class="composer-action"
                type="button"
                :disabled="saving"
                @click="saveFile"
              >
                保存文档
              </button>
              <button
                v-if="aiBusy"
                class="send-button stop-button"
                type="button"
                title="停止处理"
                aria-label="停止处理"
                @click="stopMessage"
              >
                <Square :size="14" />
              </button>
              <button
                v-else
                class="send-button"
                type="submit"
                :disabled="Boolean(sendBlockedReason) || !chatInput.trim()"
                :title="sendBlockedReason || '发送'"
                aria-label="发送"
              >
                <Send :size="15" />
              </button>
            </div>
          </form>
        </aside>
        <button
          v-else
          class="assistant-reopen"
          type="button"
          aria-label="展开文档处理面板"
          title="展开文档处理面板"
          @click="assistantOpen = true"
        >
          <Bot :size="17" />
        </button>
      </section>
    </template>

    <WorkspaceDialog
      v-model:open="recentOpen"
      label="最近文档"
      :busy="documentBusy"
    >
      <section class="settings-dialog" aria-label="最近文档">
        <header class="settings-header">
          <h2>文档空间</h2>
          <button
            class="icon-button"
            type="button"
            aria-label="关闭最近文档"
            @click="recentOpen = false"
          >
            <X :size="16" />
          </button>
        </header>
        <section class="rail-section">
          <div class="section-heading">
            <span class="section-label">最近文档</span>
            <button
              v-if="documentId"
              class="icon-button"
              type="button"
              aria-label="新建文档"
              title="新建文档"
              @click="createFromRecent"
            >
              <Plus :size="16" />
            </button>
          </div>
          <p v-if="!documents.length" class="empty-recent-note">
            当前空间暂无已保存文档。
          </p>
          <button
            v-for="doc in documents"
            :key="doc.id"
            class="recent-file"
            :class="{ active: doc.id === documentId }"
            type="button"
            :disabled="aiBusy"
            @click="openFromRecent(doc.id)"
          >
            <span class="file-icon"><FileText :size="15" /></span>
            <span class="file-copy">
              <strong>{{ doc.name }}</strong>
              <span>v{{ doc.revision }} · {{ formatTime(doc.updatedAt) }}</span>
            </span>
          </button>
        </section>
      </section>
    </WorkspaceDialog>

    <WorkspaceDialog
      v-model:open="pageSettingsOpen"
      label="页面设置"
      :busy="documentBusy"
    >
      <form
        class="settings-dialog page-dialog"
        aria-labelledby="page-settings-title"
        @submit.prevent="savePageSettings"
      >
        <header class="settings-header">
          <div>
            <h2 id="page-settings-title">页面设置</h2>
            <p>修改选定分节的页面尺寸与边距。</p>
          </div>
          <button
            class="icon-button"
            type="button"
            aria-label="关闭"
            @click="pageSettingsOpen = false"
          >
            <X :size="16" />
          </button>
        </header>
        <label v-if="sections.length > 1" class="setting-field">
          <span>应用范围</span>
          <select v-model="sectionSelection">
            <option value="all">全部分节（{{ sections.length }}）</option>
            <option
              v-for="section in sections"
              :key="section.index"
              :value="String(section.index)"
            >
              第 {{ section.index + 1 }} 节
            </option>
          </select>
        </label>
        <div class="page-row">
          <label class="setting-field">
            <span>纸张</span>
            <select
              :value="pagePreset"
              @change="
                setPagePreset(($event.target as HTMLSelectElement).value)
              "
            >
              <option value="A4">A4</option>
              <option value="Letter">Letter</option>
              <option value="Legal">Legal</option>
              <option value="custom">自定义</option>
            </select>
          </label>
          <fieldset class="orientation-field">
            <legend>方向</legend>
            <div class="segmented-control">
              <label :class="{ selected: pageOrientation === 'portrait' }">
                <input
                  v-model="pageOrientation"
                  type="radio"
                  value="portrait"
                />纵向
              </label>
              <label :class="{ selected: pageOrientation === 'landscape' }">
                <input
                  v-model="pageOrientation"
                  type="radio"
                  value="landscape"
                />横向
              </label>
            </div>
          </fieldset>
        </div>
        <div class="page-row dimensions-row">
          <label class="setting-field">
            <span>宽度（mm）</span>
            <input
              v-model.number="pageWidthMm"
              type="number"
              min="17.6"
              max="558.8"
              step="0.1"
              required
              @input="pagePreset = 'custom'"
            />
          </label>
          <label class="setting-field">
            <span>高度（mm）</span>
            <input
              v-model.number="pageHeightMm"
              type="number"
              min="17.6"
              max="558.8"
              step="0.1"
              required
              @input="pagePreset = 'custom'"
            />
          </label>
        </div>
        <div class="margin-grid">
          <label
            v-for="(_, key) in pageMarginsMm"
            :key="key"
            class="setting-field"
          >
            <span>{{ marginLabels[key] }}（mm）</span>
            <input
              v-model.number="pageMarginsMm[key]"
              type="number"
              min="0"
              max="558.8"
              step="0.1"
              required
            />
          </label>
        </div>
        <footer class="settings-actions">
          <button
            class="button button-quiet"
            type="button"
            @click="pageSettingsOpen = false"
          >
            取消
          </button>
          <button class="button button-primary" type="submit">
            <Check :size="14" />应用设置
          </button>
        </footer>
      </form>
    </WorkspaceDialog>

    <WorkspaceDialog
      v-model:open="settingsOpen"
      label="模型设置"
      :busy="settingsSaving"
    >
      <form
        class="settings-dialog"
        aria-labelledby="settings-title"
        @submit.prevent="saveModelSettings"
      >
        <header class="settings-header">
          <div>
            <h2 id="settings-title">模型设置</h2>
            <p>配置一个兼容 OpenAI Chat Completions 的服务。</p>
          </div>
          <button
            class="icon-button"
            type="button"
            aria-label="关闭"
            @click="settingsOpen = false"
          >
            <X :size="16" />
          </button>
        </header>
        <label class="setting-field">
          <span>服务地址</span>
          <input
            v-model="modelBaseUrl"
            type="url"
            required
            placeholder="https://api.openai.com/v1"
          />
        </label>
        <label class="setting-field">
          <span>模型名称</span>
          <input
            v-model="modelName"
            required
            maxlength="100"
            placeholder="gpt-4o-mini"
          />
        </label>
        <label class="setting-field">
          <span>API 密钥</span>
          <input
            v-model="modelApiKey"
            type="password"
            autocomplete="new-password"
            maxlength="500"
            :placeholder="
              modelConfigured ? '留空以保持当前密钥' : '输入服务 API 密钥'
            "
          />
        </label>
        <p class="settings-privacy">
          发送对话后，用户指令和当前文档结构摘要会传到上面的模型服务。本页修改仅对当前服务进程有效。可在项目
          .env 文件中配置启动默认值，重启后自动加载。
        </p>
        <footer class="settings-actions">
          <button
            class="button button-quiet"
            type="button"
            @click="settingsOpen = false"
          >
            取消
          </button>
          <button
            class="button button-primary"
            type="submit"
            :disabled="settingsSaving || (!modelApiKey && !modelConfigured)"
          >
            <LoaderCircle v-if="settingsSaving" :size="14" class="spin" />
            <Check v-else :size="14" />
            保存设置
          </button>
        </footer>
      </form>
    </WorkspaceDialog>

    <dialog
      ref="closeDialog"
      class="close-document-dialog"
      aria-labelledby="close-document-title"
      @cancel="saving && $event.preventDefault()"
    >
      <h2 id="close-document-title">关闭前保存修改？</h2>
      <p>
        “{{
          fileName
        }}”有未保存的修改。放弃修改只会丢弃本次未保存的内容，已保存的版本会保留。
      </p>
      <p v-if="errorMessage" role="alert" class="error-note">
        {{ errorMessage }}
      </p>
      <div class="settings-actions">
        <button
          class="button button-quiet"
          type="button"
          :disabled="saving"
          @click="closeDialog?.close()"
        >
          取消关闭
        </button>
        <button
          class="button button-quiet"
          type="button"
          :disabled="saving"
          @click="finishClose"
        >
          不保存并关闭
        </button>
        <button
          class="button button-primary"
          type="button"
          :disabled="saving"
          @click="saveAndClose"
        >
          {{ saving ? '正在保存…' : '保存并关闭' }}
        </button>
      </div>
    </dialog>
  </main>
</template>
