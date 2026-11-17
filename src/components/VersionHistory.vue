<script setup lang="ts">
import { History, ShieldCheck, TriangleAlert } from '@lucide/vue'
import type { Version, Validation, HistoryTask } from '../types'

defineProps<{
  groups: { id: string; task?: HistoryTask; versions: Version[] }[]
  revision: number
  validation: Validation | null
  blocked: boolean
}>()

defineEmits<{ restore: [revision: number] }>()

function state(value?: string) {
  return (
    ({
      finished: '处理已结束 · 页面效果待确认',
      partial: '部分完成',
      cancelled: '已取消',
    } as Record<string, string>)[value || ''] || '已保存中间结果 · 任务状态待确认'
  )
}

function time(value: string) {
  return new Date(value).toLocaleString('zh-CN', {
    month: 'numeric',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  })
}
</script>

<template>
  <div class="history-panel">
    <div v-if="!groups.length" class="chat-empty">
      <span class="empty-symbol"><History :size="22" /></span>
      <strong>暂无修改记录</strong>
      <p>这里展示明确保存的成果；处理中间检查点不作为用户版本。</p>
    </div>
    <p v-else class="panel-description">这里展示已保存的成果。内部检查点用于恢复，不代表独立结果。</p>
    <ol v-if="groups.length" class="version-list">
      <li v-for="group in groups" :key="group.id" class="version-row">
        <span class="version-track">
          <span class="version-node" :class="{ current: group.versions.some(v => v.revision === revision) }" />
        </span>
        <div class="version-copy">
          <strong>
            {{ group.task ? 'AI 文档处理' : group.versions[0]!.summary }}
            <span v-if="group.versions.some(v => v.revision === revision)" class="current-label">当前</span>
          </strong>
          <template v-if="group.task">
            <p class="task-summary-text">{{ group.task.summary }}</p>
            <p class="task-state-meta">{{ state(group.task.state) }} · 已保存 {{ group.versions.length }} 个步骤</p>
            <p v-if="group.versions.some(v => v.revision === revision) && group.versions[0]!.revision !== revision">
              预览位于中间步骤，最新结果尚未显示
            </p>
          </template>
          <time>{{ time(group.versions[0]!.time) }}</time>
          <div class="version-actions">
            <button
              v-if="group.task"
              class="button button-text"
              type="button"
              :disabled="blocked || revision === group.task.startRevision"
              @click="$emit('restore', group.task.startRevision)"
            >
              恢复到处理前
            </button>
            <button
              v-if="group.versions[0]!.revision !== revision"
              class="button button-text"
              type="button"
              :disabled="blocked"
              @click="$emit('restore', group.versions[0]!.revision)"
            >
              {{ group.task ? '恢复本次结果' : '恢复此版本' }}
            </button>
          </div>
          <details v-if="group.task" class="step-details">
            <summary>查看 {{ group.versions.length }} 个处理步骤</summary>
            <div v-for="version in [...group.versions].reverse()" :key="version.revision" class="step-version-item">
              <div class="step-version-head">
                <span>{{ version.task?.stepLabel || version.summary }} · v{{ version.revision }}</span>
                <button
                  class="button button-text"
                  type="button"
                  :disabled="blocked || version.revision === revision"
                  @click="$emit('restore', version.revision)"
                >
                  恢复到此步骤
                </button>
              </div>
              <details v-if="version.validation.operationReceipt" class="receipt-details">
                <summary>操作详情</summary>
                <pre>{{ JSON.stringify(version.validation.operationReceipt.operations, null, 2) }}</pre>
              </details>
            </div>
          </details>
        </div>
      </li>
    </ol>
    <div v-if="validation" class="validation-summary" :class="validation.status">
      <ShieldCheck v-if="validation.status === 'passed'" :size="15" />
      <TriangleAlert v-else :size="15" />
      <span>{{ validation.status === 'passed' ? '结构与操作检查通过' : '包含未验证项' }}</span>
    </div>
    <details v-if="validation" class="validation-details">
      <summary>查看质量检查与未验证项</summary>
      <ul>
        <li v-for="check in validation.checks" :key="check">{{ check }}</li>
      </ul>
      <p v-for="limit in validation.limitations" :key="limit">{{ limit }}</p>
      <p>属性与结构检查不代表页面排版已通过；页面状态：{{ validation.layoutRisks?.visualVerified ? '已验证' : '未验证' }}。</p>
    </details>
  </div>
</template>
