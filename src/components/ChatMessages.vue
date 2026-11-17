<script setup lang="ts">
import { nextTick, ref, watch } from 'vue'
import { Bot, Check, CircleAlert, MessageSquareText } from '@lucide/vue'
import MarkdownIt from 'markdown-it'
import type { ChatLine } from '../types'

const props = defineProps<{ lines: ChatLine[]; busy: boolean; hasDocument: boolean }>()
const container = ref<HTMLElement | null>(null)
const follow = ref(true)
const markdown = new MarkdownIt({ html: false, linkify: false, breaks: true })

function trackScroll() {
  const el = container.value
  if (el) follow.value = el.scrollHeight - el.scrollTop - el.clientHeight < 90
}

watch(
  () => [props.lines.length, props.lines.at(-1)?.text, props.lines.at(-1)?.steps?.map(step => step.state).join(',')],
  async () => {
    if (!follow.value) return
    await nextTick()
    container.value?.scrollTo({ top: container.value.scrollHeight })
  },
  { flush: 'post' },
)
</script>

<template>
  <div ref="container" class="chat-scroll" :aria-busy="busy" @scroll.passive="trackScroll">
    <div v-if="!lines.length" class="chat-empty">
      <span class="empty-symbol"><MessageSquareText :size="22" /></span>
      <strong>{{ hasDocument ? '提出排版要求' : '从一份文档开始' }}</strong>
      <p>{{ hasDocument ? '说明要调整的范围、字体、字号或段落格式。处理步骤与核对清单会在这里显示。' : '打开文档后，可以在这里提出排版要求，或查看处理结果。' }}</p>
      <div v-if="hasDocument" class="prompt-example">例如：将全文正文设为仿宋三号，行距固定 28 磅，首行缩进 2 字符。</div>
    </div>
    <article v-for="(line, index) in lines" :key="index" class="chat-line" :class="line.role">
      <div class="message-author">
        <Bot v-if="line.role === 'assistant'" :size="14" />
        <span>{{ line.role === 'assistant' ? '文档处理' : '你' }}</span>
      </div>
      <div class="message-body">
        <ol v-if="line.steps?.length" class="execution-steps">
          <li v-for="step in line.steps" :key="step.id" :data-state="step.state">
            <Check v-if="step.state === 'complete'" :size="13" />
            <CircleAlert v-else-if="step.state === 'error'" :size="13" />
            <span v-else class="pending-step" />
            <span class="step-label">{{ step.label }}</span>
            <small v-if="!step.id.startsWith('requirement-')">{{
              step.state === 'running'
                ? '执行中'
                : step.state === 'error'
                  ? '失败'
                  : step.state === 'pending'
                    ? '待执行'
                    : step.state === 'cancelled'
                      ? '已取消'
                      : '已提交'
            }}</small>
          </li>
        </ol>
        <div v-if="line.role === 'assistant'" class="markdown-body" v-html="markdown.render(line.text)" />
        <p v-else class="user-message">{{ line.text }}</p>
        <span v-if="line.status" class="message-status" role="status">{{ line.status }}</span>
      </div>
    </article>
  </div>
</template>
