<script setup lang="ts">
import { ref, watch } from 'vue'

const props = defineProps<{
  open: boolean
  label: string
  busy?: boolean
}>()

const emit = defineEmits<{
  'update:open': [value: boolean]
}>()

const dialogRef = ref<HTMLDialogElement | null>(null)

watch(
  () => props.open,
  value => {
    const el = dialogRef.value
    if (!el) return
    if (value && !el.open) el.showModal()
    else if (!value && el.open) el.close()
  },
)

function onCancel(event: Event) {
  if (props.busy) {
    event.preventDefault()
    return
  }
  emit('update:open', false)
}

function onClose() {
  if (props.open) emit('update:open', false)
}

function onBackdropClick(event: MouseEvent) {
  if (props.busy) return
  if (event.target === dialogRef.value) {
    emit('update:open', false)
  }
}
</script>

<template>
  <dialog
    ref="dialogRef"
    class="workspace-dialog"
    :aria-label="label"
    @cancel="onCancel"
    @close="onClose"
    @click="onBackdropClick"
  >
    <slot />
  </dialog>
</template>
