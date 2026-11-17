<script setup lang="ts">
import { FileText } from '@lucide/vue'
</script>

<template>
  <span class="studio-brand">
    <span class="studio-brand-mark"><FileText :size="20" /></span>
    <span class="studio-brand-name"><b>文序</b><span>DOCX 排版</span></span>
  </span>
</template>

<style scoped>
.studio-brand {
  display: inline-flex;
  align-items: center;
  gap: 11px;
  color: var(--studio-text);
  font-size: 16px;
  font-weight: 650;
  white-space: nowrap;
}
.studio-brand-mark {
  display: grid;
  place-items: center;
  width: 34px;
  height: 34px;
  flex-shrink: 0;
  border-radius: 7px;
  color: var(--studio-on-accent);
  background: var(--studio-accent);
}
.studio-brand-name {
  display: inline-flex;
  align-items: baseline;
  gap: 8px;
}
.studio-brand-name b {
  font-size: 18px;
  font-weight: 750;
}
.studio-brand-name > span {
  color: var(--studio-muted);
  font-size: 11px;
  font-weight: 500;
}
@media (max-width: 900px) {
  .topbar .studio-brand-name > span {
    display: none;
  }
}
</style>
