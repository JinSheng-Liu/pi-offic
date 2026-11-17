// Server execution boundary for the installed editor public Automation API.
import {blankDocumentBytes} from '@docx-editor.dev/core'
import {createServerAutomationHost} from '@docx-editor.dev/core/automation'
import {problem} from './package.mjs'

export function createBlankDocument() {
  return executeAutomationBatch(blankDocumentBytes(),host=>{
    const document=host.execute({operations:[{op:'getDocument'}]}).results[0].value.handle
    const sections=host.execute({operations:[{op:'getSections',document}]}).results[0].value.handles
    return sections.map(section=>({op:'setPageSetup',section,setup:{pageWidth:595.3,pageHeight:841.9,orientation:'portrait'}}))
  }).bytes
}

export function openAutomationDocument(bytes) {
  const opened=createServerAutomationHost(new Uint8Array(bytes))
  if(!opened.ok) throw problem(`编辑器引擎无法打开 DOCX：${opened.reason}`,422)
  return opened.host
}

export function executeAutomationBatch(bytes,operations,{revision}={}) {
  const host=openAutomationDocument(bytes)
  try {
    operations=typeof operations==='function'?operations(host):operations
    const response=host.execute({operations,expectedRevision:revision})
    if(!response.ok) throw problem(response.results.find(r=>r.status==='error')?.error?.message||'编辑器自动化批次失败',422)
    const saved=host.save()
    if(!saved.ok) throw problem(saved.error.message||'编辑器导出 DOCX 失败',422)
    return {bytes:Buffer.from(saved.bytes),response,validation:{status:'passed',checks:['Editor Automation 批次提交','编辑器保存后导出'],limitations:['页面视觉和复杂对象仍需专项验收']}}
  } finally { host.dispose() }
}

