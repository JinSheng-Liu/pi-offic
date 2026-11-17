import { createHash } from 'node:crypto'

// Only identical successful reads are elided. Revisions, writes and errors remain intact.
// The persisted messages are not changed; Pi uses this projection for model requests.
export async function deduplicateDocumentReads(messages) {
  const seen = new Map()
  const projected = messages.slice()
  for (let i = messages.length - 1; i >= 0; i--) {
    const message = messages[i]
    if (message.role !== 'toolResult' || message.toolName !== 'inspect_document' || message.isError) continue
    if (message.content?.length !== 1 || message.content[0].type !== 'text') continue
    const value = message.content[0].text
    let data
    try { data = JSON.parse(value) } catch { continue }
    if (!Number.isInteger(data.revision) || !data.hash) continue
    const key = createHash('sha256').update(value).digest('hex')
    if (!seen.has(key)) { seen.set(key, message.toolCallId); continue }
    projected[i] = { ...message, content: [{ type: 'text', text: JSON.stringify({
      revision: data.revision, hash: data.hash, identicalReadAvailableAt: seen.get(key),
      note: '完全相同的成功读取结果见后续指定 toolCallId；此处仅去重，不表示文档或任务已完成。',
    }) }] }
  }
  return projected
}

// Bound model-facing read history, never the persisted session or write receipts.
export async function budgetDocumentContext(messages,{budgetBytes=180000}={}) {
  const projected=await deduplicateDocumentReads(messages)
  const size=value=>Buffer.byteLength(JSON.stringify(value),'utf8')
  let remaining=budgetBytes,keptLatest=false
  for(let i=projected.length-1;i>=0;i--){
    const m=projected[i]
    if(m.role!=='toolResult'||m.toolName!=='inspect_document'||m.isError||m.content?.length!==1||m.content[0].type!=='text')continue
    let data;try{data=JSON.parse(m.content[0].text)}catch{continue}
    if(!Number.isInteger(data.revision)||!data.hash)continue
    // details is SDK metadata, not additional document evidence.
    projected[i]={...m,details:undefined}
    const bytes=size(projected[i])
    if(!keptLatest||bytes<=remaining){keptLatest=true;remaining=Math.max(0,remaining-bytes);continue}
    const compact={revision:data.revision,hash:data.hash,view:data.view,pagination:data.pagination,
      omittedForContextBudget:true,note:'此读取正文已从模型上下文移除；原会话保留。需要这些段落或格式时必须重新读取当前版本，不能凭本摘要执行修改。'}
    projected[i]={...m,details:undefined,content:[{type:'text',text:JSON.stringify(compact)}]}
  }
  return projected
}

export function inspectModelPayload(payload) {
  const requestBytes=Buffer.byteLength(JSON.stringify(payload),'utf8')
  const stats={requestBytes,messageCount:payload.messages?.length||0,toolCount:payload.tools?.length||0}
  // Byte bound is intentionally not presented as an exact token estimate.
  if(requestBytes>600000)throw new Error(`模型请求超过本地上下文预算（${requestBytes} 字节）；已停止发送，已保存修改保留。请开启新对话或缩小本轮范围。`)
  return stats
}
