import { format } from './editor-adapter.mjs'
import { insertCaptions } from './ooxml-captions.mjs'
import { inspect } from './structure.mjs'
import { problem } from './package.mjs'
import { qualityReport } from './quality-report.mjs'
import { restoreReferenceCover } from './reference-cover.mjs'

// Pure document computation: no storage or commit side effects.
export async function computeBatch({source, reference, batch, revision}, signal) {
  let result, summary
  if (batch.kind === 'format') {
    result = await format(source, batch.operations, signal)
    summary = `格式调整：${batch.operations.length} 项操作`
  } else if (batch.kind === 'captions') {
    result = insertCaptions(source, batch.captions, signal)
    summary = `添加${result.affected}个图表题注`
  } else if (['referenceCover','referenceAuthorization','referenceDeclaration'].includes(batch.kind)) {
    result=restoreReferenceCover(source,reference,signal,batch.kind==='referenceDeclaration'?'declaration':batch.kind==='referenceAuthorization'?'authorization':'cover')
    summary=batch.kind==='referenceDeclaration'?'按参考文档恢复原创性声明样式':batch.kind==='referenceAuthorization'?'按参考文档恢复授权书样式':'按参考文档恢复封面样式'
  } else throw problem('未知文档操作批次')
  // Both adapters must produce a document that the reader can open before commit.
  const inspected = await inspect(result.bytes)
  signal?.throwIfAborted()
  const validation = qualityReport(inspected, result.validation)
  validation.operationReceipt = {inputRevision:revision,kind:batch.kind,
    operations:batch.kind==='format'?batch.operations:['referenceCover','referenceAuthorization','referenceDeclaration'].includes(batch.kind)?[{referenceId:batch.referenceId,referenceRevision:batch.referenceRevision}]:batch.captions,
    affected:result.affected,batchChecks:result.validation.checks,
    requirementsVerified:false,visualVerified:false}
  return {result, summary, validation}
}
