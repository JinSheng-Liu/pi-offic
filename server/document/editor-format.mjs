import {readOoxmlPackage,writeOoxmlPackage,TreePackageStore} from '@docx-editor.dev/core/store'
import {compileFormatOps} from '../../shared/editor-format-ops.mjs'
import {problem} from './package.mjs'
export {paragraphProperties,fontProperties} from '../../shared/editor-format-ops.mjs'

export function formatEditorParagraphs(bytes,targets,signal) {
  const loaded=readOoxmlPackage(new Uint8Array(bytes))
  if(!loaded.ok)throw problem(`编辑器无法读取文档：${loaded.reason}`,422)
  const pkg=loaded.package,store=new TreePackageStore(pkg,pkg.parts.get(pkg.mainDocumentPart))
  for(const {indices,op} of targets) {
    signal?.throwIfAborted()
    const ops=compileFormatOps(store.bodyStore().part,indices,op)
    const result=store.transact({kind:'body'},ctx=>{for(const op of ops)if(!ctx.apply(op))return})
    if(!result.ok)throw problem(`编辑器格式事务失败：${result.reason}${result.detail?` (${result.detail})`:''}`,422)
  }
  signal?.throwIfAborted()
  return Buffer.from(writeOoxmlPackage(store.currentPackage()))
}
