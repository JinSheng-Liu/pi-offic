import {readOoxmlPackage,detectBodyTocs,bodyStoryRoot,storyParagraphs} from '@docx-editor.dev/core/store'
import {replayFingerprint} from '../../shared/replay-fingerprint.mjs'
import {digest} from './package.mjs'
function fingerprint(bytes) {
  const loaded=readOoxmlPackage(bytes)
  if(!loaded.ok)return null
  return digest(new TextEncoder().encode(replayFingerprint(loaded.package.parts.get(loaded.package.mainDocumentPart))))
}
export function createFormatReplay(before,after,targets) {
  const loaded=readOoxmlPackage(before)
  if(!loaded.ok)return undefined
  const part=loaded.package.parts.get(loaded.package.mainDocumentPart),paragraphs=storyParagraphs(bodyStoryRoot(part))
  const protectedIds=new Set(detectBodyTocs(part).flatMap(t=>[t.beginParagraphId,t.endParagraphId,...t.resultParagraphIds]))
  if(targets.some(t=>t.indices.some(i=>protectedIds.has(paragraphs[i]?.id))))return undefined
  const replay={kind:'format-v1',before:fingerprint(before),after:fingerprint(after),targets}
  // Large or unsupported changes keep the normal full-document synchronization.
  return replay.before&&replay.after&&JSON.stringify(replay).length<=200000?replay:undefined
}
