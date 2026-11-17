import {canonicalOoxmlFingerprint} from '@docx-editor.dev/core/store'
const W14='http://schemas.microsoft.com/office/word/2010/wordml'
// Browser sessions normalize duplicate paragraph identities. Replay uses current
// paragraph order, not these IDs; content, formatting and all other attrs remain.
export function replayFingerprint(part) {
 const visit=node=>{
  if(!node.children)return node
  const attributes=node.attributes?.filter(a=>!(a.namespaceUri===W14&&['paraId','textId'].includes(a.localName)))
  return {...node,attributes,children:node.children.map(visit)}
 }
 return canonicalOoxmlFingerprint({...part,root:visit(part.root)})
}
