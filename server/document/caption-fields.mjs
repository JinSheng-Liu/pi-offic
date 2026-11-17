import {W} from './package.mjs'

// Traverse the same block paragraphs as the editor reader, never textbox paragraphs.
export function captionFieldDetails(document) {
  const result=[]
  const walk=node=>{
    if(node.namespaceURI===W&&node.localName==='p') {
      const instructions=[];const stack=[]
      const visit=n=>{
        if(n!==node&&n.namespaceURI===W&&n.localName==='p')return
        if(n.namespaceURI===W&&n.localName==='fldSimple')instructions.push(n.getAttributeNS(W,'instr')||'')
        if(n.namespaceURI===W&&n.localName==='fldChar') {
          const type=n.getAttributeNS(W,'fldCharType')
          if(type==='begin')stack.push({text:'',reading:true})
          else if(type==='separate'&&stack.length)stack.at(-1).reading=false
          else if(type==='end'&&stack.length)instructions.push(stack.pop().text)
        }
        if(n.namespaceURI===W&&n.localName==='instrText'&&stack.at(-1)?.reading)stack.at(-1).text+=n.textContent||''
        for(const child of Array.from(n.childNodes||[]))visit(child)
      }
      visit(node)
      const sequences=instructions.map(s=>/^\s*SEQ\s+(?:"([^"]+)"|([^\s\\]+))/i.exec(s)).filter(Boolean).map(m=>m[1]||m[2])
      result.push({numbering:sequences.length?'sequence-field':'manual-or-unverified',sequenceIdentifiers:sequences})
      return
    }
    for(const child of Array.from(node.childNodes||[]))walk(child)
  }
  walk(document.getElementsByTagNameNS(W,'body')[0])
  return result
}
