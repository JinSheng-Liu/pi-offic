import {posix} from 'node:path'
import {W,xml} from './package.mjs'
const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const children=(n,name)=>Array.from(n?.childNodes||[]).filter(c=>c.namespaceURI===W&&c.localName===name)
export function inspectFooters(parts,sections){
  const rels=parts['word/_rels/document.xml.rels']?xml(parts['word/_rels/document.xml.rels']):null
  const relationships=new Map(Array.from(rels?.documentElement.childNodes||[]).filter(n=>n.nodeType===1).map(n=>[n.getAttribute('Id'),n]))
  const inherited=new Map()
  return sections.map((section,index)=>{
    for(const ref of children(section,'footerReference')){
      const rel=relationships.get(ref.getAttributeNS(R,'id')),target=rel?.getAttribute('Target')
      inherited.set(ref.getAttributeNS(W,'type')||'default',target&&rel.getAttribute('TargetMode')!=='External'?(target.startsWith('/')?target.slice(1):posix.normalize(posix.join('word',target))):null)
    }
    const page=children(section,'pgNumType')[0]
    return {index,format:page?.getAttributeNS(W,'fmt')||'decimal',start:page?.getAttributeNS(W,'start')||null,
      footers:[...inherited].map(([type,path])=>{
        if(!path||!parts[path])return {type,path,valid:false,pageFields:[]}
        const doc=xml(parts[path]),fields=[]
        for(const p of Array.from(doc.getElementsByTagNameNS(W,'p'))){
          const instructions=[...Array.from(p.getElementsByTagNameNS(W,'instrText')).map(n=>n.textContent),...Array.from(p.getElementsByTagNameNS(W,'fldSimple')).map(n=>n.getAttributeNS(W,'instr'))]
          const fonts=Array.from(p.getElementsByTagNameNS(W,'r')).map(r=>{
            const pr=children(r,'rPr')[0],font=children(pr,'rFonts')[0],size=children(pr,'sz')[0]
            return {ascii:font?.getAttributeNS(W,'ascii')||null,hAnsi:font?.getAttributeNS(W,'hAnsi')||null,eastAsia:font?.getAttributeNS(W,'eastAsia')||null,size:size?Number(size.getAttributeNS(W,'val'))/2:null}
          })
          for(const instruction of instructions)if(/^\s*PAGE(?:\s|$)/i.test(instruction))fields.push({instruction,alignment:children(children(p,'pPr')[0],'jc')[0]?.getAttributeNS(W,'val')||null,fonts})
        }
        return {type,path,valid:true,pageFields:fields}
      })}
  })
}
