import {unpack,xml,W,canonical} from './package.mjs'
import {referenceRegion} from './reference-cover.mjs'
const text=n=>Array.from(n.getElementsByTagNameNS(W,'t')).map(t=>t.textContent||'').join('')
function styleSignature(styles,id,seen=new Set()){
  if(seen.has(id))return 'cycle'
  const style=Array.from(styles?.getElementsByTagNameNS(W,'style')||[]).find(s=>s.getAttributeNS(W,'styleId')===id)?.cloneNode(true)
  if(!style)return 'missing:'+id
  const chain=new Set(seen);chain.add(id);style.removeAttributeNS(W,'styleId')
  for(const tag of ['name','next','link'])for(const n of Array.from(style.getElementsByTagNameNS(W,tag)))n.parentNode.removeChild(n)
  for(const n of Array.from(style.getElementsByTagNameNS(W,'basedOn')))n.setAttributeNS(W,'w:val',styleSignature(styles,n.getAttributeNS(W,'val'),chain))
  return canonical(style)
}
const props=(n,name,styles)=>Array.from(n.getElementsByTagNameNS(W,name)).map(x=>{
  const clone=x.cloneNode(true)
  for(const tag of ['pStyle','rStyle'])for(const ref of Array.from(clone.getElementsByTagNameNS(W,tag)))ref.setAttributeNS(W,'w:val',styleSignature(styles,ref.getAttributeNS(W,'val')))
  return canonical(clone)
})
export function auditFrontMatter(bytes,reference){
  const a=unpack(bytes),b=unpack(reference),target=xml(a['word/document.xml']),sample=xml(b['word/document.xml']),as=a['word/styles.xml']?xml(a['word/styles.xml']):null,bs=b['word/styles.xml']?xml(b['word/styles.xml']):null
  return {visualVerified:false,regions:['cover','authorization','declaration'].map(region=>{
    let current,expected
    try{current=referenceRegion(target,region).blocks}catch(error){return {region,status:'unresolved',reason:error.message}}
    try{expected=referenceRegion(sample,region).blocks}catch(error){return {region,status:'reference-unresolved',reason:error.message}}
    if(current.length!==expected.length||current.some((p,i)=>text(p)!==text(expected[i])))return {region,status:'content-mismatch',reason:'段落数量或原文不同，需要映射，不能复制参考内容'}
    const differences=[]
    for(let i=0;i<current.length;i++){
      const paragraph=JSON.stringify(props(current[i],'pPr',as))!==JSON.stringify(props(expected[i],'pPr',bs))
      const runs=JSON.stringify(props(current[i],'rPr',as))!==JSON.stringify(props(expected[i],'rPr',bs))
      if(paragraph||runs)differences.push({paragraphOffset:i,text:text(current[i]).slice(0,60),paragraphPropertiesDiffer:paragraph,runPropertiesDiffer:runs})
    }
    return {region,status:differences.length?'format-differences':'reference-properties-matched',differences,visualVerified:false}
  }),note:'比较结果仅表示参考属性差异；未比较最终字体替代和页面呈现，也不等同于任务执行历史。'}
}
