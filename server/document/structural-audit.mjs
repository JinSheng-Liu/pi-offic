import {unpack,xml,W} from './package.mjs'
import {inspect} from './structure.mjs'
import {hasNativeToc} from './toc-layout.mjs'

// These obligations come from actual structures in this document, not fixed document categories.
export async function structuralTargets(bytes, request) {
  if (/只.*(?:检查|分析)|不(?:要|需要).*修改/.test(request)) return []
  const full = /(?:整理|排版).*(?:文档|全文)|全文.*(?:整理|排版)/.test(request)
  const data = await inspect(bytes), targets = []
  if ((full || /(?:生成|更新|真正|正式|规范).*目录|目录.*(?:生成|更新)/.test(request)) &&
      (hasNativeToc(bytes) || data.paragraphs.some(p=>/^(目\s*录|contents)$/i.test(p.text.trim())))) {
    targets.push({id:'structure-native-toc',label:'生成并核验文档中的真实目录',kind:'toc'})
  }
  if (full || /表格|Tab|制表符/i.test(request)) {
    for (let index=0;index<data.paragraphs.length;index++) {
      const header=data.paragraphs[index]
      if(header.tableNestingLevel || !/^(问题\t形成原因\t影响|指标\t公开数据\t分析含义)$/.test(header.text.trim()))continue
      const texts=[header.text]
      while(data.paragraphs[index+1]?.tableNestingLevel===0 && data.paragraphs[index+1].text.split('\t').length===3)texts.push(data.paragraphs[++index].text)
      if(texts.length>1)targets.push({id:`structure-table-${header.index}`,label:`转换并排版“${header.text.replaceAll('\t','／')}”数据表`,kind:'table',cells:texts.flatMap(text=>text.split('\t'))})
    }
  }
  return targets
}

export async function auditStructures(bytes, targets) {
  const data=await inspect(bytes)
  const document=xml(unpack(bytes)['word/document.xml'])
  const tables=[...document.getElementsByTagNameNS(W,'body')[0].getElementsByTagNameNS(W,'tbl')]
  const children=(node,name)=>[...node.childNodes].filter(n=>n.namespaceURI===W&&n.localName===name)
  return targets.map(target=>{
    if(target.kind==='toc')return {id:target.id,passed:hasNativeToc(bytes),reason:'检查真实TOC域；普通文字目录不满足此项'}
    const table=data.tables.find(table=>{
      const cells=children(tables[table.index],'tr').flatMap(row=>children(row,'tc').map(cell=>[...cell.getElementsByTagNameNS(W,'t')].map(n=>n.textContent).join('')))
      return JSON.stringify(cells)===JSON.stringify(target.cells)
    })
    const ratios=table?.layoutRecommendation?.widthOperation?.columnRatios
    const widths=table?.gridWidthsTwips || [],total=widths.reduce((a,b)=>a+b,0)
    const sum=ratios?.reduce((a,b)=>a+b,0)
    const fitted=ratios && widths.length===ratios.length && widths.every((width,index)=>Math.abs(width-total*ratios[index]/sum)<ratios.length)
    return {id:target.id,passed:Boolean(table&&fitted),reason:'检查真实表格、完整单元格内容及分类列宽'}
  })
}
