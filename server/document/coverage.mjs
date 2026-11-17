// Evidence inventory, not inferred execution history or visual approval.
export function documentCoverage(data){
  const paragraphs=data.paragraphs||[]
  const eligible=paragraphs.filter(p=>!p.tableNestingLevel&&!/^TOC/i.test(p.style||''))
  const rules=[['cover',/^(?:本科毕业论文[（(]设计[）)]|毕业论文|硕士学位论文|博士学位论文)$/],['authorization',/^学位论文版权使用授权书$/],['declaration',/^学位论文原创性声明$/],['abstract',/^摘要$/],['englishAbstract',/^ABSTRACT$/i],['toc',/^目录$/],['references',/^参考文献$/],['acknowledgments',/^致谢$/],['appendix',/^附录(?:[A-Z一二三四五六七八九十\d]+)?$/]]
  const regions=rules.map(([region,pattern])=>{
    const matches=eligible.filter(p=>pattern.test(p.text.replace(/\s/g,'')))
    return {region,status:matches.length===1?'identified':matches.length?'ambiguous':'not-identified',headingIndices:matches.map(p=>p.index),requirementsVerified:false,visualVerified:false}
  })
  const headings=eligible.filter(p=>p.heading>0&&!rules.some(([,pattern])=>pattern.test(p.text.replace(/\s/g,''))))
  regions.push({region:'body',status:headings.length?'candidate':'not-identified',headingIndices:headings.map(p=>p.index),requirementsVerified:false,visualVerified:false})
  const objects=[{region:'tables',count:data.tables?.length||0},{region:'images',count:data.elements?.images||0},{region:'captions',count:data.captions?.length||0}].map(item=>({...item,status:item.count?'present-unverified':'not-detected',requirementsVerified:false,visualVerified:false}))
  const blockers=(data.floatingObjects||[]).filter(o=>o.textBoxCount&&o.flowCapability?.supported===false).map(o=>({objectIndex:o.objectIndex,reason:o.flowCapability.reason}))
  return {regions,objects,blockedTextboxes:blockers,requirementsVerified:false,visualVerified:false,note:'清单只提供识别证据，不是任务完成记录；未识别不等于缺失，候选标题不确定区域结束边界。按用户范围逐项核对，不能用成功操作数代表全文完成。'}
}
