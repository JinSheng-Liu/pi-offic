async page => {
  const result = await page.evaluate(async () => {
    if (!(await (await fetch('/api/health')).json()).testMode) throw new Error('仅允许隔离测试')
    let v = document.querySelector('.editor-frame').firstElementChild.__vueParentComponent
    while (v && !v.exposed?.getEditor) v = v.parent
    const editor = v.exposed.getEditor()
    const paragraphs = editor.surface.layout().pages.flatMap(p => p.fragments.filter(f => f.kind === 'paragraph').map(f => ({page:p.index+1, id:f.paragraphId, text:f.lines.flatMap(l=>l.spans.map(s=>s.text)).join('')})))
    const checks = ['一、工作目标','六、组织保障'].map(text => {
      const i = paragraphs.findIndex(p=>p.text===text)
      if (i < 0 || !paragraphs[i+1] || paragraphs[i].page !== paragraphs[i+1].page) throw new Error('标题与正文不同页：'+text)
      return {text,page:paragraphs[i].page}
    })
    const {documents} = await (await fetch('/api/documents')).json()
    const doc = documents.filter(d=>d.name==='公文排版能力验收.docx').sort((a,b)=>b.updatedAt.localeCompare(a.updatedAt))[0]
    if (document.querySelector('.document-name').textContent !== doc.name) throw new Error('当前不是验收副本')
    const response = await fetch(`/api/documents/${doc.id}/file`,{method:'PUT',headers:{'content-type':'application/vnd.openxmlformats-officedocument.wordprocessingml.document','x-document-revision':String(doc.revision)},body:await v.exposed.save()})
    if(!response.ok) throw new Error('编辑器保存失败')
    const saved = await response.json()
    editor.scrollToBlock(paragraphs.find(p=>p.text==='一、工作目标').id)
    return {checks,id:doc.id,revision:saved.revision}
  })
  await page.screenshot({path:'data/validation/layout/editor-layout.png'})
  return result
}
