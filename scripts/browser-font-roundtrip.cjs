async page => {
  return page.evaluate(async () => {
    if (!(await (await fetch('/api/health')).json()).testMode) throw new Error('仅允许隔离测试')
    let component = document.querySelector('.editor-frame').firstElementChild.__vueParentComponent
    while (component && !component.exposed?.getEditor) component = component.parent
    if (!component?.exposed?.save) throw new Error('没有找到编辑器保存接口')
    const { documents } = await (await fetch('/api/documents')).json()
    const doc = documents.filter(d => d.name === '公文字体验收副本.docx').sort((a,b) => b.updatedAt.localeCompare(a.updatedAt))[0]
    if (!doc || doc.revision < 2) throw new Error('真实模型尚未生成测试版本')
    const bytes = await component.exposed.save()
    const response = await fetch(`/api/documents/${doc.id}/file`, { method:'PUT', headers: { 'content-type':'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'x-document-revision':String(doc.revision) }, body:bytes })
    if (!response.ok) throw new Error(`编辑器保存失败 ${response.status}`)
    const saved = await response.json()
    return { id:saved.id, revision:saved.revision }
  })
}
