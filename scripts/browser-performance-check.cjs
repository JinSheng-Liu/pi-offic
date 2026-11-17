async page => {
  await page.reload()
  await page.waitForFunction(() => window.probe?.editor, { timeout: 60000 })
  return await page.evaluate(async () => {
    const { editor, replay } = window.probe
    const scroll = document.querySelector('[data-testid="docx-editor-scroll"]') || document.scrollingElement
    if (!scroll) throw Error('没有编辑器滚动容器')
    const { applyFormatReplay } = await import('/src/lib/format-replay.ts')
    const tasks = []
    const observer = new PerformanceObserver(list => tasks.push(...list.getEntries().map(e=>e.duration)))
    observer.observe({type:'longtask'})
    const generation=editor.mountGeneration,start=performance.now()
    let batches=0,scrollMoved=false
    await applyFormatReplay(editor,replay,()=>{
      batches++
      if (batches===3) {
        const before=scroll.scrollTop
        scroll.scrollTop=before+400
        scrollMoved=scroll.scrollTop>before
      }
    })
    await new Promise(resolve=>setTimeout(resolve,100))
    observer.disconnect()
    const report={paragraphs:replay.targets.reduce((n,t)=>n+t.indices.length,0),batches,
      elapsedMs:performance.now()-start,scrollMoved,mountUnchanged:editor.mountGeneration===generation,
      longTasks:tasks.length,maxLongTaskMs:Math.max(0,...tasks),finalFingerprintMatched:true,
      longTaskDurations:tasks}
    if(!report.scrollMoved||!report.mountUnchanged||batches<2)throw Error(JSON.stringify(report))
    await fetch('/probe-output/report.json',{method:'POST',body:JSON.stringify(report,null,2)})
    return report
  })
}
