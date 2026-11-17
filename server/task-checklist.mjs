// A task ledger records coverage, not independent proof of visual correctness.
export function createChecklist(request){
  const requiresPlan=!/不(?:要|需要).*修改|只.*(?:分析|检查|解释)/.test(request)&&/(?:排版|整理|统一|修改|设置|规范)/.test(request)
  return {schemaVersion:2,requiresPlan,items:[],continuations:0,stopReason:null}
}

export function resumeChecklist(previous,request){
  if(previous?.schemaVersion===2){
    const next=structuredClone(previous)
    next.continuations=0;next.stopReason=null
    for(const item of next.items)if(['blocked','needsDecision'].includes(item.state)){
      item.state='pending';delete item.evidenceId;delete item.question
    }
    return next
  }
  // Legacy fixed scopes are history, not a new task plan. Re-read committed work.
  return createChecklist(request)
}

export function updateChecklist(checklist,items,currentRevision,evidence=[]){
  if(!Array.isArray(items)||!items.length||items.length>60)throw Error('任务清单需要 1–60 项')
  const next=structuredClone(checklist.items),seen=new Set()
  for(const item of items){
    if(!item.id||seen.has(item.id)||!item.label?.trim()||!['pending','awaitingVerification','addressed','blocked','needsDecision','notApplicable'].includes(item.state))throw Error('任务项 ID、名称或状态无效')
    seen.add(item.id)
    const old=next.find(i=>i.id===item.id)
    if(old&&old.label!==item.label)throw Error('已有任务项不能改名；请新增补充项')
    if(item.state!=='pending'&&(!item.reason?.trim()||item.revision!==currentRevision))throw Error('关闭任务项必须提供当前版本和具体检查依据或阻塞原因')
    if(item.state==='blocked'){
      const failure=evidence.find(e=>e.toolCallId===item.evidenceId&&e.isError&&e.revision===currentRevision&&e.toolName!=='update_task_checklist')
      if(!failure||!item.target?.trim())throw Error('受阻必须关联当前版本真实失败的工具调用 evidenceId，并限定具体 target；未检查应保持 pending 或 awaitingVerification')
    }
    if(item.state==='needsDecision'&&!item.question?.trim())throw Error('需要确认必须给出具体 question，不能用未检查替代')
    const entry={id:item.id,label:item.label,state:item.state,...(item.state!=='pending'?{reason:item.reason,revision:item.revision}:{})}
    if(item.state==='blocked')Object.assign(entry,{evidenceId:item.evidenceId,target:item.target,failure:structuredClone(evidence.find(e=>e.toolCallId===item.evidenceId))})
    if(item.state==='needsDecision')entry.question=item.question
    if(old)next[next.indexOf(old)]=entry;else next.push(entry)
  }
  if(next.length>60)throw Error('任务清单最多 60 项')
  checklist.items=next
  return checklist
}

export async function runToTaskBoundary({prompt,request,checklist,revision,hasCommits,signal,persist,maxContinuations=12}){
  let message=request,previous='',stagnant=0,previousEvidence=''
  for(;;){
    signal?.throwIfAborted()
    await prompt(message)
    signal?.throwIfAborted()
    const currentRevision=await revision()
    const pending=checklist.items.filter(i=>['pending','awaitingVerification'].includes(i.state))
    const missing=(checklist.requiresPlan||hasCommits())&&!checklist.items.length
    if(!pending.length&&!missing){checklist.stopReason=checklist.items.some(i=>i.state==='blocked')?'blocked':checklist.items.some(i=>i.state==='needsDecision')?'needs-decision':'scope-addressed';await persist();return}
    const evidence=checklist.items.map(i=>[i.id,i.state,i.revision||null,i.evidenceId||null,i.reason||''])
      .sort((a,b)=>a[0].localeCompare(b[0]))
    const effectiveEvidence=JSON.stringify(checklist.items.map(i=>[i.id,i.revision||null,i.evidenceId||null]).sort((a,b)=>a[0].localeCompare(b[0])))
    // State labels may be corrected from addressed back to awaitingVerification
    // without any new document work. Only a new revision or new evidence counts
    // as progress; otherwise automatic continuation must stop.
    const signature=JSON.stringify([currentRevision,effectiveEvidence])
    stagnant=signature===previous?stagnant+1:0
    previous=signature;previousEvidence=effectiveEvidence
    if(stagnant>=2||checklist.continuations>=maxContinuations){
      checklist.stopReason=stagnant>=2?'no-progress':'continuation-limit';await persist();return
    }
    checklist.continuations++
    await persist()
    message=`[宿主任务检查：原需求尚未处理完，继续同一任务，不是新需求。当前版本 ${currentRevision}。${missing?'尚未建立清单，请先用 update_task_checklist 登记原需求全部范围。':`待处理项：${JSON.stringify(pending)}`}请读取当前文档，仅处理剩余项目，不重复已提交操作。完成一个范围后更新清单，提供当前版本和检查依据；已修改待检查用awaitingVerification，未检查不能标受阻。局部受阻必须关联真实工具失败及具体目标，继续其他独立项。没有修改权限的范围不得修改；不适用需说明依据。不要仅总结后结束。]`
  }
}
