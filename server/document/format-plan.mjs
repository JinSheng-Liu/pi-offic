import {runDocumentJob} from './document-jobs.mjs'
import {validateOperations} from './rules.mjs'
import {problem} from './package.mjs'

// Structural changes invalidate following paragraph/object indices. They remain
// separate tools: inspect the resulting document before planning the next phase.
const structural=new Set(['floatingObjects','tocTitle','coverLayout','convertCaptions','tabularText','numberEquations','renumberHeadings'])
function mergeTableWidthSteps(steps) {
  const merged=[]
  for(const step of steps){
    const onlyTables=step.operations.length>0&&step.operations.every(op=>op.scope==='fitTableWidth')
    const previous=merged.at(-1)
    if(onlyTables&&previous?.operations.every(op=>op.scope==='fitTableWidth')) {
      previous.operations.push(...step.operations)
      previous.labels.push(step.label)
    } else merged.push({label:step.label,labels:[step.label],operations:[...step.operations]})
  }
  return merged
}
export async function executeFormatPlan(store,id,revision,steps,signal,onProgress=()=>{}){
  if(!Array.isArray(steps)||!steps.length||steps.length>100)throw problem('计划需要1–100个步骤')
  for(const step of steps){
    if(typeof step.label!=='string'||!step.label.trim()||step.label.length>160)throw problem('计划步骤需要简短名称')
    validateOperations(step.operations)
    if(step.operations.some(op=>structural.has(op.scope)))throw problem('结构转换必须独立执行并重读索引后再创建格式计划')
  }
  const executionSteps=mergeTableWidthSteps(steps)
  const report={inputRevision:revision,currentRevision:revision,state:'running',requirementsVerified:false,visualVerified:false,
    steps:executionSteps.map(s=>({label:s.labels.length>1?`${s.labels[0]}等${s.labels.length}项`:s.label,state:'pending',mergedLabels:s.labels.length>1?s.labels:undefined}))}
  const publish=async()=>{
    try{await onProgress(report)}catch{report.progressDeliveryFailures=(report.progressDeliveryFailures||0)+1}
  }
  for(let i=0;i<executionSteps.length;i++){
    try{
      signal?.throwIfAborted()
      report.steps[i].state='running';await publish()
      const apply=()=>store.apply(id,report.currentRevision,executionSteps[i].operations,signal)
      const result=await (store.withStep?store.withStep(executionSteps[i].label,apply):apply())
      report.currentRevision=result.document.revision
      Object.assign(report.steps[i],{state:'committed',revision:report.currentRevision,affected:result.affected})
      await publish()
    }catch(error){
      Object.assign(report.steps[i],{state:signal?.aborted?'cancelled':'failed',reason:error.message})
      report.state=signal?.aborted?'cancelled':'partial';await publish()
      return report
    }
  }
  report.state='executed'
  try{
    signal?.throwIfAborted()
    report.audit=await runDocumentJob('audit',{bytes:await store.bytes(id,report.currentRevision),steps:executionSteps},signal)
    signal?.throwIfAborted()
  }catch(error){
    report.state=signal?.aborted?'cancelled':'partial'
    report.audit={planTargetsVerified:false,requirementsVerified:false,visualVerified:false,reason:error.message,
      note:'步骤已提交，最终验收未完成；请重新验收当前版本，不要重复执行已提交步骤。'}
  }
  await publish()
  return report
}
