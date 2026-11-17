import test from 'node:test'
import assert from 'node:assert/strict'
import {createChecklist,resumeChecklist,updateChecklist,runToTaskBoundary} from '../server/task-checklist.mjs'

test('whole-document work remains pending after a partial model summary',async()=>{
  const checklist=planned(),messages=[];let rev=1
  await runToTaskBoundary({checklist,request:'整理',revision:async()=>rev,hasCommits:()=>rev>1,persist:async()=>{},prompt:async message=>{
    messages.push(message);rev++
    const targets=messages.length===1?checklist.items.slice(0,1):checklist.items.filter(i=>i.state==='pending')
    updateChecklist(checklist,targets.map(i=>({...i,state:'addressed',revision:rev,reason:'已检查并提交对应范围'})),rev)
  }})
  assert.equal(messages.length,2);assert.match(messages[1],/不重复已提交/);assert.equal(checklist.stopReason,'scope-addressed')
})
test('blocked region does not stop independent pending regions',async()=>{
  const checklist=planned();let turns=0
  await runToTaskBoundary({checklist,request:'整理',revision:async()=>1,hasCommits:()=>false,persist:async()=>{},prompt:async()=>{
    turns++;const targets=turns===1?checklist.items.slice(0,1):checklist.items.filter(i=>i.state==='pending')
    updateChecklist(checklist,targets.map(i=>({...i,state:turns===1?'blocked':'notApplicable',revision:1,reason:'实际结构检查依据',evidenceId:'failed-tool',target:'目录条目'})),1,[{toolCallId:'failed-tool',toolName:'apply_formatting',isError:true,revision:1}])
  }})
  assert.equal(turns,2);assert.equal(checklist.stopReason,'blocked')
})
test('repeated summaries stop with an explicit no-progress reason',async()=>{
  const checklist=planned();let turns=0
  await runToTaskBoundary({checklist,request:'整理',revision:async()=>1,hasCommits:()=>false,persist:async()=>{},prompt:async()=>{turns++}})
  assert.equal(turns,3);assert.equal(checklist.stopReason,'no-progress');assert.ok(checklist.items.every(i=>i.state==='pending'))
})
test('ledger preserves omitted pending items and rejects stale or unexplained closure',()=>{
  const checklist=planned(),first=checklist.items[0]
  assert.throws(()=>updateChecklist(checklist,[{...first,state:'addressed',revision:1,reason:'checked'}],2),/当前版本/)
  assert.throws(()=>updateChecklist(checklist,[{...first,state:'blocked',revision:2}],2),/具体检查/)
  updateChecklist(checklist,[{...first,state:'addressed',revision:2,reason:'checked'}],2)
  assert.equal(checklist.items.length,2);assert.equal(checklist.items[1].state,'pending')
})
test('read-only conversations do not trigger unsolicited work',async()=>{
  const checklist=createChecklist('这个字体是什么');let turns=0
  await runToTaskBoundary({checklist,request:'咨询',revision:async()=>1,hasCommits:()=>false,persist:async()=>{},prompt:async()=>turns++})
  assert.equal(turns,1)
  assert.equal(createChecklist('只分析全文排版问题，不需要修改').items.length,0)
})

test('cancellation prevents another model call and leaves remaining items intact',async()=>{
  const checklist=planned(),controller=new AbortController();let turns=0
  await assert.rejects(runToTaskBoundary({checklist,request:'整理',revision:async()=>1,hasCommits:()=>false,persist:async()=>{},signal:controller.signal,prompt:async()=>{turns++;controller.abort()}}),{name:'AbortError'})
  assert.equal(turns,1);assert.ok(checklist.items.every(i=>i.state==='pending'))
})

test('continuation limit bounds changing work without declaring completion',async()=>{
  const checklist=planned();let rev=1
  await runToTaskBoundary({checklist,request:'整理',revision:async()=>rev,hasCommits:()=>true,persist:async()=>{},maxContinuations:2,prompt:async()=>rev++})
  assert.equal(rev,4);assert.equal(checklist.stopReason,'continuation-limit')
})

function planned(){
 const checklist=createChecklist('全文排版')
 updateChecklist(checklist,[{id:'body',label:'用户指定的正文格式',state:'pending'},{id:'toc',label:'实际目录格式',state:'pending'}],1)
 return checklist
}
test('new tasks have no fixed document regions and cannot silently finish without a plan',async()=>{
 const checklist=createChecklist('全文排版');assert.deepEqual(checklist.items,[])
 let turns=0
 await runToTaskBoundary({checklist,request:'全文排版',revision:async()=>1,hasCommits:()=>false,persist:async()=>{},prompt:async()=>turns++})
 assert.equal(turns,3);assert.equal(checklist.stopReason,'no-progress')
})
test('unexamined work cannot be marked blocked without current tool failure evidence',()=>{
 const checklist=planned(),item={...checklist.items[0],state:'blocked',reason:'尚未全量检查',revision:2,target:'正文',evidenceId:'fake'}
 assert.throws(()=>updateChecklist(checklist,[item],2),/真实失败/)
 assert.throws(()=>updateChecklist(checklist,[item],2,[{toolCallId:'fake',toolName:'apply_formatting',isError:true,revision:1}]),/真实失败/)
 assert.equal(checklist.items[0].state,'pending')
 updateChecklist(checklist,[{...item,state:'awaitingVerification'}],2)
 assert.equal(checklist.items[0].state,'awaitingVerification')
})
test('awaiting verification continues rather than ending as blocked',async()=>{
 const checklist=createChecklist('修改正文');let turns=0
 updateChecklist(checklist,[{id:'body',label:'正文字体',state:'awaitingVerification',reason:'已提交，待回读',revision:2}],2)
 await runToTaskBoundary({checklist,request:'检查',revision:async()=>2,hasCommits:()=>true,persist:async()=>{},prompt:async()=>{
  turns++;if(turns===2)updateChecklist(checklist,[{...checklist.items[0],state:'addressed',reason:'回读字体符合要求',revision:2}],2)
 }})
 assert.equal(turns,2)
})
test('legacy fixed scopes are not reintroduced on resume',()=>{
 assert.deepEqual(resumeChecklist({items:[{id:'scope-1',label:'页面与分节',state:'blocked'}]},'继续全文排版').items,[])
 const c=planned();c.items[0].state='blocked'
 assert.equal(resumeChecklist(c,'继续').items[0].state,'pending')
})
