import test from 'node:test'
import assert from 'node:assert/strict'
import {paragraphDocument,documentFixture,paragraph,table} from './helpers/docx-fixture.mjs'
import {format,inspect} from '../server/engine.mjs'
import {executeFormatPlan} from '../server/document/format-plan.mjs'
import {unpack,xml,W} from '../server/document/package.mjs'
import {auditPlan} from '../server/document/plan-audit.mjs'
import {taskResult} from '../server/task-result.mjs'
import {DocumentStore} from '../server/store.mjs'
import {mkdtemp,rm} from 'node:fs/promises'
import {tmpdir} from 'node:os'
import {join} from 'node:path'
import {atomicReplace} from '../server/atomic-replace.mjs'

test('atomic metadata replacement retries only temporary Windows locks with a finite bound',async()=>{
  let attempts=0;const waits=[]
  const locked=Object.assign(Error('locked'),{code:'EPERM'})
  await atomicReplace('temporary','metadata',{platform:'win32',wait:async ms=>waits.push(ms),move:async()=>{if(++attempts<3)throw locked}})
  assert.equal(attempts,3);assert.deepEqual(waits,[25,50])
  attempts=0
  await assert.rejects(atomicReplace('temporary','metadata',{platform:'win32',wait:async()=>{},move:async()=>{attempts++;throw locked}}),/locked/)
  assert.equal(attempts,5)
  attempts=0
  await assert.rejects(atomicReplace('temporary','metadata',{platform:'win32',move:async()=>{attempts++;throw Object.assign(Error('missing'),{code:'ENOENT'})}}),/missing/)
  assert.equal(attempts,1)
})

test('progress delivery failure cannot relabel or repeat a committed write',async()=>{
  let writes=0
  const bytes=(await format(paragraphDocument(['正文']),[{scope:'indices',indices:[0],font:{size:12}}])).bytes
  const store={async apply(){writes++;return {document:{revision:2},affected:1}},async bytes(){return bytes}}
  const report=await executeFormatPlan(store,'d',1,[{label:'字号',operations:[{scope:'indices',indices:[0],font:{size:12}}]}],undefined,progress=>{if(progress.steps[0].state==='committed')throw Error('连接已关闭')})
  assert.equal(writes,1);assert.equal(report.steps[0].state,'committed');assert.equal(report.audit.planTargetsVerified,true)
  assert.ok(report.progressDeliveryFailures>0)
})

test('postcommit audit failure returns retained commits and explicit unverified result',async()=>{
  const store={async apply(){return {document:{revision:2},affected:1}},async bytes(){throw Error('读取暂时失败')}}
  const report=await executeFormatPlan(store,'d',1,[{label:'字号',operations:[{scope:'indices',indices:[0],font:{size:12}}]}])
  assert.equal(report.currentRevision,2);assert.equal(report.steps[0].state,'committed')
  assert.equal(report.state,'partial');assert.equal(report.audit.planTargetsVerified,false);assert.match(report.audit.reason,/读取暂时失败/)
})

test('cancellation after final commit retains that commit and marks audit unverified',async()=>{
  const controller=new AbortController()
  const store={async apply(){return {document:{revision:2},affected:1}},async bytes(){throw Error('不应读取')}}
  const report=await executeFormatPlan(store,'d',1,[{label:'字号',operations:[{scope:'indices',indices:[0],font:{size:12}}]}],controller.signal,progress=>{if(progress.steps[0].state==='committed')controller.abort()})
  assert.equal(report.state,'cancelled');assert.equal(report.steps[0].state,'committed')
  assert.equal(report.currentRevision,2);assert.equal(report.audit.planTargetsVerified,false)
})

test('table text scope covers all cells without formatting surrounding body; prose is not a caption',async()=>{
  const source=documentFixture('<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>图5.1为系统界面，下面说明实现过程。</w:t></w:r></w:p>'+paragraph('图5.1 系统界面')+table([['字段','说明'],['id','标识']]))
  const result=await format(source,[{scope:'tableText',font:{name:'宋体',latinName:'Times New Roman',size:10.5}}])
  const data=await inspect(result.bytes)
  assert.equal(data.paragraphs[0].font.size,12)
  assert.ok(data.paragraphs.filter(p=>p.tableNestingLevel).every(p=>p.font.size===10.5))
  assert.deepEqual(data.captions.map(p=>p.text),['图5.1 系统界面'])
})
test('plan commits serially, reports pending work on failure, never claims requirements passed',async()=>{
  const seen=[]
  const store={async apply(id,revision){seen.push(revision);if(revision===3)throw Error('内容保护拒绝');return {document:{revision:revision+1},affected:2}}}
  const steps=Array.from({length:4},(_,i)=>({label:`步骤${i}`,operations:[{scope:'indices',indices:[0],font:{size:12}}]}))
  const report=await executeFormatPlan(store,'document',1,steps)
  assert.deepEqual(seen,[1,2,3]);assert.equal(report.state,'partial')
  assert.deepEqual(report.steps.map(s=>s.state),['committed','committed','failed','pending'])
  assert.equal(report.requirementsVerified,false)
})
test('plan validates every step before first commit and rejects stale-index structural plans',async()=>{
  let writes=0;const store={apply(){writes++}}
  await assert.rejects(()=>executeFormatPlan(store,'d',1,[{label:'转换',operations:[{scope:'convertCaptions',indices:[1]}]}]),/结构转换/)
  assert.equal(writes,0)
})
test('page numbers create a footer without duplicates, retaining body text',async()=>{
  let bytes=paragraphDocument(['正文保留'])
  for(let i=0;i<2;i++)bytes=(await format(bytes,[{scope:'ensurePageNumbers'}])).bytes
  const parts=unpack(bytes),footers=Object.entries(parts).filter(([n])=>/^word\/piFooter\d+\.xml$/.test(n))
  assert.equal(footers.length,1)
  const footer=xml(footers[0][1]);assert.equal(footer.getElementsByTagNameNS(W,'fldSimple').length,1)
  assert.equal(footer.getElementsByTagNameNS(W,'jc')[0].getAttributeNS(W,'val'),'center')
  assert.equal((await inspect(bytes)).paragraphs[0].text,'正文保留')
})
test('table fitting coordinates grid and cell widths and is repeatable',async()=>{
  let bytes=documentFixture(table([['字段','内容'],['a','b']]))
  for(let i=0;i<2;i++)bytes=(await format(bytes,[{scope:'fitTableWidth',tableIndex:0}])).bytes
  const doc=xml(unpack(bytes)['word/document.xml'])
  const attr=(n,k)=>Number(n.getAttributeNS(W,k))
  const total=Array.from(doc.getElementsByTagNameNS(W,'gridCol')).reduce((sum,n)=>sum+attr(n,'w'),0)
  assert.equal(attr(doc.getElementsByTagNameNS(W,'tblW')[0],'w'),total)
  assert.equal(Array.from(doc.getElementsByTagNameNS(W,'tcW')).slice(0,2).reduce((sum,n)=>sum+attr(n,'w'),0),total)
})
test('final audit catches earlier target regression and accepts intentional later overrides',async()=>{
  const bytes=documentFixture('<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>正文</w:t></w:r></w:p>')
  const steps=[{label:'字号',operations:[{scope:'indices',indices:[0],font:{size:10.5}}]}]
  const failed=await auditPlan(bytes,steps)
  assert.equal(failed.planTargetsVerified,false);assert.equal(failed.failures.length,1)
  steps.push({label:'用户后续字号',operations:[{scope:'indices',indices:[0],font:{size:12}}]})
  assert.equal((await auditPlan(bytes,steps)).planTargetsVerified,true)
})

test('formatting plan evidence remains auditable after an unrelated later revision', async () => {
  const source = documentFixture('<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>')
  const first = await format(source, [{scope:'indices', indices:[0], font:{size:10.5}}])
  const second = await format(first.bytes, [{scope:'indices', indices:[1], font:{size:12}}])
  const audit = await auditPlan(second.bytes, [{label:'第一段', operations:[{scope:'indices',indices:[0],font:{size:10.5}}]}])
  assert.equal(audit.planTargetsVerified, true)
})

test('adjacent table width steps share one document commit', async () => {
  const source = documentFixture('<w:tbl><w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:tbl><w:tr><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>')
  let writes = 0
  const store = { bytes: async () => source, apply: async (_id, revision, operations) => { writes++; return { document: { revision: revision + 1 }, affected: operations.length, validation: {} } }, withStep: (_label, run) => run() }
  const report = await executeFormatPlan(store, 'id', 1, [
    { label:'表1', operations:[{ scope:'fitTableWidth', tableIndex:0 }] },
    { label:'表2', operations:[{ scope:'fitTableWidth', tableIndex:1 }] },
  ], undefined, async () => {})
  assert.equal(writes, 1)
  assert.equal(report.steps.length, 1)
})
test('normal model completion cannot imply validated formatting completion',()=>{
  const result=taskResult('complete',1,2,[{revision:2,summary:'修改'}])
  assert.equal(result.state,'partial');assert.equal(result.conversationState,'complete');assert.equal(result.requirementsVerified,false)
})

test('empty table paragraphs are audited through their paragraph mark font',async()=>{
  const bytes=documentFixture(table([['内容','']]))
  const operations=[{scope:'tableText',font:{name:'宋体',latinName:'Times New Roman',size:10.5}}]
  const result=await format(bytes,operations)
  const audit=await auditPlan(result.bytes,[{label:'表格',operations}])
  assert.equal(audit.planTargetsVerified,true)
})
test('real store plan publishes every committed revision and audits final bytes',async()=>{
  const directory=await mkdtemp(join(tmpdir(),'plan-regression-'))
  try{
    const store=new DocumentStore(directory)
    const doc=await store.create('plan.docx',paragraphDocument(['真实计划'])),commits=[]
    store.onCommit(commit=>commits.push(commit.revision))
    const report=await executeFormatPlan(store,doc.id,1,[{label:'字号',operations:[{scope:'indices',indices:[0],font:{size:16}}]},{label:'对齐',operations:[{scope:'indices',indices:[0],paragraph:{alignment:'Centered'}}]}])
    assert.deepEqual(commits,[2,3]);assert.equal(report.audit.planTargetsVerified,true)
    assert.equal(report.currentRevision,3);assert.equal(report.requirementsVerified,false)
  }finally{await rm(directory,{recursive:true,force:true})}
})

test('TOC page tab can be safely reapplied after a page-width change',async()=>{
  let bytes=documentFixture('<w:p><w:r><w:t>第一章 绪论</w:t><w:tab/><w:t>1</w:t></w:r></w:p>')
  for(const position of [420,400,400])bytes=(await format(bytes,[{scope:'indices',indices:[0],paragraph:{tocPageTabPosition:position}}])).bytes
  const data=await inspect(bytes)
  assert.equal(data.paragraphs[0].tabStops.direct[0].positionTwips,8000)
  assert.equal(data.paragraphs[0].text,'第一章 绪论\t1')
})

test('final audit verifies leading and TOC tabs and detects later displacement',async()=>{
  for(const key of ['leadingTabPosition','tocPageTabPosition']){
    const source=documentFixture(paragraph(key==='leadingTabPosition'?'\t正文':'第一章\t1'))
    const operations=[{scope:'indices',indices:[0],paragraph:{[key]:key==='leadingTabPosition'?24:400}}]
    const steps=[{label:'制表位',operations}]
    const result=await format(source,operations)
    assert.equal((await auditPlan(result.bytes,steps)).planTargetsVerified,true)
    const changed=await format(source,[{scope:'indices',indices:[0],paragraph:{[key]:100}}])
    assert.equal((await auditPlan(changed.bytes,steps)).failures.length,1)
  }
})

test('footer audit checks final alignment and fonts and later explicit overrides',async()=>{
  const source=(await format(paragraphDocument(['正文']),[{scope:'ensurePageNumbers'}])).bytes
  const operations=[{scope:'footerPageNumber',sectionIndices:[0],alignment:'Centered',font:{name:'Times New Roman',size:10.5}}]
  const steps=[{label:'页码',operations}]
  const result=await format(source,operations)
  assert.equal((await auditPlan(result.bytes,steps)).planTargetsVerified,true)
  const override=[{scope:'footerPageNumber',sectionIndices:[0],alignment:'Right',font:{size:12}}]
  const changed=await format(result.bytes,override)
  assert.equal((await auditPlan(changed.bytes,steps)).failures.length,1)
  assert.equal((await auditPlan(changed.bytes,[...steps,{label:'新要求',operations:override}])).planTargetsVerified,true)
})

test('final audit detects column proportions overwritten after fitting', async () => {
  const source = documentFixture(table([['问题','原因','影响'],['甲','乙','丙']]))
  const operations = [{scope:'fitTableWidth',tableIndex:0,columnRatios:[24,38,38]}]
  const result = await format(source, operations)
  const steps = [{label:'调整列宽',operations}]
  const good = await auditPlan(result.bytes, steps)
  assert.equal(good.planTargetsVerified, true)
  const changed = await format(result.bytes,[{scope:'fitTableWidth',tableIndex:0,columnRatios:[1,1,1]}])
  const bad = await auditPlan(changed.bytes, steps)
  assert.ok(bad.failures.some(check=>check.target.endsWith('columnRatios')))
  assert.equal(bad.planTargetsVerified, false)
})
