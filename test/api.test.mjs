import test from 'node:test'
import assert from 'node:assert/strict'
import { createServer } from 'node:http'
import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

// Run the actual Pi provider against a deterministic local Chat Completions endpoint.
test('HTTP model stream completes, releases task lock and persists conversation', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'pi-api-'))
  process.env.PI_DOCX_DATA_DIR = directory
  let mode = 'text'
  let callCount = 0
  let releaseStream
  const model = createServer(async (req, res) => {
    let requestBody = ''
    for await (const chunk of req) requestBody += chunk
    const request = JSON.parse(requestBody)
    assert.ok(request.tools?.some(tool=>tool.function.name==='inspect_document'),'document requests must include actual tool declarations')
    assert.ok(request.tools?.some(tool=>tool.function.name==='update_task_checklist'))
    if (mode === 'commit-fail') {
      if(callCount++===0){
        res.writeHead(200,{'content-type':'text/event-stream'})
        const chunk={id:'partial',object:'chat.completion.chunk',created:1,model:'test-model',choices:[{index:0,delta:{role:'assistant',tool_calls:[{index:0,id:'partial-write',type:'function',function:{name:'apply_formatting',arguments:JSON.stringify({revision:1,operations:[{scope:'sections',sectionIndices:[0],margins:{top:1000}}]})}}]},finish_reason:null}]}
        res.write(`data: ${JSON.stringify(chunk)}\n\n`);chunk.choices[0].delta={};chunk.choices[0].finish_reason='tool_calls';res.write(`data: ${JSON.stringify(chunk)}\n\n`);res.end('data: [DONE]\n\n');return
      }
      req.socket.destroy();return
    }
    if (mode === 'resume-check') {assert.match(JSON.stringify(request.messages.filter(m=>m.role==='user').at(-1).content),/已提交版本：2/)}
    if (mode === 'disconnect') { req.socket.destroy(); return }
    if (mode === 'slow') { res.writeHead(200, { 'content-type': 'text/event-stream' }); res.flushHeaders(); return }
    res.writeHead(200, { 'content-type': 'text/event-stream' })
    const chunk = (delta, finish_reason = null) => ({ id: 'test', object: 'chat.completion.chunk', created: 1, model: 'test-model', choices: [{ index: 0, delta, finish_reason }] })
    if(mode==='blocked-evidence'){
      const step=callCount++
      if(step===2)assert.match(JSON.stringify(request.messages.at(-1)),/真实失败/)
      if(step===4){
        const result=JSON.parse(request.messages.filter(m=>m.role==='tool').at(-1).content)
        assert.equal(result.items[0].state,'blocked')
        assert.equal(result.items[0].failure.toolCallId,'evidence-2')
      }
      if(step<4){
        const item={id:'body',label:'指定正文格式',state:'pending'}
        const name=step===2?'apply_formatting':'update_task_checklist'
        const args=step===2?{revision:1,operations:[{scope:'indices',indices:[999],font:{size:12}}]}:{items:[step===0?item:{...item,state:'blocked',revision:1,reason:'目标索引无效，需要重新定位',target:'段落999',evidenceId:step===1?'invented':'evidence-2'}]}
        res.write(`data: ${JSON.stringify(chunk({role:'assistant',tool_calls:[{index:0,id:`evidence-${step}`,type:'function',function:{name,arguments:JSON.stringify(args)}}]}))}\n\n`)
        res.write(`data: ${JSON.stringify(chunk({},'tool_calls'))}\n\n`);res.end('data: [DONE]\n\n');return
      }
    }
    if(mode==='checklist'){
      const step=callCount++
      if(step===1){
        assert.match(JSON.stringify(request.messages),/宿主任务检查/)
        const args={items:[{id:'body',label:'读取当前正文并确定排版范围',state:'notApplicable',revision:1,reason:'测试空白文档无文字'}]}
        res.write(`data: ${JSON.stringify(chunk({role:'assistant',tool_calls:[{index:0,id:'task-scope',type:'function',function:{name:'update_task_checklist',arguments:JSON.stringify(args)}}]}))}\n\n`)
        res.write(`data: ${JSON.stringify(chunk({},'tool_calls'))}\n\n`);res.end('data: [DONE]\n\n');return
      }
    }
    if (mode === 'staged') {
      const gate = new Promise(resolve => { releaseStream = resolve })
      res.write(`data: ${JSON.stringify(chunk({ role: 'assistant', content: '第一段' }))}\n\n`)
      await gate
      res.write(`data: ${JSON.stringify(chunk({ content: '第二段' }))}\n\n`)
      res.write(`data: ${JSON.stringify(chunk({}, 'stop'))}\n\n`)
      res.end('data: [DONE]\n\n'); return
    }
    if (mode === 'skills') {
      const step = callCount++
      if (step === 0) {
        const system = request.messages.find(m => m.role === 'system').content
        assert.match(system, /docx-academic/)
        assert.doesNotMatch(system, /不存在通用的论文字号组合/)
      }
      if (step === 1) {
        const result = JSON.parse(request.messages.filter(m => m.role === 'tool').at(-1).content)
        assert.equal(result.name, 'docx-academic')
        assert.match(result.content, /不存在通用的论文字号组合/)
      }
      if (step === 2) {
        const result = JSON.parse(request.messages.filter(m => m.role === 'tool').at(-1).content)
        assert.equal(result.view, 'summary')
        assert.equal(result.items.length, 0)
        assert.equal(result.revision, 3)
      }
      if (step < 2) {
        const name = step === 0 ? 'load_document_skill' : 'inspect_document'
        const args = step === 0 ? {name:'docx-academic'} : {view:'summary'}
        res.write(`data: ${JSON.stringify(chunk({role:'assistant',tool_calls:[{index:0,id:`skill-${step}`,type:'function',function:{name,arguments:JSON.stringify(args)}}]}))}\n\n`)
        res.write(`data: ${JSON.stringify(chunk({}, 'tool_calls'))}\n\n`)
        res.end('data: [DONE]\n\n'); return
      }
    }
    if (mode === 'repair' && callCount < 2) {
      const repairing = callCount++ === 1
      if (repairing) {
        const error = request.messages.filter(m => m.role === 'tool').at(-1)
        assert.match(error.content, /INVALID_FORMAT_ARGUMENTS/)
        assert.match(error.content, /"committed":false/)
        assert.match(error.content, /拆成独立的段落操作/)
      }
      const section = {scope:'sections',sectionIndices:[0],page:{widthTwips:11906,heightTwips:16838},margins:{top:1440,left:1800}}
      const args = {revision:1,operations: repairing ? [section,{scope:'indices',indices:[0],paragraph:{spaceBefore:0,spaceAfter:0}}] : [{...section,paragraph:{spaceBefore:0,spaceAfter:0}}]}
      res.write(`data: ${JSON.stringify(chunk({role:'assistant',tool_calls:[{index:0,id:repairing?'repair':'invalid',type:'function',function:{name:'apply_formatting',arguments:JSON.stringify(args)}}]}))}\n\n`)
      res.write(`data: ${JSON.stringify(chunk({}, 'tool_calls'))}\n\n`)
      res.end('data: [DONE]\n\n'); return
    }
    if (mode === 'replay' && callCount++ === 0) {
      const args={revision:1,operations:[{scope:'all',paragraph:{lineSpacingMultiple:1.5},font:{size:12}}]}
      res.write(`data: ${JSON.stringify(chunk({role:'assistant',tool_calls:[{index:0,id:'replay-format',type:'function',function:{name:'apply_formatting',arguments:JSON.stringify(args)}}]}))}\n\n`)
      res.write(`data: ${JSON.stringify(chunk({},'tool_calls'))}\n\n`)
      res.end('data: [DONE]\n\n');return
    }
    if (mode === 'tool' && callCount++ === 0) {
      const args = { revision: 2, operations: [{ scope: 'sections', sectionIndices: [0], margins: { top: 900 } }] }
      res.write(`data: ${JSON.stringify(chunk({ role: 'assistant', tool_calls: [{ index: 0, id: 'call_format', type: 'function', function: { name: 'apply_formatting', arguments: JSON.stringify(args) } }] }))}\n\n`)
      res.write(`data: ${JSON.stringify(chunk({}, 'tool_calls'))}\n\n`)
      res.end('data: [DONE]\n\n'); return
    }
    res.write(`data: ${JSON.stringify(chunk({ role: 'assistant', content: '本地模型协议验证通过' }))}\n\n`)
    res.write(`data: ${JSON.stringify(chunk({}, 'stop'))}\n\n`)
    res.end('data: [DONE]\n\n')
  })
  await new Promise(resolve => model.listen(0, '127.0.0.1', resolve))
  const { createApiMiddleware } = await import('../server/api.mjs')
  const server = createServer(createApiMiddleware())
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve))
  const base = `http://127.0.0.1:${server.address().port}`
  const jsonRequest = async (path, method = 'GET', body) => {
    const response = await fetch(base + path, { method, headers: { 'content-type': 'application/json' }, body: body ? JSON.stringify(body) : undefined })
    return { status: response.status, body: await response.json() }
  }
  try {
    const setting = await jsonRequest('/api/settings/model', 'PUT', { apiKey: 'test-only', model: 'test-model', baseUrl: `http://127.0.0.1:${model.address().port}/v1` })
    assert.equal(setting.status, 200)
    assert.equal(setting.body.apiKey, undefined)
    assert.equal((await fetch(base + '/api/documents', { method: 'POST', headers: { Origin: 'https://foreign.example' } })).status, 403)
    const countBeforeBadUpload = (await jsonRequest('/api/documents')).body.documents.length
    assert.equal((await fetch(base + '/api/documents', {method:'POST',body:new Uint8Array([80,75,3,4])})).status,415)
    assert.equal((await jsonRequest('/api/documents')).body.documents.length,countBeforeBadUpload)
    const created = await jsonRequest('/api/documents', 'POST')
    assert.equal(created.status, 201)
    const path = `/api/documents/${created.body.id}`
    assert.equal((await jsonRequest(path + '/chat', 'POST', { message: '测试', revision: 99 })).status, 409)
    for (let i = 0; i < 2; i++) {
      const response = await fetch(base + path + '/chat', { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ message: '测试流式回复', revision: 1 }) })
      assert.equal(response.status, 200)
      const text = await response.text()
      assert.match(text, /本地模型协议验证通过/)
      assert.match(text, /event: done/)
      assert.doesNotMatch(text, /event: error/)
    }
    const history = await jsonRequest(path + '/chat')
    assert.equal(history.body.messages.length, 4)
    const checklistDoc=await jsonRequest('/api/documents','POST'),checklistPath=`/api/documents/${checklistDoc.body.id}`
    mode='checklist';callCount=0
    const checklistResponse=await fetch(base+checklistPath+'/chat',{method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify({message:'全文排版',revision:1})})
    const checklistEvents=await checklistResponse.text()
    assert.equal(callCount,3);assert.match(checklistEvents,/event: checklist/);assert.doesNotMatch(checklistEvents,/event: error/)
    const checklistResult=(await jsonRequest(checklistPath+'/task')).body.result
    assert.equal(checklistResult.checklist.stopReason,'scope-addressed');assert.equal(checklistResult.checklist.continuations,1)
    mode='text'
    const stalledResponse=await fetch(base+checklistPath+'/chat',{method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify({message:'全文排版',revision:1})})
    const stalledEvents=await stalledResponse.text()
    assert.match(stalledEvents,/no-progress/);assert.match(stalledEvents,/连续检查未产生/)
    assert.equal((await jsonRequest(checklistPath+'/task')).body.result.state,'partial')
    const restoredHistory=(await jsonRequest(checklistPath+'/chat')).body.messages
    assert.equal(restoredHistory.at(-1).steps?.length || 0,0)
    assert.equal((await jsonRequest(checklistPath+'/task')).body.result.checklist.items.length,0)
    assert.ok(restoredHistory.filter(m=>m.role==='user').every(m=>!m.text.startsWith('[宿主任务检查：')))
    mode='blocked-evidence';callCount=0
    const evidenceResponse=await fetch(base+checklistPath+'/chat',{method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify({message:'设置正文格式',revision:1})})
    const evidenceEvents=await evidenceResponse.text()
    assert.doesNotMatch(evidenceEvents,/event: error/)
    const evidenceResult=(await jsonRequest(checklistPath+'/task')).body.result
    assert.equal(evidenceResult.checklist.stopReason,'blocked')
    assert.equal(evidenceResult.checklist.items[0].failure.toolCallId,'evidence-2')
    assert.equal(evidenceResult.currentRevision,1)
    mode = 'staged'
    const staged = await fetch(base + path + '/chat', { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ message: '验证逐段输出', revision: 1 }), signal: AbortSignal.timeout(10000) })
    const reader = staged.body.getReader()
    const decoder = new TextDecoder()
    let received = ''
    try {
      while (!received.includes('第一段')) {
        const part = await reader.read()
        assert.equal(part.done, false)
        received += decoder.decode(part.value, { stream: true })
      }
      assert.doesNotMatch(received, /第二段|event: done/)
    } finally { releaseStream?.() }
    while (true) { const part = await reader.read(); if (part.done) break; received += decoder.decode(part.value, { stream: true }) }
    assert.match(received, /第二段/)
    assert.match(received, /event: done/)
    assert.equal((await jsonRequest(path + '/restore', 'POST', { revision: 1, target: 1 })).body.revision, 2)
    mode = 'tool'; callCount = 0
    const toolResponse = await fetch(base + path + '/chat', { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ message: '调整上边距', revision: 2 }) })
    const toolEvents = await toolResponse.text()
    assert.match(toolEvents, /apply_formatting/)
    assert.match(toolEvents, /event: document/)
    assert.match(toolEvents, /\"revision\":3/)
    assert.match(toolEvents, /修改文档格式/)
    assert.match(toolEvents, /"id":"call_format"/)
    assert.doesNotMatch(toolEvents, /event: error/)
    assert.equal((await jsonRequest(path)).body.revision, 3)
    assert.equal((await jsonRequest(path + '/outline')).body.sections[0].margins.top, 900)
    mode = 'skills'; callCount = 0
    const skillResponse = await fetch(base + path + '/chat', {method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify({message:'读取论文排版技能和结构摘要',revision:3})})
    const skillEvents = await skillResponse.text()
    assert.match(skillEvents, /加载排版技能/)
    assert.match(skillEvents, /event: done/)
    assert.doesNotMatch(skillEvents, /event: document|event: error/)
    assert.equal((await jsonRequest(path)).body.revision, 3)
    const replayDoc=await jsonRequest('/api/documents','POST'),replayPath=`/api/documents/${replayDoc.body.id}`
    mode='replay';callCount=0
    const replayResponse=await fetch(base+replayPath+'/chat',{method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify({message:'修改正文格式并显示',revision:1})})
    const replayEvents=await replayResponse.text()
    const frame=replayEvents.split('\n\n').find(f=>f.startsWith('event: document\n'))
    assert.ok(frame)
    const commit=JSON.parse(frame.split('data: ')[1])
    assert.equal(commit.previousRevision,1);assert.equal(commit.revision,2)
    assert.equal(commit.replay.kind,'format-v1');assert.equal(commit.replay.targets[0].indices[0],0)
    assert.match(commit.replay.before,/^[a-f0-9]{64}$/);assert.match(commit.replay.after,/^[a-f0-9]{64}$/)
    assert.equal((await jsonRequest(replayPath+'/task')).body.result.committedRevisions[0].replay,undefined)
    const repairDoc = await jsonRequest('/api/documents', 'POST')
    const repairPath = `/api/documents/${repairDoc.body.id}`
    mode = 'repair'; callCount = 0
    const repairResponse = await fetch(base + repairPath + '/chat', {method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify({message:'设置页面和段落间距',revision:1})})
    const repairEvents = await repairResponse.text()
    assert.match(repairEvents, /"state":"error"/)
    assert.match(repairEvents, /"id":"repair"/)
    assert.equal((repairEvents.match(/event: document/g) || []).length, 1)
    assert.equal((await jsonRequest(repairPath)).body.revision, 2)
    assert.match(repairEvents, /event: done/)
    const partialDoc=await jsonRequest('/api/documents','POST'),partialPath=`/api/documents/${partialDoc.body.id}`
    mode='commit-fail';callCount=0
    const partial=await fetch(base+partialPath+'/chat',{method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify({message:'写入后模拟断连',revision:1})})
    const partialEvents=await partial.text()
    assert.match(partialEvents,/event: task_result/)
    assert.match(partialEvents,/event: error/)
    const receipt=(await jsonRequest(partialPath+'/task')).body.result
    assert.equal(receipt.state,'failed');assert.equal(receipt.currentRevision,2);assert.deepEqual(receipt.committedRevisions.map(v=>v.revision),[2])
    assert.equal((await jsonRequest(partialPath+'/outline')).body.sections[0].margins.top,1000)
    mode='resume-check'
    const resumed=await fetch(base+partialPath+'/chat',{method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify({message:'继续检查',revision:2})})
    assert.doesNotMatch(await resumed.text(),/event: error/)
    assert.equal((await jsonRequest(partialPath)).body.revision,2)
    mode = 'disconnect'
    const failed = await fetch(base + path + '/chat', { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ message: '连接失败测试', revision: 3 }) })
    const failedEvents = await failed.text()
    assert.match(failedEvents, /event: error/)
    assert.doesNotMatch(failedEvents, /event: done/)
    assert.equal((await jsonRequest(path)).body.revision, 3)
    mode = 'slow'
    const slow = await fetch(base + path + '/chat', { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ message: '等待取消', revision: 3 }) })
    assert.equal((await jsonRequest(path + '/chat', 'POST', { message: '并发任务', revision: 3 })).status, 409)
    assert.equal((await jsonRequest(path + '/restore', 'POST', { revision: 3, target: 1 })).status, 409)
    assert.equal((await jsonRequest(path + '/cancel', 'POST')).status, 200)
    const cancelled=await slow.text()
    assert.match(cancelled,/event: done/)
    assert.match(cancelled,/cancelled/)
    assert.equal((await jsonRequest(path + '/restore', 'POST', { revision: 3, target: 1 })).body.revision, 4)
  } finally {
    releaseStream?.()
    server.closeAllConnections(); model.closeAllConnections()
    await Promise.all([new Promise(resolve => server.close(resolve)), new Promise(resolve => model.close(resolve))])
    await rm(directory, { recursive: true, force: true })
  }
})
