import test from 'node:test'
import assert from 'node:assert/strict'
import {createServer} from 'node:http'
import {mkdtemp,rm} from 'node:fs/promises'
import {tmpdir} from 'node:os'
import {join} from 'node:path'
import {Agent} from '@earendil-works/pi-agent-core'
import {streamSimple} from '@earendil-works/pi-ai/api/openai-completions'
import {SessionManager} from '../vendor/pi/packages/coding-agent/src/core/session-manager.ts'
import {attachPiSession} from '../server/pi-session.mjs'
import {budgetDocumentContext} from '../server/document/read-context.mjs'

test('native Pi compaction and restored session retain document tools and system rules',async()=>{
  const dir=await mkdtemp(join(tmpdir(),'pi-native-')),requests=[];let mode='history',called=0,runtime
  const endpoint=createServer(async(req,res)=>{
    let body='';for await(const part of req)body+=part
    const payload=JSON.parse(body);requests.push(payload)
    const chunk=(delta,finish_reason=null)=>({id:'native',object:'chat.completion.chunk',created:1,model:'test-model',choices:[{index:0,delta,finish_reason}]})
    res.writeHead(200,{'content-type':'text/event-stream'})
    if(mode==='tool'){
      mode='finish'
      assert.ok(payload.tools.some(t=>t.function.name==='inspect_document'))
      assert.match(JSON.stringify(payload.messages),/PRESERVE_DOCUMENT/)
      res.write(`data: ${JSON.stringify(chunk({role:'assistant',tool_calls:[{index:0,id:'read-restored',type:'function',function:{name:'inspect_document',arguments:'{}'}}]}))}\n\n`)
      res.write(`data: ${JSON.stringify(chunk({},'tool_calls'))}\n\n`)
    }else{
      res.write(`data: ${JSON.stringify(chunk({role:'assistant',content:mode==='summary'?'## Goal\nPreserve document; continue remaining formatting.':mode==='history'?'History '.repeat(1000):'Done'}))}\n\n`)
      res.write(`data: ${JSON.stringify(chunk({},'stop'))}\n\n`)
    }
    res.end('data: [DONE]\n\n')
  })
  await new Promise(resolve=>endpoint.listen(0,'127.0.0.1',resolve))
  try{
    const model={id:'test-model',name:'test',api:'openai-completions',provider:'openai',baseUrl:`http://127.0.0.1:${endpoint.address().port}/v1`,reasoning:false,input:['text'],contextWindow:128000,maxTokens:4096,cost:{input:0,output:0,cacheRead:0,cacheWrite:0}}
    const tools=[{name:'inspect_document',label:'读取',description:'Read document',parameters:{type:'object',properties:{}},execute:async()=>{called++;return {content:[{type:'text',text:'revision 1'}]}}}]
    const make=async manager=>{
      const agent=new Agent({initialState:{model,messages:manager.buildSessionContext().messages},transformContext:budgetDocumentContext,
        streamFn:(m,c,o)=>streamSimple(m,c,{...o,apiKey:'test-only',maxRetries:0})})
      return attachPiSession({agent,sessionManager:manager,agentDir:dir,systemPrompt:'PRESERVE_DOCUMENT',tools,apiKey:'test-only',compaction:{keepRecentTokens:100,reserveTokens:1000}})
    }
    const manager=SessionManager.create(dir,join(dir,'sessions'))
    runtime=await make(manager)
    await runtime.prompt('Preserve original text. Analyze formatting.',{expandPromptTemplates:false})
    await runtime.prompt('Continue analysis.',{expandPromptTemplates:false})
    const compactionEvents=[]
    runtime.subscribe(event=>{if(event.type==='compaction_start')compactionEvents.push(event.reason)})
    await runtime.setModel({...model,contextWindow:3000},{persist:false})
    mode='summary';await runtime.prompt('Continue with existing rules.',{expandPromptTemplates:false})
    assert.ok(compactionEvents.includes('threshold'),'native threshold automatically compacts')
    await runtime.setModel(model,{persist:false})
    mode='history';await runtime.prompt('More analysis before manual compaction.',{expandPromptTemplates:false})
    await runtime.prompt('Record remaining work.',{expandPromptTemplates:false})
    mode='summary';await runtime.compact('Keep original requirements and remaining work.')
    assert.ok(manager.getBranch().some(e=>e.type==='compaction'))
    runtime.dispose()
    runtime=await make(SessionManager.continueRecent(dir,join(dir,'sessions')))
    mode='tool';await runtime.prompt('Read current document before continuing.',{expandPromptTemplates:false})
    assert.equal(called,1)
    assert.deepEqual(runtime.agent.state.tools.map(t=>t.name),['inspect_document'])
    assert.ok(requests.some(p=>!p.tools?.length),'native summary request used')
  }finally{runtime?.dispose();endpoint.closeAllConnections();await new Promise(resolve=>endpoint.close(resolve));await rm(dir,{recursive:true,force:true})}
})
