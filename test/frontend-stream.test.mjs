import test from 'node:test'
import assert from 'node:assert/strict'
import {consumeEvents} from '../src/lib/api.ts'
import {createEditorEventHandler} from '../src/lib/editor-events.ts'

test('SSE preserves Chinese across byte boundaries and CRLF events',async()=>{
 const bytes=new TextEncoder().encode('event: text\r\ndata: {"delta":"中文输出"}\r\n\r\nevent: done\ndata: {"state":"complete"}\n\n')
 const stream=new ReadableStream({start(c){for(const b of bytes)c.enqueue(new Uint8Array([b]));c.close()}})
 const events=[];await consumeEvents(new Response(stream),(event,data)=>events.push({event,data}))
 assert.deepEqual(events,[{event:'text',data:{delta:'中文输出'}},{event:'done',data:{state:'complete'}}])
})
test('SSE propagates HTTP errors instead of presenting success',async()=>{
 await assert.rejects(()=>consumeEvents(new Response(JSON.stringify({error:'服务不可用'}),{status:503}),()=>{}),/服务不可用/)
})

test('editor queue preserves mutation order after failure while UI events bypass it',async()=>{
 let release
 const gate=new Promise(resolve=>{release=resolve}),seen=[]
 const handler=createEditorEventHandler(async(event,data)=>{
  if(event==='text'){seen.push('text');return}
  seen.push(`start-${data.id}`)
  if(data.id===1){await gate;throw Error('预览失败')}
  seen.push(`end-${data.id}`)
 })
 const first=handler('document',{id:1})
 const rejected=assert.rejects(first,/预览失败/)
 const second=handler('editor_request',{id:2})
 const third=handler('document',{id:3})
 await handler('text',{})
 assert.ok(seen.includes('text'))
 assert.equal(seen.includes('start-2'),false)
 release()
 await rejected
 await Promise.all([second,third])
 assert.deepEqual(seen.filter(x=>x!=='text'),['start-1','start-2','end-2','start-3','end-3'])
})

test('SSE does not block later UI events behind a slow document callback',async()=>{
 const encoder=new TextEncoder()
 const stream=new ReadableStream({start(controller){
  controller.enqueue(encoder.encode('event: document\ndata: {"revision":2}\n\n'))
  controller.enqueue(encoder.encode('event: text\ndata: {"delta":"已继续输出"}\n\n'))
  controller.enqueue(encoder.encode('event: done\ndata: {"state":"complete"}\n\n'))
  controller.close()
 }})
 let releaseSlowDocument
 const slowDocument=new Promise(resolve=>{releaseSlowDocument=resolve})
 const events=[]
 const consuming=consumeEvents(new Response(stream),async(event)=>{
  events.push(`${event}:start`)
  if(event==='document') await slowDocument
  events.push(`${event}:end`)
 })
 await new Promise(resolve=>setTimeout(resolve,0))
 try {
  assert.deepEqual(events,['document:start','text:start','text:end','done:start','done:end'])
 } finally { releaseSlowDocument() }
 await consuming
 assert.equal(events.at(-1),'document:end')
})

test('SSE observes handler failure immediately but waits for in-flight updates before rejecting',async()=>{
 let release
 const gate=new Promise(resolve=>{release=resolve})
 const response=new Response('event: document\ndata: {}\n\nevent: error\ndata: {}\n\n')
 let finished=false
 const consuming=consumeEvents(response,async event=>{
  if(event==='document'){await gate;finished=true}
  else throw Error('模型失败')
 })
 const rejected=assert.rejects(consuming,/模型失败/)
 await new Promise(resolve=>setTimeout(resolve,0))
 assert.equal(finished,false)
 release()
 await rejected
 assert.equal(finished,true)
})

test('SSE network failure drains pending preview work before recovery',async()=>{
 let release,transport
 const gate=new Promise(resolve=>{release=resolve})
 const stream=new ReadableStream({start(controller){transport=controller}})
 let started=false,finished=false,settled=false
 const consuming=consumeEvents(new Response(stream),async()=>{
  started=true;await gate;finished=true
 }).finally(()=>{settled=true})
 const rejected=assert.rejects(consuming,/连接断开/)
 transport.enqueue(new TextEncoder().encode('event: document\ndata: {}\n\n'))
 await new Promise(resolve=>setTimeout(resolve,0))
 assert.equal(started,true)
 transport.error(Error('连接断开'))
 await new Promise(resolve=>setTimeout(resolve,0))
 assert.equal(settled,false)
 release()
 await rejected
 assert.equal(finished,true)
 assert.equal(stream.locked,false)
})
