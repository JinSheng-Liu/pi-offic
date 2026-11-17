import test from 'node:test'
import assert from 'node:assert/strict'
import {diagnosticFetch} from '../server/model-transport.mjs'
test('transport records status and safe cause without secrets',async()=>{
 const events=[],fetch=diagnosticFetch(e=>events.push(e),async()=>{throw Object.assign(new Error('secret token'),{cause:{code:'ECONNRESET'}})})
 await assert.rejects(fetch('https://secret.example',{}))
 assert.equal(events[0].code,'ECONNRESET');assert.doesNotMatch(JSON.stringify(events),/secret/)
})
test('diagnostic write failure cannot fail successful streaming response',async()=>{
 const fetch=diagnosticFetch(()=>{throw Error('disk')},async()=>new Response('ok'))
 assert.equal(await(await fetch('unused',{})).text(),'ok')
})
test('idle stream aborts and records timeout separately from user cancellation',async()=>{
 const events=[]
 const fetch=diagnosticFetch(e=>events.push(e),async(_,init)=>new Response(new ReadableStream({start(c){init.signal.addEventListener('abort',()=>c.error(init.signal.reason))}})),{idleMs:20})
 await assert.rejects((await fetch('unused',{})).text())
 assert.equal(events.at(-1).code,'TimeoutError');assert.equal(events.at(-1).cancelled,false)
})
