import {streamSimple} from '@earendil-works/pi-ai/api/openai-completions'
import {diagnosticFetch} from '../server/model-transport.mjs'
import {writeFile} from 'node:fs/promises'
if(!process.env.OPENAI_API_KEY){console.log('MODEL_CREDENTIALS_NOT_AVAILABLE');process.exit(2)}
const events=[]
const model={id:process.env.PI_MODEL||'gpt-4o-mini',name:'probe',api:'openai-completions',provider:'openai',baseUrl:process.env.PI_API_BASE_URL||'https://api.openai.com/v1',reasoning:false,input:['text'],cost:{input:0,output:0,cacheRead:0,cacheWrite:0},contextWindow:128000,maxTokens:100}
const stream=streamSimple(model,{messages:[{role:'user',content:'请只回复：连接测试通过',timestamp:Date.now()}]},{apiKey:process.env.OPENAI_API_KEY,maxTokens:100,maxRetries:1,timeoutMs:30000,fetch:diagnosticFetch(e=>events.push(e),globalThis.fetch,{idleMs:30000})})
const result=await stream.result()
const report={stopReason:result.stopReason,hasText:result.content.some(c=>c.type==='text'&&c.text.length),events}
await writeFile('output/model-transport-live.json',JSON.stringify(report,null,2));console.log(JSON.stringify(report))
