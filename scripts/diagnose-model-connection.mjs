import {mkdtemp,readFile,writeFile,mkdir,rm} from 'node:fs/promises'
import {tmpdir} from 'node:os'
import {join} from 'node:path'
import {createServer} from 'node:http'
const temp=await mkdtemp(join(tmpdir(),'pi-connection-'));process.env.PI_DOCX_DATA_DIR=temp
const nativeFetch=globalThis.fetch;const traces=[];let lastRequest;let sequence=0
const endpoint=new URL(process.env.PI_API_BASE_URL)
globalThis.fetch=async(input,init)=>{
 const url=typeof input==='string'?input:input instanceof URL?input.href:input.url
 if(new URL(url).origin!==endpoint.origin)return nativeFetch(input,init)
 const record={request:++sequence,start:new Date().toISOString(),bytes:typeof init?.body==='string'?Buffer.byteLength(init.body):null};traces.push(record)
 lastRequest={input,init};const started=Date.now()
 try{
  const response=await nativeFetch(input,init);record.status=response.status;record.headersMs=Date.now()-started
  if(!response.body)return response
  const reader=response.body.getReader();let size=0;let cancelled=false
  const body=new ReadableStream({async pull(controller){
   let r
   try{r=await reader.read()}catch(e){if(cancelled)return;record.failurePhase='stream';record.error=e.name;record.code=e.cause?.code;controller.error(e);return}
   if(cancelled)return
   if(r.done){record.streamBytes=size;record.totalMs=Date.now()-started;record.complete=true;controller.close()}
   else{size+=r.value.byteLength;controller.enqueue(r.value)}
  },cancel(reason){cancelled=true;record.cancelledByConsumer=true;record.streamBytes=size;record.totalMs=Date.now()-started;return reader.cancel(reason)}})
  return new Response(body,{status:response.status,statusText:response.statusText,headers:response.headers})
 }catch(e){record.failurePhase='connect';record.error=e.name;record.code=e.cause?.code;record.totalMs=Date.now()-started;throw e}
}
const {createApiMiddleware}=await import('../server/api.mjs');const server=createServer(createApiMiddleware());await new Promise(r=>server.listen(0,'127.0.0.1',r));const base='http://127.0.0.1:'+server.address().port
try{
 const bytes=await readFile('D:/xiangmu/agent-office/agent-office/测试文档/党政机关公文_无格式版.docx')
 const doc=await(await nativeFetch(base+'/api/documents?name=connection-test.docx',{method:'POST',headers:{'content-type':'application/vnd.openxmlformats-officedocument.wordprocessingml.document'},body:bytes})).json()
 const response=await nativeFetch(base+`/api/documents/${doc.id}/chat`,{method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify({revision:1,message:'请先调用 inspect_document 读取文档，然后调用 analyze_formatting 分析全文格式，最后只简短报告段落数量和当前字体情况。这是连接诊断，不修改文档。'}),signal:AbortSignal.timeout(150000)})
 const stream=await response.text();const events=stream.split('\n').filter(l=>l.startsWith('event: '));console.log(JSON.stringify({phase:'product',events,traces}))
 if(lastRequest){const start=Date.now();try{const r=await nativeFetch(lastRequest.input,{...lastRequest.init,signal:AbortSignal.timeout(60000)});const text=await r.text();traces.push({phase:'directReplaySamePayload',status:r.status,totalMs:Date.now()-start,done:text.includes('[DONE]'),bytes:Buffer.byteLength(text)});}catch(e){traces.push({phase:'directReplaySamePayload',error:e.name,code:e.cause?.code})}}
 const passed=response.ok&&events.includes('event: done')&&!events.includes('event: error')&&!traces.some(t=>t.error)&&traces.some(t=>t.phase==='directReplaySamePayload'&&t.status===200&&t.done)
 await mkdir('data/validation',{recursive:true});await writeFile('data/validation/connection-diagnostic.json',JSON.stringify({passed,events,traces},null,2));console.log(JSON.stringify({phase:'complete',passed,traces}));if(!passed)process.exitCode=1
}finally{server.closeAllConnections();await new Promise(r=>server.close(r));await rm(temp,{recursive:true,force:true})}
