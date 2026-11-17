// Only allowlisted metadata is persisted; never log URLs, headers or error messages.
export function diagnosticFetch(record,fetcher=globalThis.fetch,{idleMs=90000}={}) {
 const code=e=>{const v=[e?.cause?.code,e?.code,e?.name].find(v=>typeof v==='string');return typeof v==='string'&&/^[A-Za-z0-9_]{1,60}$/.test(v)?v:'UNKNOWN'}
 return async(input,init)=>{
  const start=Date.now(),controller=new AbortController()
  const signal=init?.signal||input?.signal
  const abort=()=>controller.abort(signal?.reason)
  if(signal?.aborted)abort();else signal?.addEventListener('abort',abort,{once:true})
  let timer,consumerClosed=false
  const arm=()=>{clearTimeout(timer);timer=setTimeout(()=>controller.abort(new DOMException('Model stream idle','TimeoutError')),idleMs)}
  const clean=()=>{clearTimeout(timer);signal?.removeEventListener('abort',abort)}
  const emit=async data=>{try{await record({...data,time:new Date().toISOString(),elapsedMs:Date.now()-start})}catch{/* Diagnostics must never turn a successful request into a retry. */}}
  arm()
  try{
   const response=await fetcher(input,{...init,signal:controller.signal})
   await emit({phase:'headers',status:response.status})
   if(!response.body){clean();return response}
   const reader=response.body.getReader()
   const body=new ReadableStream({async pull(target){
    arm()
    try{const next=await reader.read();if(next.done){clean();target.close();await emit({phase:'complete',status:response.status})}else{arm();target.enqueue(next.value)}}
    catch(e){clean();await emit({phase:consumerClosed?'consumer-closed':'stream-error',code:code(controller.signal.aborted?controller.signal.reason:e),cancelled:Boolean(signal?.aborted)});target.error(e)}
   },async cancel(reason){consumerClosed=true;clean();controller.abort(reason);await reader.cancel(reason)}})
   return new Response(body,{status:response.status,statusText:response.statusText,headers:response.headers})
  }catch(e){clean();await emit({phase:'connect-error',code:code(controller.signal.aborted?controller.signal.reason:e),cancelled:Boolean(signal?.aborted)});throw e}
 }
}
