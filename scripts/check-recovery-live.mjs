import {cp,mkdir,readFile,writeFile,readdir} from 'node:fs/promises'
import {resolve} from 'node:path'
import {once} from 'node:events'
const id='ff3617e9-770e-4d7b-82cb-b5ba73b0286b',root=resolve('output/recovery-live-'+Date.now())
await mkdir(root,{recursive:true});await cp('data/documents/'+id,root+'/'+id,{recursive:true})
process.env.PI_DOCX_DATA_DIR=root
const {startServer}=await import('../server/api.mjs')
const server=startServer({port:0});await once(server,'listening')
const base='http://127.0.0.1:'+server.address().port
await writeFile('output/recovery-live-location.json',JSON.stringify({root,id,base}))
try{
 const meta=await(await fetch(base+'/api/documents/'+id)).json()
 const files=await readdir('data/documents/'+id+'/agent/sessions');let original=''
 for(const f of files)for(const line of (await readFile('data/documents/'+id+'/agent/sessions/'+f,'utf8')).trim().split('\n')){const m=JSON.parse(line).message;if(m?.role==='user')original+=(original?'\n':'')+(typeof m.content==='string'?m.content:m.content.filter(c=>c.type==='text').map(c=>c.text).join('\n'))}
 const response=await fetch(base+'/api/documents/'+id+'/chat',{method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify({revision:meta.revision,message:original+'\n\n用户已授权按常见毕业论文格式整理，不需要再次索要学校规范。采用宋体小四、英文Times New Roman、1.5倍行距、首行缩进24磅；一级黑体三号、二级黑体四号、三级黑体小四，标题无首行缩进；表格文字五号、宽度适应版心，纯图片段落居中。只对明确识别的区域执行，封面保留。继续上次未完成的论文排版要求。从当前版本读取并核对，只处理未完成的正文、标题、表格、图片、目录和页码要求，不重复已提交的题注转换及目录标题分离。按区域分批读取并执行，保留原文内容。最后验证并明确未完成项。'})})
 let transcript='',pending='';const decoder=new TextDecoder()
 for await(const chunk of response.body){const text=decoder.decode(chunk,{stream:true});transcript+=text;pending+=text;const frames=pending.split('\n\n');pending=frames.pop();for(const frame of frames){const type=frame.match(/^event: (.+)$/m)?.[1];if(['tool','document','error','done'].includes(type))console.log(frame.slice(0,700))}await writeFile(root+'/events.txt',transcript)}
 const final=await(await fetch(base+'/api/documents/'+id)).json()
 const bytes=await(await fetch(base+'/api/documents/'+id+'/file')).arrayBuffer();await writeFile(root+'/result.docx',Buffer.from(bytes))
 await writeFile(root+'/result.json',JSON.stringify({revision:final.revision,task:JSON.parse(await readFile(root+'/'+id+'/last-task.json','utf8'))},null,2))
 console.log('LIVE_FINISHED revision='+final.revision)
}finally{server.closeAllConnections();server.close()}
