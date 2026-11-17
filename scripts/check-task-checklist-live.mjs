import {mkdir,writeFile} from 'node:fs/promises'
import {resolve} from 'node:path'
import {once} from 'node:events'
import {paragraphDocument} from '../test/helpers/docx-fixture.mjs'
import {inspect} from '../server/engine.mjs'
const root=resolve('output/task-checklist-live-'+Date.now())
await mkdir(root,{recursive:true});process.env.PI_DOCX_DATA_DIR=root
const {startServer}=await import('../server/api.mjs')
const server=startServer({port:0});await once(server,'listening')
const base='http://127.0.0.1:'+server.address().port
let timer
try{
  const source=paragraphDocument(['测试说明','这是第一段正文。','This is the second paragraph.'])
  const created=await fetch(base+'/api/documents?name=checklist-smoke.docx',{method:'POST',headers:{'content-type':'application/vnd.openxmlformats-officedocument.wordprocessingml.document'},body:source})
  if(!created.ok)throw Error('上传失败')
  const doc=await created.json(),path=base+'/api/documents/'+doc.id
  timer=setTimeout(()=>{void fetch(path+'/cancel',{method:'POST'})},180000)
  const response=await fetch(path+'/chat',{method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify({revision:1,message:'请统一全文格式：全部三个段落均作为普通正文，中文宋体、英文Times New Roman，12磅，1.5倍行距。保留文字，不添加任何内容；页面、页眉页脚维持现状。本文没有封面、目录、图表和参考文献，不需要添加。请建立任务清单，执行并检查。'})})
  if(!response.ok)throw Error('模型请求失败：'+response.status)
  let events=''
  for await(const chunk of response.body){events+=new TextDecoder().decode(chunk);}
  const task=(await(await fetch(path+'/task')).json()).result
  const bytes=Buffer.from(await(await fetch(path+'/file')).arrayBuffer()),data=await inspect(bytes)
  const report={taskState:task.state,conversationState:task.conversationState,revision:task.currentRevision,checklist:task.checklist,
    textPreserved:data.paragraphs.map(p=>p.text).join('\n')===(await inspect(source)).paragraphs.map(p=>p.text).join('\n'),
    fontVerified:data.paragraphs.every(p=>p.font.size===12&&p.font.effectiveFonts.every(f=>f.eastAsia==='宋体'&&f.ascii==='Times New Roman'&&f.hAnsi==='Times New Roman')),
    spacingVerified:data.paragraphs.every(p=>p.lineSpacing.rule==='auto'&&p.lineSpacing.multiple===1.5)}
  await writeFile(root+'/report.json',JSON.stringify(report,null,2));await writeFile(root+'/events.txt',events);await writeFile(root+'/result.docx',bytes)
  console.log(JSON.stringify({root,...report}))
  if(!report.textPreserved||!report.fontVerified||!report.spacingVerified||task.checklist.stopReason!=='scope-addressed')process.exitCode=1
}finally{clearTimeout(timer);server.closeAllConnections();server.close()}
