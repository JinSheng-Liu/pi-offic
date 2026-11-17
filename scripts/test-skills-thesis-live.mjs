import {readFile,writeFile,mkdir} from 'node:fs/promises'
import {createHash} from 'node:crypto'
const base='http://127.0.0.1:5175',dir=process.env.PI_TEST_OUTPUT || 'data/validation/skills-thesis'
const source=process.env.PI_TEST_SOURCE || 'C:/Users/Administrator/Desktop/基于Spring Boot 的校园二手交易平台的设计与实现_去除格式.docx'
const health=await(await fetch(base+'/api/health')).json();if(!health.testMode)throw Error('非测试服务')
const bytes=await readFile(source);const hash=createHash('sha256').update(bytes).digest('hex')
await mkdir(dir,{recursive:true})
const r=await fetch(base+'/api/documents?name='+encodeURIComponent('无格式论文_Skills真实测试.docx'),{method:'POST',headers:{'content-type':'application/vnd.openxmlformats-officedocument.wordprocessingml.document'},body:bytes});if(r.status!==201)throw Error('上传失败 '+r.status)
const doc=await r.json();await writeFile(dir+'/task.json',JSON.stringify({id:doc.id,source,hash},null,2))
const message=process.env.PI_TEST_PROMPT || '请实际排版当前整份论文，这是功能测试，以下数值是测试要求，不代表学校规范。先按需加载论文及相关对象技能。1.封面“本科毕业论文（设计）”设黑体30pt居中，题目两行设黑体18pt，学生等信息行保留原格式，授权书、原创性声明和签署区保持原样。2.摘要标题黑体16pt居中，ABSTRACT标题Times New Roman 16pt居中；中英文摘要正文12pt、1.5倍行距，中文宋体、西文Times New Roman，首行缩进24pt；关键词保持原样。3.正文区域明确的一级章节标题黑体16pt、二级14pt、三级12pt，西文Times New Roman，设置对应Heading层级并与下段同页；绝不能把目录条目设为Heading。普通正文中文宋体12pt、西文Times New Roman、1.5倍行距、首行缩进24pt，排除目录、图表、题注、参考文献及特殊区域。4.正文内独占段落的嵌入式图片居中，不修改尺寸和浮动图位置；不要改封面或声明区图像。5.表内原文不改，明确的表头文字居中，其他单元格保持原样；不新增题注。6.保留目录域与所有文字、图片、表格、空段，不要求重算目录页码或清除空白。请真实执行并逐项回读，未完成或不能确认的范围明确报告，不用成功提交代替全部需求验收。'
await writeFile(dir+'/prompt.txt',message,'utf8')
console.log(JSON.stringify({id:doc.id,state:'uploaded',bytes:bytes.length}))
let events='';const started=Date.now()
try{
 const response=await fetch(base+`/api/documents/${doc.id}/chat`,{method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify({revision:1,message}),signal:AbortSignal.timeout(600000)})
 const decoder=new TextDecoder();for await(const chunk of response.body){events+=decoder.decode(chunk,{stream:true});await writeFile(dir+'/events.txt',events,'utf8')}
 console.log(JSON.stringify({state:'stream-ended',elapsedSeconds:Math.round((Date.now()-started)/1000),error:events.includes('event: error')}))
}catch(e){console.log(JSON.stringify({state:'request-failed',message:e.message}));await fetch(base+`/api/documents/${doc.id}/cancel`,{method:'POST'})}
const meta=await(await fetch(base+`/api/documents/${doc.id}`)).json()
const out=await(await fetch(base+`/api/documents/${doc.id}/file`)).arrayBuffer();await writeFile(dir+'/测试结果.docx',new Uint8Array(out))
await writeFile(dir+'/result-meta.json',JSON.stringify(meta,null,2))
console.log(JSON.stringify({id:doc.id,revision:meta.revision,originalUnchanged:createHash('sha256').update(await readFile(source)).digest('hex')===hash}))
