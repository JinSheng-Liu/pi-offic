import{readFile,writeFile,mkdir}from'node:fs/promises';import assert from'node:assert/strict';import{inspect}from'../server/engine.mjs';
const base='http://127.0.0.1:5175',dir='data/validation/tab-fix';await mkdir(dir,{recursive:true});assert.equal((await(await fetch(base+'/api/health')).json()).testMode,true);
const original=await readFile('data/validation/skills-thesis/测试结果.docx');const before=await inspect(original);
const upload=await fetch(base+'/api/documents?name='+encodeURIComponent('论文Tab定位回归.docx'),{method:'POST',headers:{'content-type':'application/vnd.openxmlformats-officedocument.wordprocessingml.document'},body:original});assert.equal(upload.status,201);const doc=await upload.json();await writeFile(dir+'/task.json',JSON.stringify(doc));
const indices=[126,127,128,129,130,131,132,162];
const message='请实际修改当前论文中索引126、127、128、129、130、131、132、162这8段：先定点读取完整格式，确认它们是单个前导Tab、无已有自定义制表位及缩进。保留所有文字和Tab，将Tab后文字的第一个起点设为24pt，使用leadingTabPosition，不把它声称为firstLineIndent属性。其他字体、行距和所有其他段落保持不变。修改后定点回读tabStops和lineSpacing，报告实际数值，不声称页面视觉已验收。';
const started=Date.now();const response=await fetch(base+`/api/documents/${doc.id}/chat`,{method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify({revision:1,message}),signal:AbortSignal.timeout(240000)});let events='';for await(const chunk of response.body){events+=new TextDecoder().decode(chunk);await writeFile(dir+'/events.txt',events)}
const bytes=new Uint8Array(await(await fetch(base+`/api/documents/${doc.id}/file`)).arrayBuffer());await writeFile(dir+'/论文_Tab定位修正.docx',bytes);const after=await inspect(bytes);
assert.deepEqual(after.paragraphs.map(p=>p.text),before.paragraphs.map(p=>p.text));
for(const i of indices){assert.deepEqual(after.paragraphs[i].tabStops.direct,[{positionTwips:480,alignment:'left',leader:'none'}]);assert.deepEqual(after.paragraphs[i].font,before.paragraphs[i].font);assert.deepEqual(after.paragraphs[i].lineSpacing,before.paragraphs[i].lineSpacing)}
for(let i=0;i<before.paragraphs.length;i++)if(!indices.includes(i))assert.deepEqual(after.paragraphs[i],before.paragraphs[i]);
const report={id:doc.id,elapsedSeconds:Math.round((Date.now()-started)/1000),targets:indices,textPreserved:true,otherParagraphsUnchanged:true,visualVerified:false,error:events.includes('event: error')};await writeFile(dir+'/report.json',JSON.stringify(report,null,2));console.log(JSON.stringify(report));
