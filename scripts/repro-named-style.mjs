import {readFile,readdir,writeFile,mkdir} from 'node:fs/promises'
import {format} from '../server/engine.mjs'
const root=process.argv[2]||'data/documents/e7774821-9dbc-44e8-803e-0e2e9250f7e2'
const files=await readdir(root+'/agent/sessions')
let steps
for(const file of files)for(const line of (await readFile(root+'/agent/sessions/'+file,'utf8')).trim().split('\n'))for(const c of JSON.parse(line).message?.content||[])if(c.name==='execute_format_plan')steps=c.arguments.steps
let bytes=await readFile(root+'/original.docx')
for(const [index,step] of steps.entries()){const result=await format(bytes,step.operations);bytes=result.bytes;console.log(JSON.stringify({step:index+1,affected:result.affected,status:result.validation.status}))}
await mkdir('output/named-style-task-replay',{recursive:true})
await writeFile('output/named-style-task-replay/公文排版回归.docx',bytes)
