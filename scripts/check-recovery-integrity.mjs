import {readFile,writeFile} from 'node:fs/promises'
import {readEditorDocument} from '../server/document/editor-reader.mjs'
import {unpack,digest} from '../server/document/package.mjs'
const l=JSON.parse(await readFile('output/recovery-live-location.json','utf8'))
const meta=JSON.parse(await readFile(l.root+'/'+l.id+'/metadata.json','utf8'))
const base=await readFile(l.root+'/'+l.id+'/'+meta.versions.find(v=>v.revision===5).file)
const latest=await readFile(l.root+'/'+l.id+'/'+meta.versions.at(-1).file)
const a=unpack(base),b=unpack(latest),media=Object.keys(a).filter(k=>k.startsWith('word/media/'))
const report={revision:meta.revision,textPreserved:readEditorDocument(base).text===readEditorDocument(latest).text,mediaCount:media.length,mediaPreserved:media.every(k=>b[k]&&digest(a[k])===digest(b[k])),newVersions:meta.versions.filter(v=>v.revision>5).length,taskGroups:new Set(meta.versions.filter(v=>v.revision>5).map(v=>v.task?.id)).size,scopes:meta.versions.filter(v=>v.revision>5).flatMap(v=>v.validation.operationReceipt?.operations?.map(o=>o.scope)||[])}
await writeFile(l.root+'/integrity-report.json',JSON.stringify(report,null,2));console.log(report)
