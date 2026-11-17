import {readFile,readdir,writeFile} from 'node:fs/promises'
import {budgetDocumentContext} from '../server/document/read-context.mjs'
const folder='data/documents/ff3617e9-770e-4d7b-82cb-b5ba73b0286b/agent/sessions'
const files=await readdir(folder),messages=[]
for(const f of files)for(const line of (await readFile(folder+'/'+f,'utf8')).trim().split('\n')){const x=JSON.parse(line);if(x.message)messages.push(x.message)}
const projected=await budgetDocumentContext(messages)
const report={beforeBytes:Buffer.byteLength(JSON.stringify(messages)),afterBytes:Buffer.byteLength(JSON.stringify(projected)),messages:messages.length,omittedReads:projected.filter(m=>m.content?.some?.(c=>c.text?.includes('omittedForContextBudget'))).length}
await writeFile('output/context-budget-real-report.json',JSON.stringify(report,null,2));console.log(report)
