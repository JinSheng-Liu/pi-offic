import {readFile,mkdir,writeFile} from 'node:fs/promises'
import assert from 'node:assert/strict'
import {format,inspect} from '../server/engine.mjs'
import {unpack,xml,W,canonical,digest} from '../server/document/package.mjs'
const output='output/tab-layout-verification';await mkdir(output,{recursive:true});const reports=[]
for(const [id,scope,name] of [['cdcb777d-adad-49fc-957b-9f7817e2af61','tabularText','研究报告_分列修复'],['67114bc8-8446-493d-b549-660371334a7c','tocEntries','论文_目录修复']]){
 const dir=`data/documents/${id}`,meta=JSON.parse(await readFile(`${dir}/metadata.json`,'utf8')),v=meta.versions.find(v=>v.revision===meta.revision),source=await readFile(`${dir}/${v.file}`),before=await inspect(source)
 const op=scope==='tabularText'?{scope,indices:[53,54,55]}:{scope};const result=await format(source,[op]),after=await inspect(result.bytes)
 const a=unpack(source),b=unpack(result.bytes)
 for(const part of Object.keys(a).filter(p=>p!=='word/document.xml'&&!p.endsWith('/')))assert.equal(/\.(xml|rels)$/.test(part)?canonical(xml(a[part]).documentElement):digest(a[part]),/\.(xml|rels)$/.test(part)?canonical(xml(b[part]).documentElement):digest(b[part]),part)
 const fields=bytes=>Array.from(xml(bytes).getElementsByTagNameNS(W,'instrText')).map(n=>n.textContent)
 assert.deepEqual(fields(a['word/document.xml']),fields(b['word/document.xml']))
 if(scope==='tocEntries'){
  assert.deepEqual(after.paragraphs.map(p=>p.text),before.paragraphs.map(p=>p.text))
  assert.ok(result.affected>0)
  for(const p of after.paragraphs.filter(p=>p.role==='toc'&&/\t[0-9IVXLCDMivxlcdm]+\s*$/.test(p.text)))assert.ok(p.tabStops.direct.some(t=>t.alignment==='right'&&t.leader==='dot'))
 }else{
  assert.deepEqual(after.paragraphs.map(p=>p.text),before.paragraphs.flatMap((p,i)=>op.indices.includes(i)?p.text.split('\t'):[p.text]))
  assert.equal(after.tables.length,before.tables.length+1)
 }
 await writeFile(`${output}/${name}.docx`,result.bytes)
 reports.push({name,sourceRevision:meta.revision,affected:result.affected,originalUnchanged:true,otherPartsUnchanged:true,fieldInstructionsUnchanged:true,validation:result.validation})
}
await writeFile(`${output}/report.json`,JSON.stringify(reports,null,2));console.log(JSON.stringify(reports,null,2))
