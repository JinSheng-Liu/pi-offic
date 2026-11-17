import { mkdtemp,rm,mkdir,writeFile } from 'node:fs/promises'
import {tmpdir} from 'node:os'
import {join} from 'node:path'
import {createServer} from 'node:http'
import {paragraphDocument} from '../test/helpers/docx-fixture.mjs'
import {zipSync,unzipSync} from 'fflate'
import {format,inspect,compareProtected} from '../server/engine.mjs'
import {DocumentStore} from '../server/store.mjs'
const root=await mkdtemp(join(tmpdir(),'pi-deep-audit-'));process.env.PI_DOCX_DATA_DIR=root
const findings=[]
const bytes=paragraphDocument(['目标正文','不应改变的段落'])
const probe=async(name,fn)=>{try{findings.push({name,...await fn()})}catch(e){findings.push({name,error:e.message})}}
await probe('impossible_margins',async()=>{const r=await format(bytes,[{scope:'sections',sectionIndices:[0],margins:{left:10000,right:10000}}]);const s=(await inspect(r.bytes)).sections[0];return {accepted:true,width:s.page.widthTwips,left:s.margins.left,right:s.margins.right,validation:r.validation.status}})
await probe('non_target_format_guard',async()=>{const changed=await format(bytes,[{scope:'indices',indices:[1],font:{bold:true}}]);return {accepted:compareProtected(bytes,changed.bytes).status}})
await probe('manual_save_changes_existing_part',async()=>{const parts=unzipSync(new Uint8Array(bytes));const key='word/styles.xml';let xml=new TextDecoder().decode(parts[key]);xml=xml.replace(/(val=")([^"]+)(")/, '$1' + '40' + '$3');parts[key]=new TextEncoder().encode(xml);const store=new DocumentStore(root);const d=await store.create('audit.docx',bytes);const r=await store.save(d.id,1,zipSync(parts));return {accepted:true,revision:r.revision}})
await probe('restore_validation_upgrade',async()=>{const store=new DocumentStore(root);const d=await store.create('restore.docx',bytes);const r=await store.restore(d.id,1,1);return {before:d.versions[0].validation.status,after:r.versions.at(-1).validation.status}})
const {createApiMiddleware}=await import('../server/api.mjs');const staticRoot=join(root,'static');await mkdir(staticRoot);await writeFile(join(staticRoot,'index.html'),'ok');const server=createServer(createApiMiddleware({staticRoot}));await new Promise(r=>server.listen(0,'127.0.0.1',r));const base='http://127.0.0.1:'+server.address().port
await probe('missing_static_file',async()=>{try{const r=await fetch(base+'/missing.js',{signal:AbortSignal.timeout(1200)});return {status:r.status}}catch(e){return {timeout:e.name==='TimeoutError'}}})
await probe('foreign_origin_write',async()=>{const r=await fetch(base+'/api/documents',{method:'POST',headers:{Origin:'https://foreign.example','content-type':'text/plain'}});return {status:r.status}})
server.closeAllConnections();await new Promise(r=>server.close(r));await rm(root,{recursive:true,force:true});await mkdir('data/validation',{recursive:true});await writeFile('data/validation/deep-audit.json',JSON.stringify(findings,null,2));console.log(JSON.stringify(findings,null,2))
