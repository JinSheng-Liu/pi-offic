import {createServer} from 'vite'
import fs from 'node:fs'
const out='output/editor-browser-probe'
fs.mkdirSync(out,{recursive:true})
const server=await createServer({configFile:false,cacheDir:'tmp/editor-probe-vite-cache',optimizeDeps:{esbuildOptions:{target:'esnext'}},server:{host:'127.0.0.1',port:5186,strictPort:true},plugins:[{name:'probe-output',configureServer(s){s.middlewares.use('/probe-output',async(req,res)=>{
  if(req.method!=='POST'){res.statusCode=405;res.end();return}
  const chunks=[];for await(const c of req)chunks.push(c)
  const name=req.url?.slice(1)
  if(!['roundtrip.docx','formatted.docx','undone.docx','report.json'].includes(name)){res.statusCode=400;res.end();return}
  fs.writeFileSync(`${out}/${name}`,Buffer.concat(chunks));res.end('ok')
})}}]})
await server.listen();console.log('Independent editor probe: http://127.0.0.1:5186/scripts/probe-editor-browser.html')
