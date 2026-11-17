import { mkdir, writeFile } from 'node:fs/promises'
import { paragraph, table, withInlinePicture } from '../test/helpers/docx-fixture.mjs'
import { runDocumentJob } from '../server/document/document-jobs.mjs'

const blocks = Array.from({ length: 500 }, (_, i) => paragraph(`第 ${i + 1} 段，中文正文 English 123。用于验证长文档实时修改时的阅读与滚动。`)).join('')
const equation = '<w:p xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath></w:p>'
const source = withInlinePicture(blocks, table([['项目','说明'],['字段','保留表格内容']]) + equation)
const before = performance.now()
let ticks = 0, maxGap = 0, lastTick = before
const timer = setInterval(() => {
  const now = performance.now()
  maxGap = Math.max(maxGap, now - lastTick)
  lastTick = now
  ticks++
}, 10)
try {
  const {result} = await runDocumentJob('batch', {source,revision:1,batch:{kind:'format',operations:[
    {scope:'all',font:{name:'宋体',latinName:'Times New Roman',size:14},paragraph:{lineSpacingMultiple:1.5}},
  ]}})
  const out = 'output/live-format-probe'
  await mkdir(out,{recursive:true})
  await writeFile(`${out}/source.docx`,source)
  await writeFile(`${out}/expected.docx`,result.bytes)
  await writeFile(`${out}/replay.json`,JSON.stringify(result.replay))
  const report={paragraphs:result.affected,elapsedMs:performance.now()-before,mainThreadTicks:ticks,maxTimerGapMs:maxGap}
  await writeFile(`${out}/server-performance.json`,JSON.stringify(report,null,2))
  console.log(report)
} finally { clearInterval(timer) }
