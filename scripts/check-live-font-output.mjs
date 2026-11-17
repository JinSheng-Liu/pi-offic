import assert from 'node:assert/strict'
import { readFile, writeFile } from 'node:fs/promises'
import {readEditorDocument} from '../server/document/editor-reader.mjs'
const base = 'http://127.0.0.1:5175'
assert.equal((await (await fetch(`${base}/api/health`)).json()).testMode, true)
const { documents } = await (await fetch(`${base}/api/documents`)).json()
const doc = documents.filter(d => d.name === '公文字体验收副本.docx').sort((a,b) => b.updatedAt.localeCompare(a.updatedAt))[0]
assert.ok(doc?.revision >= 2, '真实模型没有生成版本')
const bytes = new Uint8Array(await (await fetch(`${base}/api/documents/${doc.id}/file`)).arrayBuffer())
const body = readEditorDocument(bytes)
const original = readEditorDocument(await readFile('data/validation/font-regression/公文字体验收副本.docx'))
assert.deepEqual(body.paragraphs.map(p => p.text), original.paragraphs.map(p => p.text), '真实模型或编辑器保存改变了原文')
let runs = 0
for (const p of body.paragraphs) for (const run of p.runs) {
  assert.equal(run.value.rPr?.rFonts?.eastAsia, '仿宋')
  assert.equal(run.value.rPr?.rFonts?.eastAsiaTheme, undefined)
  runs++
}
await writeFile(`data/validation/font-regression/live-v${doc.revision}.docx`, bytes)
console.log(JSON.stringify({ id: doc.id, revision: doc.revision, paragraphs: body.paragraphs.length, verifiedRuns: runs, eastAsia: '仿宋' }))
