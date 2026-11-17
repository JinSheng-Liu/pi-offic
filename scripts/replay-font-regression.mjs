import { readFile, readdir, mkdir, writeFile } from 'node:fs/promises'
import assert from 'node:assert/strict'
import {readEditorDocument} from '../server/document/editor-reader.mjs'
import { format } from '../server/engine.mjs'

// Replay the actual failed operations without calling a model or changing user history.
const root = 'data/documents/4a4aee96-87ef-45f8-ba90-f66204739da1'
const metadata = JSON.parse(await readFile(`${root}/metadata.json`, 'utf8'))
const entries = []
for (const file of await readdir(`${root}/agent/sessions`)) {
  for (const line of (await readFile(`${root}/agent/sessions/${file}`, 'utf8')).trim().split('\n')) {
    const message = JSON.parse(line).message
    if (message?.role === 'assistant' && Array.isArray(message.content)) entries.push(...message.content)
  }
}
const operations = entries.find(entry => entry.type === 'toolCall' && entry.name === 'apply_formatting')?.arguments.operations
assert.ok(operations, '缺少真实失败操作记录')
const original = await readFile(`${root}/${metadata.versions[0].file}`)
const saved = await readFile(`${root}/${metadata.versions[1].file}`)
const expected = new Map()
const originalBody = readEditorDocument(original)
for (const op of operations) if (op.font?.name) {
  const indices = op.scope === 'all' ? originalBody.paragraphs.map((_, i) => i) : op.indices || []
  for (const index of indices) expected.set(index, op.font.name)
}
async function failures(bytes) {
  const body = readEditorDocument(bytes)
  assert.equal(body.text, originalBody.text, '原文发生变化')
  const failed = []
  for (const [index, name] of expected) {
    const p = body.paragraphs[index]
    const fonts = [...p.runs.map(run => run.value.rPr?.rFonts), p.p.pPr?.rPr?.rFonts]
    if (fonts.some(font => font?.eastAsia !== name || font?.eastAsiaTheme)) failed.push(index)
  }
  return failed
}
const before = await failures(saved)
assert.ok(before.length > 0, '旧输出未复现中文字体错误')
const result = await format(original, operations)
const after = await failures(result.bytes)
assert.deepEqual(after, [], '新版仍未正确写入中文字体')
await mkdir('data/validation/font-regression', { recursive: true })
await writeFile('data/validation/font-regression/公文字体验收副本.docx', result.bytes)
const report = { oldOutputFailingParagraphs: before, replayFailingParagraphs: after, checkedParagraphs: expected.size, textPreserved: true, limitations: ['原操作未包含同段标题局部格式，本检查不将其视为完成', '尚未验证 Word 实际排版和浏览器再次保存'] }
await writeFile('data/validation/font-regression/report.json', JSON.stringify(report, null, 2))
console.log(JSON.stringify(report))
