// Specific regression replay for the thesis, not a general tab normalization rule.
import {readFile, writeFile, mkdir} from 'node:fs/promises'
import assert from 'node:assert/strict'
import {readEditorDocument} from '../server/document/editor-reader.mjs'
import {format, unpack, digest} from '../server/engine.mjs'

const input = 'data/validation/latin-font/论文_英文与缩进修正版.docx'
const folder = 'data/validation/heading-indent'
const source = await readFile(input)
const before = (readEditorDocument(source)).paragraphs
const targets = before.flatMap((p, index) =>
  /^\t[^\s]/u.test(p.text) && /^Heading[23]$/.test(p.styleId) && p.firstLineIndent === 0 && [21,24].includes(p.leftIndent)
    ? [{index, left:p.leftIndent}] : [])
assert.equal(targets.length, 12, 'Source document changed; inspect before replaying')
const result = await format(source, targets.map(({index, left}) => ({scope:'indices', indices:[index], paragraph:{firstLineIndent:-left}})))
const after = (readEditorDocument(result.bytes)).paragraphs
assert.deepEqual(after.map(p=>p.text), before.map(p=>p.text))
for (const {index, left} of targets) {
  assert.equal(after[index].leftIndent, left)
  assert.equal(after[index].firstLineIndent, -left)
}
const originalParts = unpack(source), outputParts = unpack(result.bytes)
const media = Object.keys(originalParts).filter(n=>n.startsWith('word/media/')&&!n.endsWith('/'))
for (const name of media) assert.equal(digest(outputParts[name]), digest(originalParts[name]))
for (const p of after) for (const run of p.runs) {
  assert.equal(run.value.rPr?.rFonts?.ascii, 'Times New Roman')
  assert.equal(run.value.rPr?.rFonts?.hAnsi, 'Times New Roman')
}
await mkdir(folder, {recursive:true})
await writeFile(`${folder}/论文_标题缩进与西文字体修正版.docx`, result.bytes)
const report = {source:input, sourceHash:digest(source), correctedHeadings:targets, paragraphs:after.length, mediaPreserved:media.length, textPreserved:true, latinFontsVerified:true, validation:result.validation}
await writeFile(`${folder}/report.json`, JSON.stringify(report,null,2), 'utf8')
console.log(JSON.stringify(report))
