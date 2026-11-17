import assert from 'node:assert/strict'
import { readFile, writeFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { format } from '../server/engine.mjs'
import { unpack, xml, W, digest } from '../server/document/package.mjs'

// Replay captured tool arguments without a model call or changing the user's revision.
const input = process.argv[2]
if (!input) throw new Error('Usage: node scripts/replay-format-task.mjs <captured-task.json>')
const task = JSON.parse(await readFile(input, 'utf8'))
const root = dirname(dirname(dirname(task.session)))
const metadata = JSON.parse(await readFile(join(root, 'metadata.json'), 'utf8'))
const source = await readFile(join(root, metadata.versions[0].file))
const started = Date.now()
const result = await format(source, task.args.operations)
const before = unpack(source), after = unpack(result.bytes)
const a = xml(before['word/document.xml']), b = xml(after['word/document.xml'])
const elements = (doc, name) => Array.from(doc.getElementsByTagNameNS(W, name))
const paragraphText = doc => elements(doc, 'p').map(p => Array.from(p.getElementsByTagNameNS(W, 't')).map(n => n.textContent).join(''))
assert.deepEqual(paragraphText(b), paragraphText(a), 'All XML paragraph text must survive')
const counts = {}
for (const name of ['p', 'tbl', 'tr', 'tc', 'drawing', 'sectPr', 'fldChar', 'instrText']) {
  counts[name] = elements(a, name).length
  assert.equal(elements(b, name).length, counts[name], `${name} count changed`)
}
assert.deepEqual(elements(b, 'instrText').map(n => n.textContent), elements(a, 'instrText').map(n => n.textContent))
const media = Object.keys(before).filter(n => n.startsWith('word/media/') && !n.endsWith('/'))
for (const name of media) assert.equal(digest(after[name]), digest(before[name]), `Media changed: ${name}`)
const report = { operations: task.args.operations.length, ms: Date.now() - started, affected: result.affected,
  counts, unchangedMedia: media.length, allParagraphTextUnchanged: true, fieldInstructionsUnchanged: true,
  validation: result.validation }
await writeFile(join(dirname(input), 'replayed.docx'), result.bytes)
await writeFile(join(dirname(input), 'replay-result.json'), JSON.stringify(report, null, 2))
console.log(JSON.stringify(report, null, 2))
