import { readFile, writeFile, mkdir } from 'node:fs/promises'
import { createHash } from 'node:crypto'
import assert from 'node:assert/strict'
import { format, inspect } from '../server/engine.mjs'
const source = 'D:/xiangmu/agent-office/agent-office/测试文档/党政机关公文_无格式版.docx'
const bytes = await readFile(source)
const beforeHash = createHash('sha256').update(bytes).digest('hex')
const session = await readFile('data/documents/749b9681-27a8-4aea-a1ea-d9ccab7f7536/agent/sessions/2026-11-01T06-59-48-472Z_01a195e8-4078-77aa-9906-606d1f3b6ea8.jsonl','utf8')
const calls = session.trim().split('\n').flatMap(line => JSON.parse(line).message?.content || []).filter(c => c.type === 'toolCall' && c.name === 'apply_formatting')
const original = await inspect(bytes)
await mkdir('data/validation/real-document', { recursive: true })
for (let i=0; i<calls.length; i++) {
 const result = await format(bytes, calls[i].arguments.operations)
 const after = await inspect(result.bytes)
 assert.deepEqual(after.paragraphs.map(p=>p.text), original.paragraphs.map(p=>p.text))
 await writeFile(`data/validation/real-document/replay-${i+1}.docx`,result.bytes)
 console.log(JSON.stringify({replay:i+1,paragraphs:after.paragraphs.length,validation:result.validation,sections:after.sections}))
}
assert.equal(createHash('sha256').update(await readFile(source)).digest('hex'),beforeHash)
console.log('Original hash unchanged: '+beforeHash)
