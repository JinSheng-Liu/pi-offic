import test from 'node:test'
import assert from 'node:assert/strict'
import { readdir, readFile } from 'node:fs/promises'
import { projectInspection, recommendSkills } from '../server/document/skill-routing.mjs'
import { deduplicateDocumentReads } from '../server/document/read-context.mjs'
import { inspectLayoutRisks } from '../server/document/layout-risks.mjs'
import { xml } from '../server/document/package.mjs'

test('floating textbox risks remain distinct from editable paragraph indices', () => {
  const doc = xml(new TextEncoder().encode(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:body><w:p><w:r><w:drawing><wp:anchor><wp:wrapNone/><w:txbxContent><w:p><w:r><w:t>图4.1 架构</w:t></w:r></w:p></w:txbxContent></wp:anchor></w:drawing><w:instrText>TOC \\o "1-2"</w:instrText></w:r></w:p></w:body></w:document>`))
  const risks = inspectLayoutRisks(doc)
  assert.equal(risks.layoutRisks.floatingTextBoxCount,1)
  assert.equal(risks.layoutRisks.noWrapCount,1)
  assert.equal(risks.layoutRisks.tocInstructions,1)
  assert.equal(risks.layoutRisks.visualVerified,false)
  const data = {...document(['正文']),...risks}
  const projected = projectInspection(data,{view:'floatingObjects'})
  assert.equal(projected.items[0].text,'图4.1 架构')
  assert.equal(projected.items[0].readOnly,true)
  assert.equal(projected.items[0].index,undefined)
  assert.equal(projected.layoutRisks.warnings.length,3)
})

test('read deduplication preserves latest evidence, revisions, writes, errors and stored messages', async () => {
  const read = (id, revision = 1) => ({role:'toolResult',toolName:'inspect_document',toolCallId:id,
    content:[{type:'text',text:JSON.stringify({revision,hash:'bytes',items:[{index:0,text:'正文'}]})}]})
  const messages = [read('first'), read('other-revision',2), {...read('failed'),isError:true},
    {...read('write'),toolName:'apply_formatting'}, read('latest')]
  const original = structuredClone(messages)
  const result = await deduplicateDocumentReads(messages)
  assert.equal(JSON.parse(result[0].content[0].text).identicalReadAvailableAt, 'latest')
  assert.deepEqual(result.slice(1), original.slice(1))
  assert.deepEqual(messages, original)
})

function document(texts = []) {
  return { paragraphs: texts.map((text, index) => ({ index, text })), tables: [], sections: [],
    captions: [], captionTargets: { figures: [], tables: [] }, elements: {}, limitations: [] }
}
test('structure scans stay compact while targeted reads retain formatting', () => {
  const data = document(['正文'.repeat(300)])
  data.paragraphs[0].font = { name: '宋体', size: 12 }
  data.paragraphs[0].paragraph = { firstLineIndent: 24 }
  const scan = projectInspection(data)
  assert.equal(scan.items[0].font, undefined)
  assert.ok(scan.items[0].text.length <= 160)
  assert.equal(scan.items[0].textTruncated, true)
  assert.deepEqual(projectInspection(data, {indices:[0]}).items[0].font, data.paragraphs[0].font)
  assert.deepEqual(projectInspection(data, {detail:'formatting'}).items[0].font, data.paragraphs[0].font)
})

test('targeted formatting reads omit repeated document-wide context', () => {
  const data = document(['正文'])
  data.pageNumbers = [{ index: 0 }]
  data.elements = { images: 5 }
  data.layoutRisks = { warnings: ['风险'] }
  data.limitations = ['限制']
  data.tables = Array.from({ length: 20 }, () => ({}))
  const result = projectInspection(data, { indices: [0], detail: 'formatting' })
  assert.equal(result.pageNumbers, undefined)
  assert.equal(result.layoutRisks, undefined)
  assert.equal(result.recommendations, undefined)
  assert.deepEqual(result.items, data.paragraphs)
})
test('type suggestions expose evidence and ambiguity, not a forced template', () => {
  assert.equal(recommendSkills(document(['普通文章'])).uncertain, true)
  const academic = recommendSkills(document(['本科毕业论文']))
  assert.equal(academic.candidates[0].name, 'docx-academic')
  assert.equal(academic.uncertain, false)
  assert.equal(recommendSkills(document(['毕业论文', '附件：合同，甲方'])).uncertain, true)
})
test('pagination visits every paragraph once and rejects invalid targets', () => {
  const data = document(Array.from({length: 103}, (_, i) => `段落${i}`))
  let offset = 0
  const seen = []
  do {
    const page = projectInspection(data, {offset})
    assert.ok(page.items.length <= 40)
    seen.push(...page.items.map(p => p.index))
    offset = page.pagination.nextOffset
  } while (offset !== null)
  assert.deepEqual(seen, Array.from({length:103}, (_, i) => i))
  assert.equal(projectInspection(data, {view:'summary'}).pagination.complete, false)
  assert.deepEqual(projectInspection(data, {indices:[99,2]}).items.map(p => p.index), [99,2])
  assert.throws(() => projectInspection(data, {indices:[103]}))
  assert.throws(() => projectInspection(data, {view:'tables',indices:[0]}))
  assert.throws(() => projectInspection(data, {limit:1000}))
  assert.throws(() => projectInspection(data, {view:'__proto__'}))
})
test('caption target indexes remain a separate namespace', () => {
  const data = document(['图'])
  data.captionTargets.figures = [{index:24,floating:true}]
  assert.deepEqual(projectInspection(data, {view:'figures'}).items, [{index:24,floating:true}])
  assert.equal(projectInspection(data).items[0].index, 0)
})
test('skills meet length, structure, UTF-8 and registered-reference requirements', async () => {
  const dirs = await readdir(new URL('../docs/skills/', import.meta.url), {withFileTypes:true})
  const skills = dirs.filter(d => d.isDirectory())
  const names = new Set(skills.map(d => d.name))
  assert.equal(skills.length, 10)
  for (const dir of skills) {
    const content = await readFile(new URL(`../docs/skills/${dir.name}/SKILL.md`, import.meta.url), 'utf8')
    assert.ok(content.replace(/\r\n/g, "\n").includes(`name: ${dir.name}\n`))
    assert.match(content, /[\u4e00-\u9fff]/u)
    assert.ok(!content.includes('\uFFFD'))
    const lines = content.trimEnd().split(/\r?\n/)
    assert.ok(lines.length >= 70 && lines.length <= 90, `${dir.name}: ${lines.length} lines`)
    for (const heading of ['适用范围', '输入与前置检查', '识别与决策', '执行流程', '验收要求', '案例与失败边界']) {
      assert.ok(lines.includes(`## ${heading}`), `${dir.name}: missing ${heading}`)
    }
    for (const [reference] of content.matchAll(/docx-[a-z-]+/g)) {
      assert.ok(names.has(reference), `${dir.name}: unknown skill ${reference}`)
    }
  }
})
