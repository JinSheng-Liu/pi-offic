import test from 'node:test'
import assert from 'node:assert/strict'
import {documentFixture,richParagraph} from './helpers/docx-fixture.mjs'
import {format,inspect} from '../server/engine.mjs'
import {unpack,xml,W} from '../server/document/package.mjs'
import {executeFormatPlan} from '../server/document/format-plan.mjs'

test('TOC layout formats all native entries at section width and preserves page text',async()=>{
  const source=documentFixture(richParagraph('目录')+richParagraph('第一章\t1','<w:pStyle w:val="TOC1"/>')+richParagraph('1.1 背景\t2','<w:pStyle w:val="TOC2"/>')+richParagraph('正文\t3'))
  const result=await format(source,[{scope:'tocEntries'}]),data=await inspect(result.bytes)
  for(const p of data.paragraphs.slice(1,3))assert.deepEqual(p.tabStops.direct,[{positionTwips:9026,alignment:'right',leader:'dot'}])
  assert.deepEqual(data.paragraphs.map(p=>p.text),(await inspect(source)).paragraphs.map(p=>p.text))
  assert.equal(data.paragraphs[3].tabStops.direct.length,0)
  const again=await format(result.bytes,[{scope:'tocEntries'}])
  assert.deepEqual((await inspect(again.bytes)).paragraphs.map(p=>p.tabStops),data.paragraphs.map(p=>p.tabStops))
})

test('tabular text becomes independent wrapping cells preserving runs and surrounding content',async()=>{
  const rows=['问题\t形成原因\t影响','目录质量不高\t编目与业务脱节、更新机制弱\t数据难发现、共享申请反复沟通','标准碎片化\t部门系统长期独立建设\t跨系统关联、统计和核验成本高']
  const source=documentFixture(richParagraph('前文')+rows.map(s=>richParagraph(s,'','<w:rFonts w:eastAsia="宋体"/><w:b/>')).join('')+richParagraph('后文'))
  const out=await format(source,[{scope:'tabularText',indices:[1,2,3]}]),data=await inspect(out.bytes)
  assert.equal(data.tables.length,1)
  assert.deepEqual(data.cells.map(c=>c.text),rows.flatMap(s=>s.split('\t')))
  assert.equal(data.paragraphs[0].text,'前文');assert.equal(data.paragraphs.at(-1).text,'后文')
  const doc=xml(unpack(out.bytes)['word/document.xml'])
  assert.equal(doc.getElementsByTagNameNS(W,'tab').length,0)
  assert.equal(doc.getElementsByTagNameNS(W,'b').length,9)
  assert.equal(doc.getElementsByTagNameNS(W,'tblW')[0].getAttributeNS(W,'w'),'9026')
  assert.equal(doc.getElementsByTagNameNS(W,'noWrap').length,0)
})

test('tabular conversion rejects ambiguous, noncontiguous, and complex content',async()=>{
  for(const blocks of [richParagraph('a\tb')+richParagraph('c\td\te'),richParagraph('a\tb','<w:pStyle w:val="TOC1"/>')+richParagraph('c\td'),richParagraph('a\tb').replace('</w:p>','<w:bookmarkStart w:id="1" w:name="keep"/></w:p>')+richParagraph('c\td')]){
    await assert.rejects(format(documentFixture(blocks),[{scope:'tabularText',indices:[0,1]}]))
  }
  await assert.rejects(format(documentFixture(richParagraph('a\tb')+richParagraph('中间')+richParagraph('c\td')),[{scope:'tabularText',indices:[0,2]}]))
})

test('native TOC fields survive and each section supplies its own right tab position',async()=>{
  const field='<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>'
  const end='<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
  const section='<w:sectPr><w:pgSz w:w="10000" w:h="16000"/><w:pgMar w:left="1000" w:right="1000"/></w:sectPr>'
  const source=documentFixture(field+richParagraph('章节甲\t9')+end+richParagraph('分节段落',section)+richParagraph('章节乙\t10','<w:pStyle w:val="TOC1"/>'))
  const out=await format(source,[{scope:'tocEntries'}]),data=await inspect(out.bytes)
  assert.equal(data.paragraphs.find(p=>p.text==='章节甲\t9').tabStops.direct[0].positionTwips,8000)
  assert.equal(data.paragraphs.find(p=>p.text==='章节乙\t10').tabStops.direct[0].positionTwips,9026)
  const doc=xml(unpack(out.bytes)['word/document.xml'])
  assert.equal(doc.getElementsByTagNameNS(W,'fldChar').length,3)
  assert.equal(doc.getElementsByTagNameNS(W,'instrText')[0].textContent,' TOC \\o "1-3" ')
})

test('plans cannot reuse stale paragraph indices after tabular conversion',async()=>{
  await assert.rejects(executeFormatPlan({},'unused',1,[{label:'转表格',operations:[{scope:'tabularText',indices:[0,1]}]}]),/结构转换必须独立执行/)
})

test('TOC styles and static chapter lists are not native directory fields', async () => {
  const {hasNativeToc} = await import('../server/document/toc-layout.mjs')
  const source = documentFixture(richParagraph('目录') + richParagraph('1. 研究摘要与方法','<w:pStyle w:val="TOC1"/>'))
  assert.equal(hasNativeToc(source), false)
  const result = await format(source, [{scope:'tocEntries'}])
  assert.equal(result.affected, 0)
  assert.equal(result.validation.status, 'unverified')
  assert.equal(hasNativeToc(documentFixture('<w:p><w:fldSimple w:instr=" TOC "><w:r><w:t>目录结果</w:t></w:r></w:fldSimple></w:p>')), true)
})

for (const headers of [['问题','形成原因','影响'],['指标','公开数据','分析含义']]) {
  test(`tab conversion applies narrative layout atomically: ${headers[0]}`, async () => {
    const source = documentFixture([headers.join('\t'),'甲\t乙\t丙'].map(value=>richParagraph(value)).join(''))
    const result = await format(source,[{scope:'tabularText',indices:[0,1]}])
    const data = await inspect(result.bytes)
    assert.equal(data.tables.length,1)
    assert.ok(data.tables[0].gridWidthsTwips[0]<data.tables[0].gridWidthsTwips[2])
    assert.equal(data.paragraphs.find(p=>p.text==='甲').paragraph.alignment,'Left')
    assert.equal(data.paragraphs.find(p=>p.text===headers[0]).paragraph.alignment,'Centered')
    assert.deepEqual(data.cells.map(c=>c.text),[...headers,'甲','乙','丙'])
  })
}
