import test from 'node:test'
import assert from 'node:assert/strict'
import {documentFixture,paragraph,table} from './helpers/docx-fixture.mjs'
import {inspect,format} from '../server/engine.mjs'

test('database table centers short cells but retains long descriptions and unrelated tables',async()=>{
  const bytes=documentFixture(paragraph('正文不改变')+table([['字段名','类型','约束','注释'],['id','int','主键','订单 ID'],['address','varchar(255)','','这里是一段较长的收货地址字段说明，需要保留左对齐方便逐行阅读和检查']])+table([['编号','步骤','预期结果','备注'],['1','登录','成功','正常']])),before=await inspect(bytes)
  assert.equal(before.tables[0].layoutRecommendation.kind,'database-schema')
  assert.equal(before.tables[1].layoutRecommendation,null)
  const out=await format(bytes,before.tables[0].layoutRecommendation.operations),after=await inspect(out.bytes)
  assert.deepEqual(after.paragraphs.map(p=>p.text),before.paragraphs.map(p=>p.text))
  assert.equal(after.paragraphs.find(p=>p.text==='订单 ID').paragraph.alignment,'Centered')
  assert.equal(after.paragraphs.find(p=>p.text.startsWith('这里是一段')).paragraph.alignment,'Left')
  assert.equal(after.paragraphs.find(p=>p.text==='id').paragraph.firstLineIndent,0)
  assert.ok(after.tables[0].cellDetails.every(c=>c.verticalAlignment==='center'))
  for(const p of before.paragraphs.filter(p=>['正文不改变','登录','成功','正常'].includes(p.text)))assert.deepEqual(after.paragraphs[p.index],p)
})

test('analysis matrix uses narrative layout without changing unrelated tables or text', async () => {
  const bytes = documentFixture(table([['问题','形成原因','影响'],['目录质量不高','编目与业务脱节、更新机制弱','数据难发现、共享申请反复沟通']]) + table([['项目','数量','备注'],['甲','1','保持']]))
  const before = await inspect(bytes), recommendation = before.tables[0].layoutRecommendation
  assert.equal(recommendation.kind, 'analysis-matrix')
  assert.equal(before.tables[1].layoutRecommendation, null)
  const formatted = await format(bytes, recommendation.operations)
  const out = await format(formatted.bytes, [recommendation.widthOperation])
  const after = await inspect(out.bytes)
  assert.deepEqual(after.paragraphs.map(p => p.text), before.paragraphs.map(p => p.text))
  assert.deepEqual(after.tables[1], before.tables[1])
  assert.equal(after.paragraphs.find(p => p.text === '目录质量不高').paragraph.alignment, 'Left')
  assert.equal(after.paragraphs.find(p => p.text === '问题').paragraph.alignment, 'Centered')
  const { unpack, xml, W } = await import('../server/document/package.mjs')
  const doc = xml(unpack(out.bytes)['word/document.xml'])
  const cols = [...doc.getElementsByTagNameNS(W, 'tbl')[0].getElementsByTagNameNS(W, 'gridCol')].map(n => Number(n.getAttributeNS(W,'w')))
  assert.ok(cols[0] < cols[1])
  assert.ok(Math.abs(cols[1] - cols[2]) <= 2)
  await assert.rejects(format(bytes, [{scope:'fitTableWidth',tableIndex:0,columnRatios:[1,1]}]), /网格/)
  await assert.rejects(format(bytes, [{scope:'fitTableWidth',tableIndex:0,columnRatios:[0,1,1]}]), /比例/)
})
