import { canonical, xml as parsePackageXml } from '../server/document/package.mjs'
import test from 'node:test'
import assert from 'node:assert/strict'
import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import {paragraphDocument,documentFixture,paragraph,richParagraph,table,withInlinePicture} from './helpers/docx-fixture.mjs'
import {readEditorDocument,findTextRanges} from '../server/document/editor-reader.mjs'
import { unzipSync, zipSync } from 'fflate'
import { DocumentStore } from '../server/store.mjs'
import { validateOperations } from '../server/document/rules.mjs'
import { inspect, format, insertCaptions, compareProtected } from '../server/engine.mjs'

test('ZIP directory entries are optional but actual document parts must survive', async () => {
  const original = await fixture()
  const parts = unzipSync(original)
  parts['word/'] = new Uint8Array()
  parts['word/media/'] = new Uint8Array()
  assert.equal(compareProtected(zipSync(parts), original).status, 'passed')
  parts['word/media/keep.bin'] = new Uint8Array([1,2,3])
  assert.throws(() => compareProtected(zipSync(parts), original), /word\/media\/keep.bin/)
})

test('known OOXML boolean spellings are equivalent but value changes remain protected', async () => {
  const source = unzipSync(await fixture())
  const original = new TextDecoder().decode(source['word/document.xml'])
  const make = value => { const parts = {...source}; parts['word/document.xml'] = new TextEncoder().encode(original.replace('<w:pPr>', `<w:pPr><w:bidi w:val="${value}"/>`)); return zipSync(parts) }
  assert.equal(compareProtected(make('0'),make('false')).status,'passed')
  assert.throws(()=>compareProtected(make('0'),make('true')),/正文结构/)
})

async function fixture() {
  return documentFixture(richParagraph('一级标题：验证文档','<w:pStyle w:val="Heading1"/><w:outlineLvl w:val="0"/>')+paragraph('正文保持原文 English 123。'))
}
test('single leading tab positioning preserves text and does not change other paragraphs', async () => {
  const source=documentFixture(richParagraph('\t正文文字')+paragraph('保留原格式'))
  const before = await inspect(source)
  const out = await format(source, [{scope:'indices',indices:[0],paragraph:{leadingTabPosition:24}}])
  const after = await inspect(out.bytes)
  assert.deepEqual(after.paragraphs.map(p=>p.text),before.paragraphs.map(p=>p.text))
  assert.equal(after.paragraphs[0].tabStops.direct[0].positionTwips,480)
  assert.deepEqual(after.paragraphs[1],before.paragraphs[1])
  assert.throws(() => compareProtected(source, out.bytes), /结构/)
  await assert.rejects(format(out.bytes,[{scope:'indices',indices:[0],paragraph:{leadingTabPosition:30}}]), /制表位/)
  await assert.rejects(format(source,[{scope:'indices',indices:[1],paragraph:{leadingTabPosition:24}}]), /单个前导/)
  await assert.rejects(format(source,[{scope:'all',paragraph:{leadingTabPosition:24}}]), /明确索引/)
})
test('table row and cell layout edits are scoped, preserve content, and verify on reopen', async () => {
  const source=documentFixture(table([['字段','说明'],['编号','内容'],['1','示例']])),before=await inspect(source)
  const out=await format(source,[
    {scope:'tableRows',tableIndex:0,rowIndices:[0],table:{repeatHeader:true,cantSplit:true}},
    {scope:'tableCells',tableIndex:0,cellTargets:[{rowIndex:0,cellIndex:0}],table:{verticalAlignment:'Center'}},
  ])
  const after=await inspect(out.bytes)
  assert.deepEqual(after.paragraphs.map(p=>p.text),before.paragraphs.map(p=>p.text))
  assert.equal(after.tables[0].rowDetails[0].header,true)
  assert.equal(after.tables[0].rowDetails[0].cantSplit,true)
  const cells=parsePackageXml(unzipSync(out.bytes)['word/document.xml']).getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main','tc')
  assert.equal(cells[0].getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main','vAlign')[0].getAttribute('w:val'),'center')
  assert.equal(cells[1].getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main','vAlign').length,0)
  assert.equal(after.tables[0].rowDetails[1].header,false)
  assert.equal(after.tables[0].rowDetails[1].cantSplit,false)
  assert.equal(after.tables[0].cellDetails[0].verticalAlignment,'center')
  assert.equal(after.tables[0].cellDetails[1].verticalAlignment,null)
  assert.throws(()=>compareProtected(source,out.bytes),/结构/)
  assert.throws(()=>validateOperations([{scope:'tableRows',tableIndex:0,rowIndices:[0],table:{repeatHeader:true},indices:[1]}]))
  assert.throws(()=>validateOperations([{scope:'tableRows',tableIndex:0,rowIndices:[0],table:{cantSplit:'yes'}}]))
  assert.throws(()=>validateOperations([{scope:'tableCells',tableIndex:0,cellTargets:[{rowIndex:0,cellIndex:0,paragraphIndex:4}],table:{verticalAlignment:'Center'}}]))
  await assert.rejects(format(source,[{scope:'tableRows',tableIndex:0,rowIndices:[1],table:{repeatHeader:true}}]),/重复表头必须/)
  await assert.rejects(format(source,[{scope:'tableRows',tableIndex:0,rowIndices:[99],table:{cantSplit:true}}]),/索引超出/)
  const disabled=await format(out.bytes,[{scope:'tableRows',tableIndex:0,rowIndices:[0],table:{repeatHeader:false,cantSplit:false}}])
  const row=(await inspect(disabled.bytes)).tables[0].rowDetails[0]
  assert.equal(row.header,false)
  assert.equal(row.cantSplit,false)
})
test('a later invalid table target leaves the entire batch uncommitted', async () => {
  const directory=await mkdtemp(join(tmpdir(),'pi-docx-tables-'))
  try {
    const source=documentFixture(table([['名称','说明'],['1','正文']])),store=new DocumentStore(directory)
    const document=await store.create('表格.docx',source)
    let commits=0
    store.onCommit(()=>commits++)
    await assert.rejects(store.apply(document.id,1,[
      {scope:'tableRows',tableIndex:0,rowIndices:[0],table:{cantSplit:true}},
      {scope:'tableCells',tableIndex:0,cellTargets:[{rowIndex:99,cellIndex:0}],table:{verticalAlignment:'Center'}},
    ]),/索引超出/)
    assert.equal(commits,0)
    assert.equal((await store.metadata(document.id)).revision,1)
    assert.deepEqual(Buffer.from(await store.bytes(document.id)),Buffer.from(source))
  } finally {await rm(directory,{recursive:true,force:true})}
})
test('inspection distinguishes multiple and exact line spacing and exposes direct tab stops', async () => {
  const source = await fixture()
  const multiple = await format(source, [{scope:'indices',indices:[1],paragraph:{lineSpacingMultiple:1.5}}])
  const exact = await format(source, [{scope:'indices',indices:[1],paragraph:{lineSpacing:18}}])
  assert.equal((await inspect(multiple.bytes)).paragraphs[1].lineSpacing.rule, 'auto')
  assert.equal((await inspect(multiple.bytes)).paragraphs[1].lineSpacing.multiple, 1.5)
  assert.equal((await inspect(exact.bytes)).paragraphs[1].lineSpacing.rule, 'exact')
  assert.equal((await inspect(exact.bytes)).paragraphs[1].lineSpacing.points, 18)
  const info=(await inspect(documentFixture(richParagraph('测试','<w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="right" w:pos="1440" w:leader="dot"/></w:tabs>')))).paragraphs[0]
  assert.deepEqual(info.tabStops.direct, [
    {positionTwips:720,alignment:'left',leader:'none'},
    {positionTwips:1440,alignment:'right',leader:'dot'},
  ])
  assert.deepEqual(info.tabStops.effective, info.tabStops.direct)
})

test('body formatting excludes cover styles, table cells, captions, TOC, lists and pictures', async () => {
  const bytes=withInlinePicture(richParagraph('封面标题','<w:pStyle w:val="Title"/>','<w:sz w:val="60"/>')+richParagraph('需要设置的正文','','<w:sz w:val="22"/>')+richParagraph('图1 系统结构','','<w:sz w:val="18"/>'),richParagraph('第一章 …… 1','<w:pStyle w:val="TOC1"/>','<w:sz w:val="20"/>')+richParagraph('列表条目','<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>')+table([['编号','说明']]))
  const before = await inspect(bytes)
  assert.equal(before.paragraphs[3].inlinePictureCount, 1)
  const result = await format(bytes, [{scope:'body',font:{size:16},paragraph:{firstLineIndent:32}}])
  const after = await inspect(result.bytes)
  assert.equal(after.paragraphs[0].font.size, 30)
  assert.equal(after.paragraphs[1].font.size, 16)
  assert.equal(after.paragraphs[2].font.size, 9)
  assert.deepEqual(after.paragraphs[3].paragraph, before.paragraphs[3].paragraph)
  for (let i = 4; i < before.paragraphs.length; i++) {
    assert.deepEqual(after.paragraphs[i].paragraph, before.paragraphs[i].paragraph)
    assert.deepEqual(after.paragraphs[i].font, before.paragraphs[i].font)
  }
  assert.equal(after.elements.images, before.elements.images)
  const centered = await format(bytes, [{scope:'indices',indices:[3],paragraph:{alignment:'Centered'}}])
  const imageResult = await inspect(centered.bytes)
  assert.equal(imageResult.paragraphs[3].paragraph.alignment, 'Centered')
  assert.equal(imageResult.elements.images, before.elements.images)
  assert.deepEqual(imageResult.paragraphs[1].paragraph, before.paragraphs[1].paragraph)
  assert.equal(before.cells.length, 2)
  const firstCell = before.cells[0].paragraphIndices
  const secondCell = before.cells[1].paragraphIndices
  assert.equal(before.paragraphs[firstCell[0]].text, '编号')
  assert.equal(before.paragraphs[secondCell[0]].text, '说明')
  const tableResult = await format(bytes, [{scope:'indices',indices:firstCell,paragraph:{alignment:'Centered'}}])
  const tableAfter = await inspect(tableResult.bytes)
  assert.equal(tableAfter.paragraphs[firstCell[0]].paragraph.alignment, 'Centered')
  assert.deepEqual(tableAfter.paragraphs[secondCell[0]].paragraph, before.paragraphs[secondCell[0]].paragraph)
})

test('Latin font override preserves Chinese slots and handles later whole font changes', async () => {
  const original=await fixture()
  const result=await format(original,[{scope:'all',font:{name:'宋体',latinName:'Times New Roman'}}])
  const body=readEditorDocument(result.bytes)
  for(const p of body.paragraphs) for(const holder of [...p.runs.map(r=>r.value),p.p.pPr]) {
    assert.equal(holder.rPr.rFonts.ascii,'Times New Roman')
    assert.equal(holder.rPr.rFonts.hAnsi,'Times New Roman')
    assert.equal(holder.rPr.rFonts.eastAsia,'宋体')
  }
  assert.equal(body.paragraphs[1].text,'正文保持原文 English 123。')
  const inspection = await inspect(result.bytes)
  assert.deepEqual(inspection.paragraphs[1].font.asciiDirect, ['Times New Roman'])
  assert.deepEqual(inspection.paragraphs[1].font.hAnsiDirect, ['Times New Roman'])
  await format(original,[{scope:'all',font:{name:'宋体',latinName:'Times New Roman'}},{scope:'all',font:{name:'黑体'}}])
  const partial=await format(result.bytes,[{scope:'text',indices:[1],text:'English',font:{latinName:'Arial'}}])
  const p=(readEditorDocument(partial.bytes)).paragraphs[1]
  assert.equal(findTextRanges(p,'English')[0].font.name,'Arial')
  assert.equal(findTextRanges(p,'English')[0].runs[0].value.rPr.rFonts.eastAsia,'宋体')
})

test('inspection exposes leading tabs and indentation conflicts without changing text', async () => {
  const data=await inspect(documentFixture(richParagraph('\t正文示例','<w:ind w:firstLine="420"/>')))
  assert.equal(data.paragraphs[0].leadingWhitespace.tabs,1)
  assert.match(data.paragraphs[0].indentWarning,/制表符/)
  assert.equal(data.paragraphs[0].text,'\t正文示例')
})

test('heading tabs can keep the required left indent without doubling the first line', async () => {
  const original=documentFixture([['Heading2',24],['Heading3',21]].map(([style,indent])=>richParagraph('\t标题 English',`<w:pStyle w:val="${style}"/><w:ind w:left="${indent*20}" w:firstLine="0"/>`)).join(''))
  const out = await format(original, [
    {scope:'indices', indices:[0], paragraph:{firstLineIndent:-24}},
    {scope:'indices', indices:[1], paragraph:{firstLineIndent:-21}},
  ])
  const paragraphs = (readEditorDocument(out.bytes)).paragraphs
  for (const [i, indent] of [24, 21].entries()) {
    assert.equal(paragraphs[i].leftIndent, indent)
    assert.equal(paragraphs[i].firstLineIndent, -indent)
    assert.equal(paragraphs[i].text, '\t标题 English')
  }
})

test('captions insert adjacent to top-level figures and tables with chapter SEQ fields', async () => {
  const parts=unzipSync(documentFixture(richParagraph('第三章 系统设计','<w:pStyle w:val="Heading1"/>')+paragraph('图片目标')))
  const sourceXml=new TextDecoder().decode(parts['word/document.xml'])
  parts['word/document.xml']=new TextEncoder().encode(sourceXml.replace('</w:body>',
    '<w:p><w:r><w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:inline/></w:drawing></w:r></w:p>'+
    '<w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>数据</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body>'))
  const original=zipSync(parts)
  const result=insertCaptions(original,[
    {kind:'figure',targetIndex:2,text:'系统结构图'},
    {kind:'table',targetIndex:0,text:'系统参数'},
  ])
  assert.equal(result.affected,2)
  assert.equal(result.captions[0].sequence,1)
  assert.equal(result.captions[1].sequence,1)
  assert.match(result.validation.checks.join(','),/Word SEQ 域保存验证/)
  const output=unzipSync(result.bytes)
  const xml=new TextDecoder().decode(output['word/document.xml'])
  assert.match(xml,/SEQ PiFigure3 \\r 1 \\\* ARABIC/)
  assert.match(xml,/SEQ PiTable3 \\r 1 \\\* ARABIC/)
  assert.equal((xml.match(/w:fldCharType="begin"/g)||[]).length,2)
  assert.match(xml,/w:pStyle w:val="Caption"/)
  assert.equal(compareProtected(original,zipSync({...output,'word/document.xml':parts['word/document.xml']})).status,'passed')
})

test('caption insertion refuses ambiguous chapter and floating figures', async () => {
  const parts=unzipSync(documentFixture(richParagraph('普通标题','<w:pStyle w:val="Heading1"/>')+paragraph('目标')))
  const sourceXml=new TextDecoder().decode(parts['word/document.xml'])
  parts['word/document.xml']=new TextEncoder().encode(sourceXml.replace('</w:body>',
    '<w:p><w:r><w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:anchor/></w:drawing></w:r></w:p></w:body>'))
  assert.throws(()=>insertCaptions(zipSync(parts),[{kind:'figure',targetIndex:2,text:'浮动图'}]),/浮动图片/)
  const noChapterParts=unzipSync(documentFixture(richParagraph('普通标题','<w:pStyle w:val="Heading1"/>')+paragraph('图片段落')))
  const noChapterXml=new TextDecoder().decode(noChapterParts['word/document.xml'])
  noChapterParts['word/document.xml']=new TextEncoder().encode(noChapterXml.replace('</w:body>',
    '<w:p><w:r><w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:inline/></w:drawing></w:r></w:p></w:body>'))
  assert.throws(()=>insertCaptions(zipSync(noChapterParts),[{kind:'figure',targetIndex:2,text:'未知章节'}]),/章节号/)
})

test('caption creation commits a new document revision through the store', async () => {
  const folder=await mkdtemp(join(tmpdir(),'pi-caption-'))
  try {
    const parts=unzipSync(documentFixture(richParagraph('第三章 系统设计','<w:pStyle w:val="Heading1"/>')))
    const xml=new TextDecoder().decode(parts['word/document.xml'])
    const table='<w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>参数</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    parts['word/document.xml']=new TextEncoder().encode(xml.replace('<w:sectPr',`${table}<w:sectPr`))
    const store=new DocumentStore(folder), created=await store.create('题注测试.docx',zipSync(parts))
    const result=await store.addCaptions(created.id,1,[{kind:'table',targetIndex:0,text:'测试参数'}])
    assert.equal(result.document.revision,2)
    assert.equal(result.affected,1)
    const inspection=await inspect(await store.bytes(created.id))
    assert.equal(inspection.captions.length,1)
    assert.equal(inspection.captions[0].text,'表3-1 测试参数')
  } finally { await rm(folder,{recursive:true,force:true}) }
})

test('captions can be appended in separate calls and batch numbering follows document order', async () => {
  const parts=unzipSync(documentFixture(richParagraph('第十二章 测试','<w:pStyle w:val="Heading1"/>')))
  const table='<w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:p><w:r><w:t>数据</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
  parts['word/document.xml']=new TextEncoder().encode(new TextDecoder().decode(parts['word/document.xml']).replace('<w:sectPr',`${table}${table}<w:sectPr`))
  const source=zipSync(parts)
  const first=insertCaptions(source,[{kind:'table',targetIndex:0,text:'第一张'}])
  const second=insertCaptions(first.bytes,[{kind:'table',targetIndex:1,text:'第二张'}])
  assert.deepEqual((await inspect(second.bytes)).captions.map(p=>p.text),['表12-1 第一张','表12-2 第二张'])
  const batch=insertCaptions(source,[{kind:'table',targetIndex:1,text:'第二张'},{kind:'table',targetIndex:0,text:'第一张'}])
  assert.deepEqual((await inspect(batch.bytes)).captions.map(p=>p.text),['表12-1 第一张','表12-2 第二张'])
  assert.throws(()=>insertCaptions(first.bytes,[{kind:'table',targetIndex:0,text:'重复'}]),/已有同章|已有题注/)
})

test('single and multiple line spacing use auto units and reject zero exact spacing', async () => {
  const original=await fixture()
  await assert.rejects(format(original,[{scope:'all',paragraph:{lineSpacing:0}}]),/固定行距/)
  await assert.rejects(format(original,[{scope:'all',paragraph:{lineSpacing:12,lineSpacingMultiple:1}}]),/同时/)
  for(const multiple of [1,1.5,2]) {
    const out=await format(original,[{scope:'all',paragraph:{lineSpacingMultiple:multiple}}])
    const body=readEditorDocument(out.bytes)
    for(const p of body.paragraphs) {
      assert.equal(p.p.pPr.spacing.lineRule,'auto')
      assert.equal(p.p.pPr.spacing.line,240*multiple)
    }
  }
})

test('whole paragraph font formatting includes empty paragraph marks', async () => {
  const bytes=documentFixture(paragraph('正文')+richParagraph('','<w:rPr><w:sz w:val="24"/><w:b/></w:rPr>'))
  const out=await format(bytes,[{scope:'all',font:{name:'宋体',size:10.5,bold:false}}])
  const result=(readEditorDocument(out.bytes)).paragraphs
  assert.equal(result.length,2)
  assert.equal(result[1].text,'')
  for(const p of result) {
    assert.equal(p.font.size,10.5)
    assert.equal(p.font.bold,false)
    assert.equal(p.p.pPr.rPr.sz.val,21)
    assert.equal(p.p.pPr.rPr.rFonts.eastAsia,'宋体')
  }
})

test('citation range formatting preserves tabs and all kinds of breaks', async () => {
  const p={text:'前文\t第一行\u000b第二行\u000C分页\u000E分栏\r引用[1]后文\t末尾'}
  const original=documentFixture(richParagraph(p.text,'','<w:rFonts w:ascii="仿宋" w:hAnsi="仿宋"/>'))
  const out = await format(original,[{scope:'text',indices:[0],text:'[1]',font:{superscript:true}}])
  const result = (readEditorDocument(out.bytes)).paragraphs[0]
  assert.equal(result.text,p.text)
  assert.equal(findTextRanges(result,'[1]')[0].font.superscript,true)
  assert.equal(findTextRanges(result,'后文')[0].font.superscript,false)
  const xmlBefore = new TextDecoder().decode(unzipSync(original)['word/document.xml'])
  const xmlAfter = new TextDecoder().decode(unzipSync(out.bytes)['word/document.xml'])
  assert.equal((xmlBefore.match(/<w:tab\b/g)||[]).length,(xmlAfter.match(/<w:tab\b/g)||[]).length)
  assert.equal((xmlBefore.match(/<w:br\b/g)||[]).length,(xmlAfter.match(/<w:br\b/g)||[]).length)
})

test('requested heading style has an actual definition when absent in source', async () => {
  const parts = unzipSync(await fixture())
  parts['word/styles.xml'] = new TextEncoder().encode(new TextDecoder().decode(parts['word/styles.xml']).replace(/<w:style\b(?=[^>]*w:styleId="Heading1")[\s\S]*?<\/w:style>/g,''))
  const out = await format(zipSync(parts),[{scope:'indices',indices:[0],style:'Heading1'}])
  const styles = new TextDecoder().decode(unzipSync(out.bytes)['word/styles.xml'])
  assert.match(styles,/w:styleId="Heading1"/)
  assert.equal((await inspect(out.bytes)).paragraphs[0].heading,1)
})

test('explicit heading style overrides old direct body outline level', async () => {
  const out=await format(documentFixture(richParagraph('第一章 绪论','<w:outlineLvl w:val="9"/>')),[{scope:'all',paragraph:{outlineLevel:10}},{scope:'indices',indices:[0],style:'Heading1'}])
  assert.equal((await inspect(out.bytes)).paragraphs[0].heading,1)
})

test('range font changes only exact text across runs and survives sequential overlapping operations', async () => {
  const p={text:'一、提高统筹认识。后面的正文保持仿宋。'}
  const bytes=documentFixture('<w:p><w:r><w:t>一、提高</w:t></w:r><w:r><w:t>统筹认识。后面的正文保持仿宋。</w:t></w:r></w:p>')
  const result = await format(bytes, [
    { scope: 'all', font: { name: '仿宋', size: 16 } },
    { scope: 'text', indices: [0], text: '一、提高统筹认识。', font: { name: '黑体', size: 16 } },
  ])
  const body = readEditorDocument(result.bytes)
  assert.equal(body.paragraphs[0].text, p.text)
  assert.equal(findTextRanges(body.paragraphs[0],'一、提高统筹认识。')[0].font.name, '黑体')
  assert.equal(findTextRanges(body.paragraphs[0],'后面的正文保持仿宋。')[0].font.name, '仿宋')
  for (const r of findTextRanges(body.paragraphs[0],'一、提高统筹认识。')[0].runs) assert.equal(r.value.rPr.rFonts.eastAsia, '黑体')
  for (const r of findTextRanges(body.paragraphs[0],'后面的正文保持仿宋。')[0].runs) assert.equal(r.value.rPr.rFonts.eastAsia, '仿宋')
  await assert.rejects(format(result.bytes, [{scope:'text',indices:[0],text:'不存在',font:{bold:true}}]), /唯一匹配/)
})

test('ambiguous text is refused and explicit occurrence selects only one match', async () => {
  const bytes=documentFixture(richParagraph('标题。正文。标题。','','<w:rFonts w:ascii="仿宋" w:hAnsi="仿宋"/>'))
  await assert.rejects(format(bytes, [{scope:'text',indices:[0],text:'标题',font:{bold:true}}]), /唯一匹配/)
  const out = await format(bytes, [{scope:'text',indices:[0],text:'标题',occurrence:1,font:{bold:true}}])
  const matches = findTextRanges(readEditorDocument(out.bytes).paragraphs[0],'标题')
  assert.equal(matches[0].font.bold, false)
  assert.equal(matches[1].font.bold, true)
})

test('pagination flags preserve text and can be explicitly disabled', async () => {
  const bytes = await fixture()
  const flags = {keepNext:true, keepLines:true, widowControl:true, pageBreakBefore:false}
  const out = await format(bytes, [{scope:'indices',indices:[0],paragraph:flags}])
  assert.deepEqual((await inspect(out.bytes)).paragraphs[0].pagination, flags)
  const off = await format(out.bytes, [{scope:'indices',indices:[0],paragraph:{keepNext:false}}])
  assert.equal((await inspect(off.bytes)).paragraphs[0].pagination.keepNext, false)
})

test('failed range stage leaves original revision and all preceding operations uncommitted', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'pi-range-atomic-'))
  try {
    const store = new DocumentStore(directory)
    const original = await fixture()
    const doc = await store.create('局部失败.docx',original)
    await assert.rejects(store.apply(doc.id,1,[{scope:'all',font:{name:'仿宋'}},{scope:'text',indices:[1],text:'不存在的文字',font:{bold:true}}]),/唯一匹配/)
    assert.equal((await store.metadata(doc.id)).revision,1)
    assert.deepEqual(Buffer.from(await store.bytes(doc.id)),Buffer.from(original))
  } finally { await rm(directory,{recursive:true,force:true}) }
})

test('range formatting refuses fields rather than damaging protected inline structures', async () => {
  const parts = unzipSync(await fixture())
  const xml = new TextDecoder().decode(parts['word/document.xml'])
  parts['word/document.xml'] = new TextEncoder().encode(xml.replace('<w:t>正文保持原文 English 123。</w:t>', '<w:fldChar w:fldCharType="begin"/><w:t>正文保持原文 English 123。</w:t>'))
  await assert.rejects(format(zipSync(parts),[{scope:'text',indices:[1],text:'正文',font:{bold:true}}]),/复杂结构/)
})

test('inspection distinguishes Latin font from direct East Asian font and theme', async () => {
  const result=await inspect(documentFixture(richParagraph('中文 English','','<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="ＭＳ 明朝" w:eastAsiaTheme="minorEastAsia"/>')))
  assert.equal(result.paragraphs[0].font.name, 'Arial')
  assert.deepEqual(result.paragraphs[0].font.eastAsiaDirect, ['ＭＳ 明朝'])
  assert.deepEqual(result.paragraphs[0].font.eastAsiaThemes, ['minorEastAsia'])
})

test('font changes set East Asian fonts and clear overriding theme references on selected paragraphs', async () => {
  const source=documentFixture(['中文正文 English','保持原字体'].map(text=>richParagraph(text,'','<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="ＭＳ 明朝" w:eastAsiaTheme="minorEastAsia"/>')).join(''))
  const output=await format(source,[{scope:'indices',indices:[0],font:{name:'仿宋'}}])
  const paragraphs=readEditorDocument(output.bytes).paragraphs
  for (const run of paragraphs[0].runs) {
    assert.equal(run.value.rPr.rFonts.eastAsia, '仿宋')
    assert.equal(run.value.rPr.rFonts.eastAsiaTheme, undefined)
    assert.equal(run.value.rPr.rFonts.ascii, '仿宋')
  }
  assert.equal(paragraphs[1].runs[0].value.rPr.rFonts.eastAsia, 'ＭＳ 明朝')
  assert.equal(paragraphs[0].text, '中文正文 English')
})

test('navigation outline level survives save without changing paragraph appearance', async () => {
  const original=documentFixture(richParagraph('四、评价指标','<w:ind w:firstLine="640"/>','<w:rFonts w:ascii="仿宋" w:hAnsi="仿宋"/><w:sz w:val="32"/><w:color w:val="0000FF"/>')+paragraph('目录数据完整率不低于98%。'))
  const body=readEditorDocument(original)
  const output = await format(original, [{ scope: 'indices', indices: [0], paragraph: { outlineLevel: 1 } }])
  const result = readEditorDocument(output.bytes)
  assert.equal(result.paragraphs[0].outlineLevel, 1)
  assert.equal(result.paragraphs[1].outlineLevel, 10)
  assert.equal(result.paragraphs[0].font.name, '仿宋')
  assert.equal(result.paragraphs[0].font.size, 16)
  assert.equal(result.paragraphs[0].font.color, '#0000FF')
  assert.equal(result.paragraphs[0].firstLineIndent, 32)
  assert.equal(result.text, body.text)
  assert.equal((await inspect(output.bytes)).paragraphs[0].heading, 1)
  const stylesXml = new TextDecoder().decode(unzipSync(output.bytes)['word/styles.xml'])
  assert.match(result.paragraphs[0].styleId, /^PiNav1_/)
  assert.ok(stylesXml.includes(`w:styleId="${result.paragraphs[0].styleId}"`))
  assert.match(stylesXml, /<w:basedOn w:val="Normal"\s*\/>/)
  assert.match(stylesXml, /<w:outlineLvl w:val="0"\s*\/>/)
})

test('formats commit immutable revisions, enforce revisions and restore as a new version', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'pi-docx-store-'))
  try {
    const store = new DocumentStore(directory)
    const original = await fixture()
    const document = await store.create('测试.docx', original)
    assert.equal(document.revision, 1)

    const result = await store.apply(document.id, 1, [{ scope: 'headings', heading: 1,
      font: { size: 18, bold: true }, paragraph: { alignment: 'Centered' } }])
    assert.equal(result.document.revision, 2)
    assert.equal(result.affected, 1)
    assert.equal(result.validation.status, 'partial')
    assert.equal(result.validation.coverage.visualLayout, 'not-verified')
    assert.deepEqual((await store.metadata(document.id)).versions[1].validation, result.validation)
    const formatted = await store.bytes(document.id, 2)
    const formattedBody = readEditorDocument(formatted)
    assert.equal(formattedBody.paragraphs[0].font.size, 18)
    assert.equal(formattedBody.paragraphs[0].font.bold, true)
    assert.equal(formattedBody.paragraphs[0].alignment, 'Centered')
    assert.deepEqual(Buffer.from(await store.bytes(document.id, 1)), Buffer.from(original))

    await assert.rejects(store.apply(document.id, 1, [{ scope: 'body', font: { size: 12 } }]), { status: 409 })
    await assert.rejects(store.apply(document.id, 2, [{ scope: 'all', font: { unsupportedProperty: true } }]), /不支持字体属性/)
    assert.equal((await store.metadata(document.id)).revision, 2)

    const restored = await store.restore(document.id, 2, 1)
    assert.equal(restored.revision, 3)
    assert.deepEqual(Buffer.from(await store.bytes(document.id)), Buffer.from(original))
    assert.deepEqual(Buffer.from(await store.bytes(document.id, 2)), Buffer.from(formatted))
  } finally { await rm(directory, { recursive: true, force: true }) }
})

test('manual editor save creates a version but stale revision cannot overwrite it', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'pi-docx-store-'))
  try {
    const store = new DocumentStore(directory)
    const original = await fixture()
    const document = await store.create('manual.docx', original)
    const parts=unzipSync(original)
    parts['word/document.xml']=new TextEncoder().encode(new TextDecoder().decode(parts['word/document.xml']).replace('正文保持原文 English 123。','正文保持原文 English 123。 已修改'))
    const edited=zipSync(parts)
    const saved = await store.save(document.id, 1, edited)
    assert.equal(saved.revision, 2)
    assert.equal(saved.versions[1].validation.status, 'partial')
    await assert.rejects(store.save(document.id, 1, original), { status: 409 })
    assert.equal((await store.metadata(document.id)).revision, 2)
  } finally { await rm(directory, { recursive: true, force: true }) }
})

test('structure inspection reports tables, drawing elements and section settings', async () => {
  const source=unzipSync(documentFixture(richParagraph('结构索引标题','<w:pStyle w:val="Heading1"/>')))
  const xml = new TextDecoder().decode(source['word/document.xml'])
  const enriched = xml.replace('<w:document ', '<w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ')
    .replace('<w:sectPr>', '<w:p><w:pPr><w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr></w:pPr></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>表格文本</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>单元格</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:drawing xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:blip r:embed="image1"/></w:drawing><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x</m:t></m:r></m:oMath></w:r></w:p><w:sectPr>')
  source['word/document.xml'] = new TextEncoder().encode(enriched)
  source['word/footnotes.xml'] = new TextEncoder().encode('<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="1"><w:p/></w:footnote></w:footnotes>')
  source['word/endnotes.xml'] = new TextEncoder().encode('<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:id="1"><w:p/></w:endnote></w:endnotes>')
  const result = await inspect(zipSync(source))
  assert.equal(result.sections.length, 2)
  assert.equal(result.sections[0].page.orientation, 'landscape')
  assert.deepEqual(result.sections[0].margins, { top: 720, right: 720, bottom: 720, left: 720, header: 0, footer: 0, gutter: 0 })
  assert.deepEqual(result.sections[0].headers, [{ type: 'default', relationshipId: 'rIdHeader' }])
  assert.equal(result.tables.length, 1)
  assert.equal(result.tables[0].rows, 1)
  assert.equal(result.tables[0].columns, 2)
  assert.equal(result.tables[0].text, '表格文本单元格')
  assert.equal(result.nestedParagraphs.length, 2)
  assert.equal(result.elements.images, 1)
  assert.equal(result.elements.equations, 1)
  assert.equal(result.elements.footnotes, 1)
  assert.equal(result.elements.endnotes, 1)
})

test('section formatting updates only requested page properties and rejects invalid indices', async () => {
  const original = await fixture()
  const changed = await format(original, [{ scope: 'sections', sectionIndices: [0],
    page: { orientation: 'landscape' }, margins: { top: 900, gutter: 180 } }])
  const result = await inspect(changed.bytes)
  assert.equal(changed.affected, 1)
  assert.equal(result.sections[0].page.orientation, 'landscape')
  assert.equal(result.sections[0].margins.top, 900)
  assert.equal(result.sections[0].margins.gutter, 180)
  assert.equal(result.sections[0].margins.left, 1440)
  assert.equal(result.paragraphs[0].text, '一级标题：验证文档')
  assert.equal(result.paragraphs[1].text, '正文保持原文 English 123。')
  await assert.rejects(format(original, [{ scope: 'sections', sectionIndices: [1], margins: { left: 720 } }]), /分节索引无效/)
  await assert.rejects(format(original, [{ scope: 'sections', sectionIndices: [0], page: { widthTwips: 500 } }]), /页面宽高/)
})

test('new blank documents include an editable empty paragraph', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'pi-docx-store-'))
  try {
    const store = new DocumentStore(directory)
    const document = await store.create()
    const structure = await inspect(await store.bytes(document.id))
    assert.equal(structure.paragraphs.length, 1)
    assert.equal(structure.paragraphs[0].text, '')
    assert.equal(document.revision, 1)
  } finally { await rm(directory, { recursive: true, force: true }) }
})


test('character effects round trip and mutually exclusive scripts are rejected', async () => {
  const original = await fixture()
  for (const font of [{ underline: 'Single' }, { strikeThrough: true }, { superscript: true }, { subscript: true }]) {
    const result = await format(original, [{ scope: 'body', font }])
    const body = readEditorDocument(result.bytes)
    for (const [key, value] of Object.entries(font)) assert.equal(body.paragraphs[1].font[key], value)
  }
  await assert.rejects(format(original, [{ scope: 'body', font: { subscript: true, superscript: true } }]), /不能同时/)
})


test('hex colors accept optional hash and round trip without changing text', async () => {
  const original = await fixture()
  for (const color of ['FF0000', '000000', '#FF0000', '#000000']) {
    const operations = [{ scope: 'body', font: { color } }]
    const result = await format(original, operations)
    const body = readEditorDocument(result.bytes)
    assert.equal(body.paragraphs[1].font.color, color.startsWith('#') ? color : `#${color}`)
    assert.equal(body.paragraphs[1].text, '正文保持原文 English 123。')
    assert.equal(operations[0].font.color, color)
  }
  for (const color of ['red', '#GG0000', '000', 0, null]) {
    await assert.rejects(format(original, [{ scope: 'body', font: { color } }]), /颜色/)
  }
})


test('rejects margins that leave no writable page area', async () => {
  const original = await fixture()
  await assert.rejects(format(original, [{ scope: 'sections', sectionIndices: [0], margins: { left: 10000, right: 10000 } }]), /页边距总和/)
})

 test('workflow emits only committed revisions and cancellation leaves bytes unchanged', async () => {
  const folder=await mkdtemp(join(tmpdir(),'workflow-'))
  try {
    const store=new DocumentStore(folder), created=await store.create('工作流.docx',await fixture())
    const events=[]; const stop=store.onCommit(event=>events.push(event))
    const original=await store.bytes(created.id)
    const controller=new AbortController(); controller.abort()
    await assert.rejects(store.apply(created.id,1,[{scope:'all',font:{size:12}}],controller.signal))
    assert.equal(events.length,0)
    assert.deepEqual(await store.bytes(created.id),original)
    const output=await store.apply(created.id,1,[{scope:'all',font:{size:12}}])
    assert.equal(output.document.revision,2)
    assert.deepEqual(events.map(e=>e.revision),[2])
    await assert.rejects(store.apply(created.id,1,[{scope:'all',font:{size:14}}]),/版本/)
    assert.equal(events.length,1)
    stop()
  } finally {await rm(folder,{recursive:true,force:true})}
})

test('canonical fingerprints stay bounded for deeply nested drawing structures', () => {
  const make = text => parsePackageXml(new TextEncoder().encode('<root>'+ '<group>'.repeat(18) + '<text>'+text+'</text>' + '</group>'.repeat(18)+'</root>')).documentElement
  const value=canonical(make('picture'))
  assert.ok(value.length<=64, `fingerprint grew to ${value.length} characters`)
  assert.equal(value,canonical(make('picture')))
  assert.notEqual(value,canonical(make('changed')))
})


test('serialization equivalence preserves boolean meaning and embedded drawing data', () => {
  const fingerprint = source => canonical(parsePackageXml(new TextEncoder().encode(source)).documentElement)
  const cases = [
    ['http://schemas.openxmlformats.org/drawingml/2006/main', 'picLocks', 'noChangeArrowheads'],
    ['http://schemas.microsoft.com/office/drawing/2010/main', 'useLocalDpi', 'val'],
    ['http://schemas.microsoft.com/office/word/2010/wordprocessingShape', 'cNvSpPr', 'txBox'],
    ...['anchorCtr','fromWordArt','rtlCol','compatLnSpc','forceAA','spcFirstLastPara'].map(name => ['http://schemas.microsoft.com/office/word/2010/wordprocessingShape', 'bodyPr', name]),
  ]
  for (const [ns, tag, attribute] of cases) {
    const node = value => `<x:${tag} xmlns:x="${ns}" ${attribute}="${value}"/>`
    assert.equal(fingerprint(node('0')), fingerprint(node('false')))
    assert.equal(fingerprint(node('1')), fingerprint(node('true')))
    assert.notEqual(fingerprint(node('0')), fingerprint(node('1')))
  }
  const link = value => `<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:history="${value}" w:anchor="bookmark"/>`
  assert.equal(fingerprint(link('1')), fingerprint(link('true')))
  assert.notEqual(fingerprint(link('0')), fingerprint(link('true')))
  const shape = value => `<v:shape xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" o:gfxdata="${value}"/>`
  assert.equal(fingerprint(shape('YWJj&#10;ZA==')), fingerprint(shape('YWJjZA==')))
  assert.notEqual(fingerprint(shape('YWJjZA==')), fingerprint(shape('YWJjZQ==')))
  assert.notEqual(fingerprint('<unknown val="0"/>'), fingerprint('<unknown val="false"/>'))
})


test('range formatting follows body order through content controls and tables with revision metadata', async () => {
  const parts = unzipSync(await fixture())
  let source = new TextDecoder().decode(parts['word/document.xml'])
  source = source.replace('<w:body>', '<w:body><w:sdt><w:sdtPr/><w:sdtContent><w:p><w:r><w:t>目录</w:t></w:r></w:p></w:sdtContent></w:sdt><w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:tcPr/><w:p><w:r w:rsidR="12345678"><w:t>关键词 test</w:t></w:r></w:p></w:tc></w:tr></w:tbl>')
  parts['word/document.xml'] = new TextEncoder().encode(source)
  const bytes = zipSync(parts)
  const before = (readEditorDocument(bytes)).paragraphs
  assert.equal(before[1].text, '关键词 test')
  const output = await format(bytes, [{scope:'text', indices:[1], text:'关键词', font:{bold:true}}])
  const after = (readEditorDocument(output.bytes)).paragraphs
  assert.deepEqual(after.map(p=>p.text), before.map(p=>p.text))
  assert.equal(findTextRanges(after[1],'关键词')[0].font.bold, true)
  assert.notEqual(findTextRanges(after[1],'test')[0].font.bold, true)
  const document = parsePackageXml(unzipSync(output.bytes)['word/document.xml'])
  const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
  const cell = document.getElementsByTagNameNS(ns, 'tc')[0]
  for (const run of Array.from(cell.getElementsByTagNameNS(ns, 'r'))) assert.equal(run.getAttributeNS(ns, 'rsidR'), '12345678')
})


test('section page and margins may coexist but paragraph fields require a separate operation', async () => {
  const original = await fixture()
  const section = {scope:'sections', sectionIndices:[0], page:{widthTwips:11906,heightTwips:16838,orientation:'portrait'}, margins:{top:1440,bottom:1440,left:1800,right:1800}}
  await assert.rejects(format(original, [{...section, paragraph:{spaceBefore:0,spaceAfter:0}}]), /拆成独立的段落操作/)
  const result = await format(original, [section, {scope:'body',paragraph:{spaceBefore:0,spaceAfter:0}}])
  assert.equal(result.validation.status, 'passed')
  const data = await inspect(result.bytes)
  assert.equal(data.sections[0].margins.left, 1800)
})


test('protected comparison detects changes to whitespace-only text runs', async () => {
  const parts = unzipSync(await fixture())
  const source = new TextDecoder().decode(parts['word/document.xml']).replace('正文保持原文 English 123。', ' ')
  parts['word/document.xml'] = new TextEncoder().encode(source)
  const before = zipSync(parts)
  parts['word/document.xml'] = new TextEncoder().encode(source.replace('> </w:t>', '>  </w:t>'))
  assert.throws(() => compareProtected(before, zipSync(parts)), /正文结构/)
})

test('range selectors cannot silently broaden an operation', async () => {
  const bytes = await fixture()
  await assert.rejects(format(bytes, [{scope:'all',indices:[0],font:{bold:true}}]), /indices/)
})


for (const verticalAlignment of ['Top', 'Center', 'Bottom']) {
  test(`native cell alignment targets the second table: ${verticalAlignment}`, async () => {
    const source = documentFixture(table([['保留表格']]) + table([['目标', '保留单元格']]))
    const before = await inspect(source)
    const out = await format(source, [{
      scope: 'tableCells', tableIndex: 1,
      cellTargets: [{ rowIndex: 0, cellIndex: 0 }, { rowIndex: 0, cellIndex: 0 }],
      table: { verticalAlignment },
    }])
    const after = await inspect(out.bytes)
    assert.deepEqual(after.paragraphs.map(p => p.text), before.paragraphs.map(p => p.text))
    assert.deepEqual(after.tables[0], before.tables[0])
    assert.equal(after.tables[1].cellDetails[0].verticalAlignment, verticalAlignment.toLowerCase())
    assert.equal(after.tables[1].cellDetails[1].verticalAlignment, null)
  })
}
