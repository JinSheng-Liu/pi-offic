import test from 'node:test'
import assert from 'node:assert/strict'
import {zipSync} from 'fflate'
import {createBlankDocument} from '../server/document/editor-automation.mjs'
import {readEditorDocument} from '../server/document/editor-reader.mjs'
import {paragraphRole} from '../server/document/paragraph-role.mjs'
import {format} from '../server/document/editor-adapter.mjs'
import {unpack,xml,W,canonical} from '../server/document/package.mjs'

function fixture(body) {
  const parts=unpack(createBlankDocument())
  parts['word/document.xml']=new TextEncoder().encode(`<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`)
  return zipSync(parts)
}
test('reader preserves symbol codes and special text without inventing a glyph',async()=>{
  const source=fixture('<w:p><w:r><w:t>A</w:t><w:sym w:font="Wingdings 2" w:char="00A3"/><w:noBreakHyphen/><w:softHyphen/><w:tab/><w:br/><w:t>B</w:t></w:r></w:p>')
  assert.equal(readEditorDocument(source).paragraphs[0].text,'A\u00a3\u2011\u00ad\t\u000bB')
  const result=await format(source,[{scope:'all',paragraph:{lineSpacingMultiple:1.5}}])
  const symbol=bytes=>xml(unpack(bytes)['word/document.xml']).getElementsByTagNameNS(W,'sym')[0]
  assert.equal(canonical(symbol(source)),canonical(symbol(result.bytes)))
  await assert.rejects(format(source,[{scope:'text',indices:[0],text:'A',font:{bold:true}}]),/复杂结构/)
})
test('editor formatting preserves unrelated indentation and spacing attributes',async()=>{
  const source=fixture('<w:p><w:pPr><w:ind w:left="480" w:firstLineChars="200"/><w:spacing w:before="120" w:after="80" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:t>正文 English</w:t></w:r></w:p>')
  const result=await format(source,[{scope:'all',font:{name:'宋体',latinName:'Times New Roman'},paragraph:{firstLineIndent:-24,lineSpacing:18}}])
  const p=readEditorDocument(result.bytes).paragraphs[0]
  assert.equal(p.leftIndent,24)
  assert.equal(p.firstLineIndent,-24)
  assert.equal(p.spaceBefore,6)
  assert.equal(p.spaceAfter,4)
  assert.equal(p.p.pPr.ind.firstLineChars,undefined)
  assert.equal(p.p.pPr.spacing.line,360)
  assert.equal(p.p.pPr.spacing.lineRule,'exact')
  assert.equal(p.runs[0].value.rPr.rFonts.eastAsia,'宋体')
  assert.equal(p.runs[0].value.rPr.rFonts.ascii,'Times New Roman')
})

test('generated TOC paragraphs are excluded from body formatting without TOC styles',async()=>{
  const source=fixture('<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p><w:p><w:r><w:t>第一章 1</w:t></w:r></w:p><w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p><w:p><w:r><w:t>实际正文</w:t></w:r></w:p>')
  const before=readEditorDocument(source)
  assert.equal(paragraphRole(before.paragraphs[1]),'toc')
  const result=await format(source,[{scope:'body',font:{size:16}}])
  assert.equal(result.affected,1)
  assert.ok(result.replay)
  const after=readEditorDocument(result.bytes)
  assert.equal(after.paragraphs[3].font.size,16)
  assert.equal(after.paragraphs[1].font.size,before.paragraphs[1].font.size)
})
