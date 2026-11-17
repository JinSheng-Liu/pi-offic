import test from 'node:test'
import assert from 'node:assert/strict'
import {format} from '../server/engine.mjs'
import {documentFixture, paragraph} from './helpers/docx-fixture.mjs'
import {W, M, xml, canonical} from '../server/document/package.mjs'
import {unzipSync} from 'fflate'
import {validateOperations} from '../server/document/rules.mjs'

test('renumberHeadings corrects explicit heading prefixes only', async () => {
  const source = documentFixture([
    '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>第1章 绪论</w:t></w:r></w:p>',
    '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>1.1 背景</w:t></w:r></w:p>',
    '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>第3章 结果</w:t></w:r></w:p>'
  ].join(''))
  const result = await format(source, [{scope:'renumberHeadings'}])
  const text = new TextDecoder().decode(unzipSync(result.bytes)['word/document.xml'])
  assert.match(text, /第1章 绪论/)
  assert.match(text, /1\.1 背景/)
  assert.match(text, /第2章 结果/)
})

test('numberEquations adds native SEQ fields to independent display equations', async () => {
  const source = documentFixture([
    '<w:p><m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara></w:p>',
    '<w:p><m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath><m:r><m:t>y</m:t></m:r></m:oMath></m:oMathPara></w:p>'
  ].join(''))
  const result = await format(source, [{scope:'numberEquations', format:'arabic', prefix:'(', suffix:')'}])
  const document = xml(unzipSync(result.bytes)['word/document.xml'])
  assert.equal(document.getElementsByTagNameNS(W, 'instrText').length, 2)
  assert.match(document.getElementsByTagNameNS(W, 'instrText')[0].textContent, /SEQ PiEquation/)
  const tabs = Array.from(document.getElementsByTagNameNS(W, 'tab'))
  assert.equal(tabs.filter(node => node.parentNode.localName === 'tabs').length, 2)
  assert.equal(tabs.filter(node => node.parentNode.localName === 'r').length, 2)
  for (const tab of tabs.filter(node => node.parentNode.localName === 'tabs')) {
    assert.equal(tab.getAttributeNS(W, 'val'), 'right')
    assert.equal(tab.getAttributeNS(W, 'pos'), '9026')
  }
  const before = xml(unzipSync(source)['word/document.xml'])
  assert.deepEqual(Array.from(document.getElementsByTagNameNS(M, 'oMathPara')).map(canonical), Array.from(before.getElementsByTagNameNS(M, 'oMathPara')).map(canonical))
  await assert.rejects(format(result.bytes, [{scope:'numberEquations'}]), /重复编号|已有/)

})

test('numbering parameters cannot escape their scope or mix with ordinary formatting', () => {
  assert.throws(() => validateOperations([{scope:'body', font:{size:12}, prefix:'('}]), /仅用于/)
  assert.throws(() => validateOperations([{scope:'numberEquations', format:'invalid'}]), /无效/)
  assert.throws(() => validateOperations([{scope:'numberEquations', prefix:'\n'}]), /无效/)
  assert.throws(() => validateOperations([{scope:'numberEquations'}, {scope:'renumberHeadings'}]), /独立/)
  assert.throws(() => validateOperations([{scope:'renumberHeadings', indices:[0]}]), /仅指定/)
})

test('heading correction preserves ordinary paragraphs and refuses native lists', async () => {
  const heading = '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>3 章节</w:t></w:r></w:p>'
  const source = documentFixture(paragraph('2026 正文保持') + heading)
  const result = await format(source, [{scope:'renumberHeadings'}])
  const before = xml(unzipSync(source)['word/document.xml'])
  const after = xml(unzipSync(result.bytes)['word/document.xml'])
  assert.equal(canonical(before.getElementsByTagNameNS(W,'p')[0]), canonical(after.getElementsByTagNameNS(W,'p')[0]))
  await assert.rejects(format(result.bytes, [{scope:'renumberHeadings'}]), /未发现/)
  const native = documentFixture(heading.replace('</w:pPr>', '<w:numPr><w:numId w:val="1"/></w:numPr></w:pPr>'))
  await assert.rejects(format(native, [{scope:'renumberHeadings'}]), /原生列表/)
})

test('formula numbering preserves numeric affixes and rejects mixed content', async () => {
  const math = `<m:oMathPara xmlns:m="${M}"><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara>`
  const source = documentFixture(`<w:p>${math}</w:p>`)
  const result = await format(source, [{scope:'numberEquations', format:'upperRoman', prefix:'1', suffix:'1'}])
  const document = xml(unzipSync(result.bytes)['word/document.xml'])
  assert.deepEqual(Array.from(document.getElementsByTagNameNS(W,'t')).map(node => node.textContent), ['1','I','1'])
  for (const [name, bytes] of Object.entries(unzipSync(source))) {
    if (name !== 'word/document.xml') assert.deepEqual(unzipSync(result.bytes)[name], bytes)
  }
  const mixed = documentFixture(`<w:p>${math}<w:r><w:t>说明</w:t></w:r></w:p>`)
  await assert.rejects(format(mixed, [{scope:'numberEquations'}]), /混排/)
  const controller = new AbortController()
  controller.abort()
  await assert.rejects(format(source, [{scope:'numberEquations'}], controller.signal), /abort/i)
})

