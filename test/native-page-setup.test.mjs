import test from 'node:test'
import assert from 'node:assert/strict'
import { applySectionOperations } from '../server/document/ooxml-sections.mjs'
import { paragraphDocument } from './helpers/docx-fixture.mjs'
import { unpack, xml, W } from '../server/document/package.mjs'

for (const gutter of [undefined, 0, 240]) {
  test(`page settings preserve content with gutter=${gutter}`, () => {
    const source = paragraphDocument(['原文 English 123'])
    const result = applySectionOperations(source, [{
      sectionIndices: [0],
      page: { widthTwips: 11906, heightTwips: 16838, orientation: 'portrait' },
      margins: { top: 1501, left: 1800, ...(gutter === undefined ? {} : { gutter }) },
    }])
    const document = xml(unpack(result.bytes)['word/document.xml'])
    const margin = document.getElementsByTagNameNS(W, 'pgMar')[0]
    assert.equal(margin.getAttributeNS(W, 'top'), '1501')
    assert.equal(margin.getAttributeNS(W, 'left'), '1800')
    if (gutter !== undefined) assert.equal(margin.getAttributeNS(W, 'gutter'), String(gutter))
    assert.equal([...document.getElementsByTagNameNS(W, 't')].map(n => n.textContent).join(''), '原文 English 123')
  })
}

test('gutter-only edits retain the page settings', () => {
  const source = paragraphDocument(['仅修改装订线'])
  const result = applySectionOperations(source, [{ sectionIndices: [0], margins: { gutter: 120 } }])
  const document = xml(unpack(result.bytes)['word/document.xml'])
  assert.equal(document.getElementsByTagNameNS(W, 'pgMar')[0].getAttributeNS(W, 'gutter'), '120')
  assert.equal(document.getElementsByTagNameNS(W, 't')[0].textContent, '仅修改装订线')
})
