import test from 'node:test'
import assert from 'node:assert/strict'
import {createBlankDocument,executeAutomationBatch,openAutomationDocument} from '../server/document/editor-automation.mjs'
import {unpack,xml,W,canonical} from '../server/document/package.mjs'
function ask(host,operation) {
  const result=host.execute({operations:[operation]})
  assert.equal(result.ok,true,JSON.stringify(result.results))
  return result.results[0].value
}
function firstParagraph(host) {
  const document=ask(host,{op:'getDocument'}).handle
  const body=ask(host,{op:'getBody',document}).handle
  return ask(host,{op:'getParagraphs',body}).handles[0]
}
test('blank document is A4 after save and reopen',()=>{
  const bytes=createBlankDocument()
  const size=xml(unpack(bytes)['word/document.xml']).getElementsByTagNameNS(W,'pgSz')[0]
  assert.equal(Number(size.getAttributeNS(W,'w')),11906)
  assert.equal(Number(size.getAttributeNS(W,'h')),16838)
  const host=openAutomationDocument(bytes)
  try { assert.ok(firstParagraph(host)) } finally { host.dispose() }
})
test('formatting persists alignment and indent; spacing API retains the existing rule',()=>{
  const result=executeAutomationBatch(createBlankDocument(),host=>[
    {op:'setParagraphFormat',paragraph:{paragraph:firstParagraph(host)},format:{alignment:'Justified',firstLineIndent:24,lineSpacing:18,spaceBefore:0,spaceAfter:0}},
  ])
  const host=openAutomationDocument(result.bytes)
  try {
    const format=ask(host,{op:'getParagraphFormat',paragraph:{paragraph:firstParagraph(host)}}).format
    assert.equal(format.alignment,'Justified')
    assert.equal(format.firstLineIndent,24)
    assert.equal(format.lineSpacing,18)
    const spacing=xml(unpack(result.bytes)['word/document.xml']).getElementsByTagNameNS(W,'spacing')[0]
    assert.equal(spacing.getAttributeNS(W,'line'),'360')
    assert.equal(spacing.getAttributeNS(W,'lineRule'),'auto')
  } finally { host.dispose() }
})
test('failed batch leaves prior formatting and revision intact',()=>{
  const host=openAutomationDocument(createBlankDocument())
  try {
    const paragraph=firstParagraph(host),revision=host.revision()
    const before=ask(host,{op:'getParagraphFormat',paragraph:{paragraph}}).format
    const result=host.execute({expectedRevision:revision,operations:[
      {op:'setParagraphFormat',paragraph:{paragraph},format:{alignment:'Centered'}},
      {op:'setPageSetup',section:paragraph,setup:{pageWidth:600}},
    ]})
    assert.equal(result.ok,false)
    assert.equal(host.revision(),revision)
    assert.deepEqual(ask(host,{op:'getParagraphFormat',paragraph:{paragraph}}).format,before)
  } finally { host.dispose() }
})
test('stale revision cannot overwrite a newer formatting edit',()=>{
  const host=openAutomationDocument(createBlankDocument())
  try {
    const paragraph=firstParagraph(host),revision=host.revision()
    ask(host,{op:'setParagraphFormat',paragraph:{paragraph},format:{alignment:'Centered'}})
    const result=host.execute({expectedRevision:revision,operations:[
      {op:'setParagraphFormat',paragraph:{paragraph},format:{alignment:'Justified'}},
    ]})
    assert.equal(result.ok,false)
    assert.equal(ask(host,{op:'getParagraphFormat',paragraph:{paragraph}}).format.alignment,'Centered')
  } finally { host.dispose() }
})


test('Ignorable prefix order is equivalent but namespace changes remain protected',()=>{
  const parse=text=>xml(new TextEncoder().encode(text)).documentElement
  const before=parse('<root xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:a="urn:a" xmlns:b="urn:b" mc:Ignorable="a b"/>')
  const reordered=parse('<root xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:a="urn:a" xmlns:b="urn:b" mc:Ignorable="b a"/>')
  const changed=parse('<root xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:a="urn:changed" xmlns:b="urn:b" mc:Ignorable="b a"/>')
  assert.equal(canonical(before),canonical(reordered))
  assert.notEqual(canonical(before),canonical(changed))
})
