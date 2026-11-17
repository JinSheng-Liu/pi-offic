import test from 'node:test'
import assert from 'node:assert/strict'
import {documentCoverage} from '../server/document/coverage.mjs'
test('coverage distinguishes heading evidence, missing boundaries and blocked captions from completion',()=>{
  const data={paragraphs:[{index:0,text:'摘要',style:'TOC1'},{index:1,text:'摘 要',heading:1},{index:2,text:'研究内容提及参考文献',heading:0},{index:3,text:'第一章 绪论',heading:1},{index:4,text:'参考文献',tableNestingLevel:1}],tables:[{}],elements:{images:2},captions:[],floatingObjects:[{objectIndex:7,textBoxCount:1,flowCapability:{supported:false,reason:'多段题注'}}]}
  const r=documentCoverage(data)
  assert.deepEqual(r.regions.find(x=>x.region==='abstract').headingIndices,[1])
  assert.equal(r.regions.find(x=>x.region==='references').status,'not-identified')
  assert.deepEqual(r.regions.find(x=>x.region==='body').headingIndices,[3])
  assert.equal(r.objects[0].status,'present-unverified')
  assert.equal(r.blockedTextboxes[0].objectIndex,7)
  assert.equal(r.requirementsVerified,false)
  data.paragraphs.push({index:5,text:'摘要'})
  assert.equal(documentCoverage(data).regions.find(x=>x.region==='abstract').status,'ambiguous')
})
