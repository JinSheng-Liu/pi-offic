import test from 'node:test'
import assert from 'node:assert/strict'
import {reviewLayout} from '../src/layoutReview.ts'
const paragraph = (id,text,y=0,height=100) => ({kind:'paragraph',paragraphId:id,box:{y,height},lines:[{spans:[{text}]}]})
const page = (index,fragments) => ({index,fragments,contentBox:{height:600}})
test('native page review detects isolated heading without guessing pagination', () => {
  const bad = [page(0,[paragraph('h','六、组织保障',500)]),page(1,[paragraph('b','正文')])]
  assert.ok(reviewLayout(bad,new Set(['h'])).some(f=>f.message.includes('分处两页')))
  const fixed = [page(0,[paragraph('h','六、组织保障'),paragraph('b','正文',100,500)])]
  assert.equal(reviewLayout(fixed,new Set(['h'])).length,0)
})
test('short title tail and nonfinal page whitespace are review hints', () => {
  const title = paragraph('h','项目实施')
  title.lines.push({spans:[{text:'要点'}]})
  const results = reviewLayout([page(0,[title]),page(1,[paragraph('b','正文')])],new Set(['h']))
  assert.ok(results.some(f=>f.message.includes('最后一行')))
  assert.ok(results.some(f=>f.message.includes('留白')))
})
