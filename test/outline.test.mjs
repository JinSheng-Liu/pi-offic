import test from 'node:test'
import assert from 'node:assert/strict'
import { recognizeHeading, visibleOutline } from '../src/outline.ts'

test('recognizes explicit heading patterns without interpreting mixed prose, dates or numbers as headings', () => {
  for (const value of ['第一章 绪论', '一、工作目标', '摘要', 'Abstract', '参考文献']) assert.equal(recognizeHeading(value), 0)
  assert.equal(recognizeHeading('1.1 项目建设'), 1)
  assert.equal(recognizeHeading('1.1.1 数据共享'), 2)
  assert.equal(recognizeHeading('第一章 节能目标'), 0)
  assert.equal(recognizeHeading('第一节 研究对象'), 1)
  for (const value of ['2026年9月15日', '98.5%', '一、工作目标。本年度我们完成了多项重点任务。', '这是没有标题标记的普通正文', '']) assert.equal(recognizeHeading(value), null)
})

test('collapse hides only descendants and retains independently collapsed branches', () => {
  const items = [
    { text: '第一章', level: 0, blockId: 'a' }, { text: '小节', level: 1, blockId: 'b' },
    { text: '细节', level: 2, blockId: 'c' }, { text: '第二章', level: 0, blockId: 'd' },
    { text: '小节', level: 1, blockId: 'e' },
  ]
  assert.deepEqual(visibleOutline(items, new Set(['a'])).map(item => item.blockId), ['a', 'd', 'e'])
  assert.deepEqual(visibleOutline(items, new Set(['b'])).map(item => item.blockId), ['a', 'b', 'd', 'e'])
})
