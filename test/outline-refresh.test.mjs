import test from 'node:test'
import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import vm from 'node:vm'
import ts from 'typescript'

function mountOutline(scroll) {
  const events = {}, watchers = [], timers = []
  const loading = { value: false }
  let headings = [{ text: '第一章', level: 0, blockId: 'a' }]
  const editor = { value: { getOutline: () => headings, query: () => [], findMatches: () => [] } }
  const file = readFileSync(new URL('../src/DocumentOutline.vue', import.meta.url), 'utf8')
  const script = file.match(/<script setup lang="ts">([\s\S]*?)<\/script>/)[1].replace(/^import .*$/gm, '')
  const context = vm.createContext({
    defineProps() {}, ref: value => ({ value }), computed: getter => ({ get value() { return getter() } }),
    nextTick: () => Promise.resolve(), onMounted() {}, onUnmounted() {},
    watch: (sources, callback, options) => { watchers.push(callback); if(options?.immediate) callback() },
    useDocxEditor: () => editor, useEditorState: () => loading,
    useEditorEvent: (name, callback) => { events[name] = callback },
    setTimeout: callback => { timers.push(callback); return timers.length }, clearTimeout() {},
    ...(scroll ? { document: { querySelector: () => scroll } } : {}),
  })
  vm.runInContext(ts.transpile(script + '\n globalThis.state = { items, collapsed, activeId, updateActive };', { target: ts.ScriptTarget.ES2022 }), context)
  return { state: context.state, loading, events, watchers, timers, setHeadings: value => { headings = value } }
}

test('outline stays visible during change debounce and reflects actual title removal', () => {
  const app = mountOutline()
  app.state.activeId.value = 'a'
  app.events.change()
  assert.equal(app.state.items.value.length, 1, 'change must not blank the visible outline')
  assert.equal(app.state.activeId.value, 'a')
  app.setHeadings([{ text: '新标题', level: 0, blockId: 'a' }])
  app.timers.at(-1)()
  assert.equal(app.state.items.value[0].text, '新标题')
  app.setHeadings([])
  app.events.change()
  app.timers.at(-1)()
  assert.equal(app.state.items.value.length, 0, 'deleted headings must not remain cached')
})

test('successive scroll events reuse the heading DOM index', () => {
  let scans=0
  const scroll={scrollTop:0,addEventListener(){},removeEventListener(){},
    getBoundingClientRect:()=>({top:0}),
    querySelectorAll:()=>{scans++;return [{dataset:{paragraphId:'a'},getBoundingClientRect:()=>({top:20-scroll.scrollTop})}]},
  }
  const app=mountOutline(scroll)
  for(let i=0;i<100;i++){scroll.scrollTop=i;app.state.updateActive()}
  assert.equal(scans,1,'scroll must not scan every document paragraph each frame')
  assert.equal(app.state.activeId.value,'a')
})

test('temporary loading retains outline and collapsed branches', () => {
  const app = mountOutline()
  app.state.collapsed.value.add('a')
  app.loading.value = true
  app.watchers.forEach(callback => callback())
  assert.equal(app.state.items.value.length, 1)
  assert.equal(app.state.collapsed.value.has('a'), true)
})
