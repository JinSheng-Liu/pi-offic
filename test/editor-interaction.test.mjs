import test from 'node:test'
import assert from 'node:assert/strict'
import {blockBusyEditorInput} from '../src/lib/editor-interaction.ts'
function event(type,props={}) {
 const e=new Event(type,{cancelable:true})
 for(const [key,value] of Object.entries(props))Object.defineProperty(e,key,{value})
 return e
}
test('busy editor preserves scrollbar pointer input and reading keys',()=>{
 for(const e of [event('pointerdown',{target:{closest:()=>null}}),event('keydown',{key:'PageDown'}),event('keydown',{key:'ArrowDown'}),event('keydown',{key:'Tab'})]) {
  blockBusyEditorInput(e);assert.equal(e.defaultPrevented,false)
 }
})
test('busy editor blocks document writes and toolbar activation',()=>{
 for(const e of [event('beforeinput'),event('paste'),event('cut'),event('drop'),event('compositionstart'),event('keydown',{key:'a'}),event('keydown',{key:'z',ctrlKey:true}),event('click',{target:{closest:()=>({})}}),event('pointerdown',{target:{closest:()=>({})}})]) {
  blockBusyEditorInput(e);assert.equal(e.defaultPrevented,true)
 }
})


test('busy toolbar events cannot reach subsequent command handlers', () => {
  const target = new EventTarget()
  let executed = false
  target.addEventListener('click', blockBusyEditorInput)
  target.addEventListener('click', () => { executed = true })
  const click = event('click', { target: { closest: () => ({}) } })
  target.dispatchEvent(click)
  assert.equal(executed, false)
})
