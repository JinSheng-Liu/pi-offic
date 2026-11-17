import test from 'node:test'
import assert from 'node:assert/strict'
import {budgetDocumentContext,inspectModelPayload} from '../server/document/read-context.mjs'
const read=(id)=>({role:'toolResult',toolName:'inspect_document',toolCallId:id,details:{payload:'x'.repeat(200)},content:[{type:'text',text:JSON.stringify({revision:1,hash:'abc',view:'paragraphs',items:[{index:id,text:'内容'.repeat(1000)}]})}]})
test('context budget bounds old reads without changing source or tool identities',async()=>{
 const messages=[{role:'user',content:'要求'},read('old'),{role:'toolResult',toolName:'apply_formatting',content:[{type:'text',text:'已提交 V2'}]},read('latest')]
 const saved=structuredClone(messages),result=await budgetDocumentContext(messages,{budgetBytes:8000})
 assert.equal(JSON.parse(result[1].content[0].text).omittedForContextBudget,true)
 assert.deepEqual(result[3],{...messages[3],details:undefined});assert.deepEqual(result[2],messages[2]);assert.deepEqual(messages,saved)
 assert.equal(result[1].toolCallId,'old')
})
test('payload guard counts bytes and rejects excessive payloads before sending',()=>{
 assert.equal(inspectModelPayload({messages:[],tools:[]}).messageCount,0)
 assert.throws(()=>inspectModelPayload({messages:['中'.repeat(210000)]}),/上下文预算/)
})

test('large document projection preserves system instructions and tool declarations',async()=>{
 const system={role:'system',content:'Document safety rules',toolsAdded:[{name:'inspect_document',parameters:{type:'object'}}]}
 const messages=[system,{role:'user',content:'x'.repeat(310000)},read('latest')]
 const projected=await budgetDocumentContext(messages)
 assert.deepEqual(projected[0],system)
 assert.equal(projected[1].content,messages[1].content)
})
