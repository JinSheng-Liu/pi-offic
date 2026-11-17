import test from 'node:test'
import assert from 'node:assert/strict'
import {createInspectionReader} from '../server/document/inspection-cache.mjs'
test('inspection snapshot reuses identical bytes, changes on writes, and never caches failure',async()=>{
  let calls=0
  const read=createInspectionReader(async bytes=>{calls++;if(bytes[0]===0)throw Error('bad');return {value:bytes[0]}})
  const results=await Promise.all([read(new Uint8Array([1])),read(new Uint8Array([1]))])
  assert.deepEqual(results,[{value:1},{value:1}]);assert.equal(calls,1)
  assert.deepEqual(await read(new Uint8Array([2])),{value:2});assert.equal(calls,2)
  await assert.rejects(read(new Uint8Array([0])));await assert.rejects(read(new Uint8Array([0])));assert.equal(calls,4)
})
