import assert from 'node:assert/strict'
import { fileURLToPath } from 'node:url'
import { Agent } from '@earendil-works/pi-agent-core'
import { Type, createAssistantMessageEventStream } from '@earendil-works/pi-ai'
import { SessionManager } from '../vendor/pi/packages/coding-agent/src/core/session-manager.ts'

const resolved = import.meta.resolve('@earendil-works/pi-agent-core')
assert.equal(resolved, new URL('../vendor/pi/packages/agent/src/index.ts', import.meta.url).href)
assert.equal(import.meta.resolve('@earendil-works/pi-ai'),
  new URL('../vendor/pi/packages/ai/src/index.ts', import.meta.url).href)
const deadline = setTimeout(() => {
  console.error('Pi source verification timed out after 15 seconds')
  process.exit(1)
}, 15_000)
let requests = 0
let inspections = 0
const events = []
const session = SessionManager.inMemory(process.cwd())
const agent = new Agent({
  initialState: {
    systemPrompt: '这是离线源码验证。仅允许读取测试文档摘要。',
    tools: [{
      name: 'inspect_document', label: 'Inspect fixture',
      description: 'Read the in-memory test fixture.',
      parameters: Type.Object({}),
      execute: async () => {
        inspections++
        return { content: [{ type: 'text', text: '测试文档：revision=1' }], details: { revision: 1 } }
      },
    }],
  },
  streamFn: (_model, context) => {
    requests++
    assert.ok(requests <= 2, 'Unexpected extra model turn')
    const toolUse = requests === 1
    if (!toolUse) assert.ok(context.messages.some(m => m.role === 'toolResult'))
    const message = {
      role: 'assistant',
      content: toolUse
        ? [{ type: 'toolCall', id: 'fixture-call', name: 'inspect_document', arguments: {} }]
        : [{ type: 'text', text: '离线源码工具调用验证完成' }],
      api: 'openai-responses', provider: 'offline-test', model: 'fixture',
      usage: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, totalTokens: 0,
        cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 } },
      stopReason: toolUse ? 'toolUse' : 'stop', timestamp: Date.now(),
    }
    const stream = createAssistantMessageEventStream()
    queueMicrotask(() => stream.push({ type: 'done', reason: message.stopReason, message }))
    return stream
  },
})
agent.subscribe(event => {
  events.push(event.type)
  if (event.type === 'message_end') session.appendMessage(event.message)
})
await agent.prompt('读取测试文档摘要')
assert.equal(requests, 2)
assert.equal(inspections, 1)
assert.ok(events.includes('tool_execution_end'))
assert.ok(events.includes('agent_end'))
assert.ok(session.getEntries().length >= 3)
assert.equal(agent.state.isStreaming, false)
assert.equal(agent.state.messages.at(-1).content[0].text, '离线源码工具调用验证完成')
clearTimeout(deadline)
console.log(JSON.stringify({ status: 'passed', source: fileURLToPath(resolved),
  requests, inspections, sessionEntries: session.getEntries().length, network: 'not used' }, null, 2))
