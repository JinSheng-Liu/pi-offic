async page => {
  const origin = new URL(page.url()).origin
  const health = await (await page.request.get(`${origin}/api/health`)).json()
  if (!health.testMode) throw new Error('浏览器测试必须使用 npm run dev:test 启动的隔离工作区')
  await page.unroute('**/api/documents/*/chat')
  await page.route('**/api/documents/*/chat', route => {
    if (route.request().method() !== 'POST') return route.continue()
    const frames = [
      ['status', { state: 'running' }],
      ['tool', { id: 'ui-test', name: 'inspect_document', label: '读取文档', state: 'running' }],
      ['tool', { id: 'ui-test', name: 'inspect_document', label: '读取文档', state: 'complete' }],
      ['text', { delta: '## 排版检查\n\n**已完成**\n\n- 第一项\n- 第二项\n\n<script>alert(1)</script>' }],
      ['done', { state: 'complete' }],
    ]
    return route.fulfill({ contentType: 'text/event-stream', body: frames.map(([event, data]) => `event: ${event}\ndata: ${JSON.stringify(data)}\n\n`).join('') })
  })
  try {
    await page.getByRole('textbox', { name: '输入文档处理需求' }).fill('前端展示验证，不修改文档')
    await page.getByRole('button', { name: '发送', exact: true }).click()
    await page.locator('.markdown-body strong').filter({ hasText: '已完成' }).waitFor()
    if (await page.locator('.markdown-body script').count()) throw new Error('HTML was not escaped')
    if (await page.locator('.markdown-body li').count() !== 2) throw new Error('Markdown list missing')
    await page.locator('.execution-steps').getByText('已完成', { exact: true }).waitFor()
    return { markdown: true, htmlEscaped: true, toolSteps: true }
  } finally { await page.unroute('**/api/documents/*/chat') }
}
