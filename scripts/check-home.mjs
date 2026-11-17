import assert from 'node:assert/strict'
import { writeFile } from 'node:fs/promises'
import { pathToFileURL } from 'node:url'
import { paragraphDocument } from '../test/helpers/docx-fixture.mjs'

// Keep browser tooling outside the product dependencies.
const modulePath = process.env.HOME_TEST_PLAYWRIGHT
if (!modulePath) throw new Error('Set HOME_TEST_PLAYWRIGHT to the Playwright index.mjs path')
const { chromium } = await import(pathToFileURL(modulePath).href)
const browser = await chromium.launch({
  headless: true,
  ...(process.env.HOME_TEST_BROWSER ? { executablePath: process.env.HOME_TEST_BROWSER } : {}),
})
const baseUrl = process.env.HOME_TEST_URL || 'http://127.0.0.1:5175'
const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } })
const errors = []
const checks = []
page.on('pageerror', error => errors.push(error.message))

try {
  await page.goto(baseUrl)
  await page.getByRole('heading', { name: /让 AI 帮你完成/ }).waitFor()
  assert.equal(await page.locator('.workspace').count(), 0)
  checks.push('No editor workspace before upload')
  await page.screenshot({ path: 'output/home-desktop.png', fullPage: true })

  const input = page.locator('input[type="file"]')
  await input.setInputFiles({ name: 'invalid.txt', mimeType: 'text/plain', buffer: Buffer.from('invalid') })
  await page.getByRole('alert').filter({ hasText: '请选择非空且不超过 20 MB' }).waitFor()
  assert.equal(await page.locator('.workspace').count(), 0)
  checks.push('Invalid file stays on home with an actionable error')

  await input.setInputFiles({ name: 'broken.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.from('not a document') })
  await page.locator('.home-error').filter({ hasNotText: '请选择非空且不超过 20 MB' }).waitFor()
  await page.locator('.home-upload-button:not([disabled])').waitFor()
  assert.equal(await page.locator('.workspace').count(), 0)
  checks.push('Server rejection leaves home ready to retry')

  const bytes = paragraphDocument(['首页上传验收', '上传后进入真实文档工作台。'])
  let releaseUpload
  let observeUpload
  const uploadGate = new Promise(resolve => { releaseUpload = resolve })
  const uploadObserved = new Promise(resolve => { observeUpload = resolve })
  await page.route('**/api/documents?name=*', async route => {
    observeUpload()
    await uploadGate
    await route.continue()
  }, { times: 1 })
  await input.setInputFiles({ name: '首页交互验收.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.from(bytes) })
  await uploadObserved
  await page.getByRole('heading', { name: '正在打开你的文档…' }).waitFor()
  assert.equal(await page.locator('.workspace').count(), 0)
  releaseUpload()
  checks.push('Upload stays on home with loading state until the document is ready')
  await page.locator('.workspace').waitFor({ timeout: 60000 })
  await page.getByRole('textbox', { name: '文档内容', exact: true }).waitFor({ timeout: 60000 })
  assert.equal(await page.locator('.studio-home').count(), 0)
  checks.push('Valid upload enters the real editor workspace')
  await page.screenshot({ path: 'output/home-workspace.png', fullPage: true })

  await page.getByRole('link', { name: 'Pi DOCX Studio 首页' }).click()
  await page.getByRole('heading', { name: /让 AI 帮你完成/ }).waitFor()
  checks.push('Workspace home link closes document and returns home')
  await page.getByRole('button', { name: /首页交互验收.docx/ }).first().click()
  await page.locator('.workspace').waitFor()
  checks.push('Recent document resumes the workspace')
  await page.getByRole('link', { name: 'Pi DOCX Studio 首页' }).click()

  const chooserEvent = page.waitForEvent('filechooser')
  await page.getByRole('button', { name: '上传 DOCX', exact: true }).click()
  const chooser = await chooserEvent
  await chooser.setFiles({ name: 'empty.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.alloc(0) })
  await page.getByRole('alert').filter({ hasText: '请选择非空且不超过 20 MB' }).waitFor()
  checks.push('Upload button opens picker and empty files are rejected')
  await input.setInputFiles({ name: 'large.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.alloc(20 * 1024 * 1024 + 1) })
  await page.getByRole('alert').filter({ hasText: '请选择非空且不超过 20 MB' }).waitFor()
  checks.push('Files over 20 MB are rejected before upload')

  await page.evaluate(data => {
    const transfer = new DataTransfer()
    transfer.items.add(new File([new Uint8Array(data)], '拖拽上传验收.docx', { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }))
    document.querySelector('.home-upload').dispatchEvent(new DragEvent('drop', { bubbles: true, cancelable: true, dataTransfer: transfer }))
  }, Array.from(bytes))
  await page.locator('.workspace').waitFor({ timeout: 60000 })
  await page.getByRole('textbox', { name: '文档内容', exact: true }).waitFor({ timeout: 60000 })
  checks.push('Dropped DOCX enters the real workspace')
  await page.getByRole('link', { name: 'Pi DOCX Studio 首页' }).click()

  await page.setViewportSize({ width: 390, height: 844 })
  await page.screenshot({ path: 'output/home-mobile.png', fullPage: true })
  const size = await page.evaluate(() => ({ width: document.documentElement.clientWidth, scroll: document.documentElement.scrollWidth }))
  assert.ok(size.scroll <= size.width, 'Mobile home must not overflow horizontally')
  checks.push('390px home has no horizontal overflow')
  assert.deepEqual(errors, [])
  checks.push('No browser page errors')
  await writeFile('output/home-browser-check.json', JSON.stringify({ status: 'passed', checks }, null, 2), 'utf8')
  console.log(JSON.stringify({ status: 'passed', checks }, null, 2))
} finally {
  await browser.close()
}
