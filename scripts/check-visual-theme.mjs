import assert from 'node:assert/strict'
import { writeFile } from 'node:fs/promises'
import { pathToFileURL } from 'node:url'
import { documentFixture, richParagraph } from '../test/helpers/docx-fixture.mjs'

const modulePath = process.env.HOME_TEST_PLAYWRIGHT
if (!modulePath) throw new Error('Set HOME_TEST_PLAYWRIGHT to Playwright index.mjs')
const { chromium } = await import(pathToFileURL(modulePath).href)
const browser = await chromium.launch({
  headless: true,
  ...(process.env.HOME_TEST_BROWSER ? { executablePath: process.env.HOME_TEST_BROWSER } : {}),
})
const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } })
const checks = []
const errors = []
page.on('pageerror', error => errors.push(error.message))

async function screenshot(name, options = {}) {
  await page.screenshot({ path: `output/style-${name}.png`, fullPage: true, ...options })
}

async function matchesToken(selector, property, token) {
  const result = await page.locator(selector).first().evaluate((element, args) => {
    const probe = document.createElement('span')
    probe.style[args.property] = `var(--studio-${args.token})`
    element.append(probe)
    const actual = getComputedStyle(element)[args.property]
    const expected = getComputedStyle(probe)[args.property]
    probe.remove()
    return { actual, expected }
  }, { property, token })
  assert.equal(result.actual, result.expected, `${selector} ${property}`)
}

async function noOverflow() {
  const bounds = await page.evaluate(() => ({
    viewport: document.documentElement.clientWidth,
    content: document.documentElement.scrollWidth,
  }))
  assert.ok(bounds.content <= bounds.viewport, JSON.stringify(bounds))
}

try {
  await page.goto(process.env.HOME_TEST_URL || 'http://127.0.0.1:5175')
  await page.locator('.studio-home').waitFor()
  assert.equal(await page.locator('html').getAttribute('data-theme'), 'light')
  await matchesToken('.home-header', 'backgroundColor', 'surface')
  await matchesToken('.home-upload-button', 'backgroundColor', 'accent')
  await screenshot('home-light')
  checks.push('Light home uses shared surface and primary tokens')

  await page.getByRole('button', { name: '模型设置', exact: true }).click()
  await page.locator('.workspace-dialog[open]').waitFor()
  await matchesToken('.workspace-dialog[open]', 'backgroundColor', 'surface')
  await screenshot('settings-light', { mask: [page.locator('.workspace-dialog input')] })
  await page.getByRole('button', { name: '关闭', exact: true }).click()

  await page.getByRole('button', { name: '切换深色界面' }).click()
  await matchesToken('.home-header', 'backgroundColor', 'surface')
  await screenshot('home-dark')
  await page.reload()
  await page.locator('.studio-home').waitFor()
  assert.equal(await page.locator('html').getAttribute('data-theme'), 'dark')
  checks.push('Homepage theme toggle persists after reload')
  await page.getByRole('button', { name: '切换浅色界面' }).click()

  const bytes = documentFixture(
    richParagraph('第一章 文档排版', '<w:pStyle w:val="Heading1"/>') +
    richParagraph('真实文档预览。中文、English 与数字 2026 保留原有内容。') +
    richParagraph('1.1 版式与结构', '<w:pStyle w:val="Heading2"/>') +
    richParagraph('这是一段带有原始红色的内容。', '', '<w:color w:val="B93838"/>'),
  )
  await page.locator('input[type="file"]').setInputFiles({
    name: '主题与界面验收.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    buffer: Buffer.from(bytes),
  })
  const editor = page.getByRole('textbox', { name: '文档内容', exact: true })
  await editor.waitFor({ timeout: 60000 })
  await page.locator('.outline-link').first().waitFor()
  await matchesToken('.topbar', 'backgroundColor', 'surface')
  await matchesToken('.assistant-panel', 'backgroundColor', 'surface')
  await matchesToken('.outline-file', 'color', 'muted')
  await matchesToken('.outline-link[aria-current]', 'backgroundColor', 'accent-soft')
  await matchesToken('.top-actions .open-document', 'backgroundColor', 'surface')
  await matchesToken('.outline-link[aria-current]', 'color', 'accent')
  await screenshot('workspace-light')
  checks.push('Workspace and teleported outline use shared tokens without native-editor collisions')

  const authoredRun = page.locator('.docx-page span').filter({ hasText: '这是一段带有原始红色的内容。' }).last()
  const authoredColor = await authoredRun.evaluate(element => getComputedStyle(element).color)
  await page.getByRole('button', { name: '切换深色界面' }).click()
  await page.locator('.editor-frame .docx-editor.dark').first().waitFor()
  await matchesToken('.assistant-panel', 'backgroundColor', 'surface')
  await matchesToken('.outline-file', 'color', 'muted')
  assert.equal(await authoredRun.evaluate(element => getComputedStyle(element).color), authoredColor)
  assert.equal(await page.locator('.docx-page-content').first().evaluate(element => getComputedStyle(element).filter), 'none')
  assert.equal(await page.locator('.docx-page').first().evaluate(element => getComputedStyle(element).backgroundColor), 'rgb(255, 255, 255)')
  await screenshot('workspace-dark')
  checks.push('Dark native chrome follows app theme while paper and authored text colours stay intact')

  await page.getByRole('tab', { name: /修改历史/ }).click()
  await page.locator('.history-panel').waitFor()
  await screenshot('history-dark')
  await page.getByRole('tab', { name: '对话助手', exact: true }).click()
  await page.getByRole('button', { name: '切换浅色界面' }).click()

  await editor.click()
  await page.keyboard.press('Control+End')
  await page.keyboard.type(' Edit check.')
  await page.getByRole('button', { name: '保存', exact: true }).click()
  await page.getByRole('button', { name: '保存', exact: true }).and(page.locator('[disabled]')).waitFor()
  await page.getByRole('tab', { name: /修改历史/ }).click()
  await screenshot('history-light')
  checks.push('Real DOCX remains editable and save action completes; history renders in both themes')

  await page.getByLabel('更多文档操作', { exact: true }).click()
  await page.getByRole('button', { name: '页面设置', exact: true }).click()
  await page.getByRole('dialog', { name: '页面设置' }).waitFor()
  await matchesToken('.workspace-dialog[open]', 'backgroundColor', 'surface')
  await screenshot('page-settings')
  await page.getByRole('button', { name: '关闭', exact: true }).click()

  await editor.click()
  await page.keyboard.press('Control+End')
  await page.keyboard.type(' Unsaved check.')
  await page.getByRole('link', { name: 'Pi DOCX Studio 首页' }).click()
  await page.getByRole('dialog', { name: '关闭前保存修改？' }).waitFor()
  await matchesToken('.close-document-dialog', 'backgroundColor', 'surface')
  await screenshot('close-document')
  await page.getByRole('button', { name: '取消关闭', exact: true }).click()
  await page.getByRole('button', { name: '保存', exact: true }).click()
  await page.getByRole('button', { name: '保存', exact: true }).and(page.locator('[disabled]')).waitFor()
  checks.push('Page settings and unsaved-close dialogs use the shared visual system')

  for (const width of [360, 768, 1024, 1920]) {
    await page.setViewportSize({ width, height: 1000 })
    await noOverflow()
  }
  checks.push('Workspace has no page overflow at 360, 768, 1024 and 1920px')

  await page.setViewportSize({ width: 390, height: 844 })
  await noOverflow()
  await screenshot('mobile-editor')
  await page.getByRole('button', { name: '大纲', exact: true }).click()
  await page.locator('.rail').waitFor({ state: 'visible' })
  await screenshot('mobile-outline')
  await page.getByRole('button', { name: '助手', exact: true }).click()
  await page.locator('.assistant-panel').waitFor({ state: 'visible' })
  await page.getByRole('tab', { name: '对话助手', exact: true }).click()
  await screenshot('mobile-chat')
  await noOverflow()
  checks.push('390px workspace switches editor, outline and assistant without page overflow')

  await page.getByRole('link', { name: 'Pi DOCX Studio 首页' }).click()
  await page.locator('.studio-home').waitFor()
  await noOverflow()
  await screenshot('mobile-home')
  for (const width of [360, 768, 1024, 1920]) {
    await page.setViewportSize({ width, height: 1000 })
    await noOverflow()
  }
  checks.push('Homepage has no page overflow at 360, 768, 1024 and 1920px')
  await page.setViewportSize({ width: 390, height: 844 })
  await page.getByRole('button', { name: '切换深色界面' }).click()
  await screenshot('mobile-home-dark')
  await page.getByRole('button', { name: '模型设置', exact: true }).click()
  await page.locator('.workspace-dialog[open]').waitFor()
  await matchesToken('.workspace-dialog[open]', 'backgroundColor', 'surface')
  await screenshot('mobile-settings-dark', { mask: [page.locator('.workspace-dialog input')] })
  checks.push('Mobile home and settings share light/dark palette')
  assert.deepEqual(errors, [])
  checks.push('No uncaught browser errors')
  await writeFile('output/style-browser-check.json', JSON.stringify({ status: 'passed', checks }, null, 2), 'utf8')
  console.log(JSON.stringify({ status: 'passed', checks }, null, 2))
} finally {
  await browser.close()
}
