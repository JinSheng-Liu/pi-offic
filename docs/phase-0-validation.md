# Phase 0 技术验证记录

日期：2026-11-01。此报告只记录已执行的检查，不代表 V1.0 验收。

## Pi 本地源码

- 上游： https://github.com/earendil-works/pi
- 标签：v1.4.0；提交：4a080252f06caf63a525d7873c5d94d3c23887e0。
- 路径：vendor/pi；MIT 许可证原文保留于 vendor/pi/LICENSE。
- 当前为本地源码基线，未发布远程 Fork，未修改上游核心。
- Node.js：24.19.0；npm：12.0.2。
- 安装命令：`npm --prefix vendor/pi ci --ignore-scripts`，成功。
- 验证命令：`npm run check:pi`，成功。

验证通过上游 source-resolver 加载本地 TypeScript。断言 Agent 与 pi-ai 的解析路径指向 vendor/pi/packages 下的 src 文件。
Agent Loop 执行两轮模拟响应、一次 inspect_document 测试工具调用，SessionManager 记录四条内存会话条目。
这证明最小进程内源码运行链可用；产品随后已接入 Pi Agent、项目技能、文档工具和基于文档的会话文件。真实模型、真实 API 流式响应与取消仍未在有凭据的环境中端到端验证。
测试工具读取内存固定数据，并非产品文档服务；模拟响应不进入产品界面。

## 依赖审计

`npm audit --json`：15 项受影响包，9 中等、6 高危。
`npm audit --omit=dev --json`：5 项受影响包，1 中等、4 高危。

| 包 | 生产依赖审计等级 |
| --- | --- |
| @anthropic-ai/sandbox-runtime | high |
| minimatch | moderate |
| node-forge | high |
| undici | high |
| zod | high |

审计范围是整个 Pi monorepo，包括示例和其他工作区，不是最终产品依赖闭包。
不能将这些告警一概视为仅开发依赖，也不能据此认定产品已存在相同可利用路径。
尚未升级依赖或变更上游锁文件。部分修复涉及主版本或上游声明为无直接修复；需针对实际运行路径验证后制定升级方案。
此状态不满足商用发布的依赖验收条件。

## DOCX 编辑器

此前已完成 Vue 原型的类型检查、生产打包以及简单中英文文本的编辑、导出、ZIP 主文档内容检查、重新打开显示。当前应用已接入本地文档 API、不可变版本保存/恢复及 Pi 助手界面；这些产品路径的最新集成验证以交付时实际运行的检查为准。
使用 @docx-editor.dev/vue 与 core 2.33.0（包元数据为 Apache-2.0），未使用 Pro editor-api。
未验证复杂 DOCX 保真、Word 重新打开、中文字体兼容、大文档性能、服务端版本保存及自动恢复。

## TypeScript docx4j 来源与基础验证

用户已确认来源为 https://github.com/plutext/docx4j-core-ts 。准确包名是 `@docx4j/core-ts`，已安装固定发布版 0.6.0，许可证 Apache-2.0。
此前调查的 docx4js 不采用。GitHub 主分支与 npm 发布版的依赖不同，实际验证以发布包为准。

执行 `npm run check:docx` 通过：创建两段中英文文档，保存、重新加载、修改首段为加粗/18 pt/居中、保存并再次加载。
断言正文及段落数保持一致，首段格式正确，第二段原有 11 pt 与斜体保持，原文件字节未变。
测试文件在 data/validation/，重复执行会重建这两个专用测试文件。
本次安装后根项目 npm 审计为 0 漏洞；独立 Pi 依赖审计结果不变。

这仅验证合成简单文档，不证明真实复杂文档、未识别 XML、图片、公式、分节、修订或 Word 排版兼容。下一步以真实文档和跨引擎往返测试确定写入边界。
