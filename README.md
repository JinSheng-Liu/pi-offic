# Pi DOCX Studio

本地运行的 DOCX 编辑与受控格式处理工作台。当前提供 Vue 编辑器、Node 本地文档服务及 Pi 源码 Agent 运行时；产品处于 V0.2 阶段，尚未完成 V1.0 验收。

## 启动

需要 Node.js 24.19 或更新兼容版本。源码使用 UTF-8。

```powershell
npm ci --ignore-scripts
npm --prefix vendor/pi ci --ignore-scripts
npm run prepare:pi
npm run dev
```

默认服务地址为 http://127.0.0.1:5173。生产模式使用 `npm run build` 和 `npm start`。文档和会话数据写入 `data/documents/`，可通过 `PI_DOCX_DATA_DIR` 指定其他本地目录。

`prepare:pi` 调用 Pi 官方模型数据生成脚本，首次安装或更新 Pi 源码后联网执行；日常启动无需重复执行。发布包需包含生成的 `vendor/pi/packages/ai/src/providers/data/` 数据。

会话生命周期、持久化、自动上下文压缩与取消复用 Pi `AgentSession`。项目仅精简文档读取结果，不自行截断系统消息或重建通用压缩摘要。`npm run check:session` 验证自动/手动压缩及恢复后工具仍可调用。

## 本地模型配置

复制 `.env.example` 为项目根目录的 `.env`，填写 `OPENAI_API_KEY`。
`npm run dev` 和 `npm start` 均会在启动时读取 `.env`；文件不存在时仍可启动。
系统环境变量优先于 `.env`。修改文件后需要重启服务。
服务重启后会重新加载文件，因此不需要每次在页面填写密钥。
页面“模型设置”仅临时覆盖当前进程，不会写回 `.env`。
`.env` 已被 Git 忽略；密钥不要使用 `VITE_` 前缀，以免暴露给前端。

```powershell
Copy-Item .env.example .env
# 编辑 .env 后启动
npm run dev -- --port 5174
```

## 当前功能

- 新建空白 DOCX、打开本地 DOCX、编辑正文、保存到本机文档空间、导出副本。
- 文档采用不可变版本存储；保存检查 revision，防止旧版本覆盖；支持恢复为新的版本。
- Pi 通过本地上游源码运行，提供读取结构、格式分析、受限格式应用及验证四个工具。
- 在“模型设置”中配置 OpenAI Chat Completions 兼容服务。也可通过 `OPENAI_API_KEY`、`PI_API_BASE_URL`、`PI_MODEL` 设置启动默认值。页面输入的密钥只保存在服务进程内存，文件配置则在每次启动时加载；浏览器聊天会向所配置的模型服务发送指令和文档结构摘要。
- 格式写入经过白名单校验、结构保护比较和保存重读；未验证的复杂文档边界会在验证信息中注明。

## 验证

```powershell
npm run check:server
npm run check:pi
npm run check:docx
npm run build
```

`check:docx` 会重建 `data/validation/` 中的专用合成测试文件。`check:pi` 使用模拟模型验证 Agent Loop 和本地源码解析，不会调用真实模型。完整执行证据与风险见 [docs/phase-0-validation.md](docs/phase-0-validation.md)。

## Pi 源码与依赖

Pi 源码位于 `vendor/pi`，上游 https://github.com/earendil-works/pi ，MIT 许可证；基线为 v1.4.0，提交 `4a080252f06caf63a525d7873c5d94d3c23887e0`。运行时通过上游 source-resolver 加载本地 TypeScript 源码，并非另装 Pi SDK。

当前文档处理使用 DOCX Editor 适配层及必要的 OOXML 补充操作，不再依赖 `@docx4j/core-ts`。编辑器使用 Apache-2.0 的 `@docx-editor.dev/vue` 和 `core` 2.33.0。

`vendor/pi` 以源码快照纳入仓库，克隆后不需要初始化子模块；其上游许可证和源码一并保留。首次安装仍需执行 `npm run prepare:pi` 生成模型数据。

## V1.0 尚待验证

尚未完成真实模型端到端调用与取消验证、真实复杂 DOCX 集合的往返保真、Microsoft Word 排版与兼容测试、大文档性能、分节/页面设置界面，以及最终产品依赖闭包的安全审计。Pi monorepo 独立审计曾报告告警；详见验证记录，不能作为商用依赖验收结论。

## 仓库提交边界

提交应用源码、测试、脚本、文档与技能、公共静态资源、锁文件和第三方许可证。
不提交 `.env`、依赖安装目录、构建产物、`data/` 中的用户文档与会话、`output/` 和 `tmp/` 中的测试结果、浏览器调试记录或日志。这些文件仍保留在本机。
`.env.example` 仅提供无密钥的通用配置示例；实际服务地址与模型请按自己的服务填写。
