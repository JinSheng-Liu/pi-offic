# 编辑器引擎迁移记录

日期：2026-11-11。

## 当前生产架构

- DocumentWorkflow → editor-adapter.mjs：业务范围、原子提交与保护验证。
- editor-reader.mjs → DOCX Editor Store / Layout：包解析、段落遍历、样式级联、运行字体读取；转换为业务快照。
- editor-format.mjs → DOCX Editor TreePackageStore：字体、局部文字格式、对齐、缩进、固定／倍数行距、段间距与分页标记。
- 窄范围 OOXML 补充：标题大纲、特殊制表位，以及现有封面、题注、分节、目录和表格专项处理。
- editor-automation.mjs：新建 A4 文档与高层 Automation 入口。

生产、测试和验证脚本均已移除 @docx4j/core-ts 导入，package.json/package-lock.json 中已卸载该依赖。旧 docx4j-adapter.mjs 已由 editor-adapter.mjs 替换。测试夹具使用编辑器空白包与明确的 OOXML 数据，未再维护另一套 SDK。

## 已验证

- 服务端、API、前端逻辑、正式文档引擎检查、Pi 集成与 TypeScript/Vite 构建均已运行通过（新增特殊字符测试后的最终计数以检查日志为准）。
- 真实论文 1173 段经过新生产链路整体字体和行距修改，正文结构与非目标部分保护检查通过。
- 原文件 SHA-256：fa61bdd12a2547f0541bf5ff00c74ac2a342fd7befc27d2e0063591c311be17a。
- 编辑器与旧 SDK 读取同一论文的段落数、表格归属一致；唯一文本表示差异为 Wingdings 2 的 w:sym，现快照保留原始编码，旧 SDK 映射为括号。文件中的符号节点保持不变。
- 固定行距明确写 exact，多倍行距明确写 auto；不再把 Automation.lineSpacing 参数直接当作倍率。
- 页眉等 XML 的 mc:Ignorable 顺序变化按命名空间集合比较；实际命名空间变化仍被拒绝。
- 空白文档尺寸验证为 A4：11906 × 16838 twips。
- 批次失败与过期修订不能覆盖已提交内容。

## 已提交格式的增量预览

- 服务端基础格式批次验证、提交后，通过 SSE 发送带起止指纹和前后修订号的回放数据。
- 浏览器共享同一格式编译器，通过编辑器公开 applyAutomationOps 每次应用 20 段，并等待实际绘制帧。保留每个提交事件，不再合并跳过中间版本。
- 这属于“已提交批次的原位分段呈现”，不是后端计算尚未完成时就将未验证结果写入编辑器。
- 标题样式、分节、目录、封面、题注等专项操作仍同步已提交 DOCX；不宣称全部操作均支持原位增量。
- 指纹只忽略编辑器运行时段落身份属性，文字、格式变化仍校验。过期回放和只读事务被拒绝，失败后恢复已保存版本。
- AI 运行期间阻止人工修改；预览修订落后时禁止保存、导出与继续 AI 请求，避免交付部分回放结果。
- 原生目录检测标记生成目录段落，正文范围排除这些段落，不绕过编辑器的只读目录保护。

## 浏览器验证

- 合成 45 段：3 批，无重新挂载，只读与过期回放拒绝。
- 真实论文正文 213 段：11 批，约 2.45 秒，无重新挂载，最终指纹相符；可见段落高度随格式变化。
- 首批 20 段后故障注入：确实发生部分修改；官方刷新恢复后指纹与已提交结果一致。
- 最终 npm run check 已通过：服务端、API、前端逻辑、引擎检查、Pi 集成、类型检查与生产构建。

## 尚未验收的边界

1. 未完成本轮真实模型联网端到端测试，也未验收 Word/WPS 页面视觉质量。
2. 表格条件样式、主题字体以及复杂兼容分支的完整读取语义仍需扩展专项验证，不能以合成测试覆盖全部 DOCX。
3. 专项结构操作尚非逐段原位预览；完整产品工作区的浏览器端到端验收不能由独立预览函数测试替代。
4. 原有大包构建提示尚存；npm 安装检查报告一项低风险依赖问题，未自动升级。

## 验证产物

- output/editor-migration-full-check.log：完整检查日志。
- output/editor-production-regression/编辑器引擎格式回归.docx：整篇基础格式回归产物，不是最终论文排版交付。
- output/editor-production-regression/report.json：真实文档保护验证。
- test/editor-reader.test.mjs：特殊符号、缩进和间距独立回归，无 core-ts 依赖。

- output/final-engine-migration-check.log：卸载依赖并完成预览保护后的完整检查日志。
- output/live-format-probe/real-body-report.json：真实正文分批预览证据。
- output/live-format-probe/partial-recovery-report.json：部分修改后的恢复证据。
- test/format-replay.test.mjs：回放一致性、提交事件和指纹保护测试。
