# 文档排版技能体系

本目录是产品 Pi Runtime 使用的技能，不是额外的文档处理引擎。规则执行仍通过DOCX Editor 公开 Store / Layout 适配器和小范围 OOXML 补充完成。

## 编写与维护标准

每份 `SKILL.md` 保持 70–90 行，按文件实际行数统计，包含元信息、标题和正常段间空行；禁止用重复文字或多余空行凑数。当前 10 份各为 72 行。

统一包含适用范围、输入与前置检查、识别与决策、执行流程、验收要求、案例与失败边界六部分。类型技能说明特有决策，通用字体、版本和错误处理在公共技能维护。

修改后运行 `node --test test/skills.test.mjs` 检查长度、章节、UTF-8 与技能引用，再运行 `npm run check:api` 验证 Pi 加载兼容性。静态检查不能证明模型在全部文档上遵守规则；行为质量仍需真实任务与文档回读验证。

当前加载工具只返回 `SKILL.md`，尚无配套参考文件读取入口，因此必要流程保持自包含。未来确需拆出参考资料时，应同时实现并测试按需读取，不能只创建模型无法访问的文件。

## 加载方式

- `docx-formatting`、`docx-validation` 为常驻公共规则。
- 其余技能仅将名称和说明加入发现目录，通过 `load_document_skill` 按需读取正文。
- `inspect_document` 提供类型线索与对象技能建议。线索来自前 160 段文字，不是确定分类，不直接修改文档。混合文档可加载多个技能。
- 用户当前任务决定是否需要对象技能；“论文”不意味着自动要求封面、目录、参考文献全部重做。
- 目录由 Pi 现有技能加载器发现，工具只接受注册名称，不接受文件路径。

## 技能分工

| 层级 | 技能 | 重点 |
|---|---|---|
| 公共 | docx-formatting / docx-validation | 字体槽、范围、版本、覆盖清单、诚实验收 |
| 类型 | docx-academic | 论文分区、学校模板优先、封面保护 |
| 类型 | docx-report | 报告阅读层级、摘要、数据表达 |
| 类型 | docx-official | 公文要素分区，不冒称国标合规 |
| 类型 | docx-contract | 条款、编号、签署留空保护 |
| 类型 | docx-technical | 步骤、代码、参数表、警示 |
| 对象 | docx-objects | 图片、题注、单元格定位与对齐 |
| 对象 | docx-toc-sections | 目录域、分节、页眉页脚边界 |
| 参考 | docx-reference | 样稿按区域匹配、参考可读性约束 |

## 工具与能力边界

`inspect_document` 默认返回 40 项，最多 80 项。返回 `pagination.nextOffset` 时继续读取。`summary` 只给统计，不代表读过全文。`indices` 可回读已修改段落。

`cells` 视图的 `paragraphIndices` 可用于现有格式工具；单元格编号不能当段落编号。图片段落提供 `inlinePictureCount`，可用其段落索引设置水平对齐。浮动图位置不等于段落对齐。题注工具使用单独的顶层索引空间。

`body` 排除已知表格、图片、标题、目录样式、题注、列表、空段，但无法可靠识别全部 Normal 样式封面和签署区。复杂文档须检查分区并使用明确索引。

当前技能不会赋予工具原本没有的能力：目录页码重算、制表位写入、浮动图位置、表格垂直对齐、页眉页脚内容修改尚未在本轮新增。需求覆盖清单目前是技能约束，不是独立的机器强制验收系统。页面留白、分页和字体实际渲染需要另行验收。

## 参考来源与适用范围

核对时间：2026-11-01。

- [Microsoft：自定义或创建样式](https://support.microsoft.com/en-us/office/customize-or-create-new-styles-d38d6e47-f6fc-48eb-a607-1eb120dec563)：支持使用样式统一标题、段落、列表的字体、间距、缩进等；不提供适用于全部文档的固定数值模板。
- [Google：技术文档标题](https://developers.google.com/style/headings)：参考层级与可导航性；不自动重写中文标题、不照搬英文大小写和编号约定。
- [Google：技术文档表格](https://developers.google.com/style/tables)：参考表头语义和题注可读性；HTML 标记、响应式和“不合并单元格”等网页规则不用于重构现有 DOCX。
- [Pandoc 手册源码](https://raw.githubusercontent.com/jgm/pandoc/main/MANUAL.txt)：参考 reference-doc 的样式模板思路；没有引入 Pandoc 做复杂 DOCX 往返转换。

APA 页面未获取到可用规则正文，公文候选链接返回 404，不把它们记为已核实规范。合同、报告与论文具体数值须来自用户指定的模板或可核实标准，不能伪造“通用标准”。检索证据保存在 `data/validation/skills-redesign/`。

结构读取默认仅含精简文本和定位信息；`indices` 定点读取默认含完整格式，也可显式指定 `detail=formatting`。`lineSpacing` 返回继承后规则及倍数/磅值，`tabStops` 返回直接和继承制表位；这些数值不等于页面视觉验收。完全重复的成功读取在发送模型前去重，原始会话记录保留。

受限的 `leadingTabPosition` 仅用于明确段落索引、单个前导 Tab、无已有自定义制表位/缩进/编号的普通非表格段落，单位 pt。它添加段落左制表位并保留原文，不修改 `firstLineIndent`，不代表复杂制表位或目录引导符编辑已开放。
