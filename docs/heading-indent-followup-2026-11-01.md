# 论文标题缩进与字体检查补充修复

## 根因与修改

同级标题中有12段包含前导Tab，同时又设置21或24 pt左缩进；首行实际多缩进。依据原要求保留标题两个汉字的左缩进，在这份论文中将这12段的firstLineIndent设为-leftIndent，保留所有原文、Tab及左缩进。该方案不是任意文档通用规则，自定义制表位必须单独验证。

inspect此前仅输出eastAsiaDirect，未提供ascii/hAnsi槽。现新增asciiDirect、hAnsiDirect，模型提示要求分别检查中西文字体。单测先复现缺失字段失败，补齐后通过。

## 验证

- 26项服务端测试、1项API测试、类型检查和生产构建通过。构建仍有约4.69 MB大包提示。
- 99段全文与工作区v2一致，7个媒体文件字节一致；副本全部文本run的ascii/hAnsi为Times New Roman。
- LibreOffice渲染8页；12个目标标题逐个匹配PDF文本并测量x起点，全部等于正文左边界70.95 pt加各自21/24 pt，误差小于0.15 pt。
- 修复副本已作为原文档3501e5c4-ed8d-4857-96e4-294b79ce0146的v3保存，历史v1/v2保留。
- 重放脚本：scripts/check-thesis-heading-indent.mjs。
- 测量报告：data/validation/heading-indent/report.json。

## 尚未验证或改变

没有重启用户5173服务；后端新增检查字段与提示需重启生效。没有进行Microsoft Word和编辑器内的最终视觉验收。截图裁切无法证明画布截断故障。保留两端对齐以及原文空格（例如A T），未擅自改写正文或宣称字距问题全部解决。
