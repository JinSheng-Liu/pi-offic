# 论文文字重叠：固定零行距缺陷

## 已证实根因

用户截图出现严重重叠。此前提供的 data/validation/thesis-size/最新论文任务_验收副本.docx 有 99 个 w:spacing，均为 line=0、lineRule=exact。原始用户要求明确为单倍行距，而模型把它表达为 lineSpacing:0。core-ts 的该 setter 始终按固定磅数写入，旧后端允许 0 且只验证值相等，错误通过了文件级验收。此前副本不能作为交付。

## 修复

- lineSpacing 仅为大于 0 的固定磅数。
- 新增 lineSpacingMultiple：单倍1、1.5倍1.5、双倍2，对应 OOXML auto 规则和240倍数单位。
- 两种模式互斥，后写操作覆盖前写的预期，保存后同时核对规则与数值。
- 模型工具 schema 明确语义并拒绝0；加入 explicit-line-spacing 能力标识提示旧服务需要重启。
- 回归先证明旧实现接受0，再确认修复后拒绝0、互斥参数，并验证1/1.5/2倍的原始XML。

## 实际验证

最新记录 754cb3bd-64ca-428c-b6a1-84aefd70d825 的21项操作，依据原始单倍要求将其中6项 lineSpacing:0 替换为 lineSpacingMultiple:1，生成独立副本。99段均写入auto/240，全部正文字符一致，7个媒体文件字节一致。原文v1没有改动。

LibreOffice 使用独立测试配置将修正版转换为8页PDF。逐页总览及首页图像检查未再出现成片文字重叠，流程图/电路图在总览中可见。此检查不等于Microsoft Word保真或全部论文规范验收。

- DOCX：data/validation/line-spacing/论文_单倍行距修正版.docx
- PDF：data/validation/line-spacing/论文_单倍行距修正版.pdf
- 总览：data/validation/line-spacing/contact.png
- 服务端23项测试及接口集成测试通过。

历史重放脚本中若仍有lineSpacing:0，应明确根据原始用户要求迁移参数，不能在生产后端偷偷把所有0解释为单倍。没有停止用户5173进程。
