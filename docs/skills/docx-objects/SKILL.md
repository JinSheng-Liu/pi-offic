---
name: docx-objects
description: 图表对象：图片段落、图注表注和表格单元格的独立处理与限制。
---

# 图表对象

## 适用范围

- 用于图片段落、图注表注和表格内容的格式处理。
- 对象类型由实际结构判断，不由“图”或“表”一字决定。
- 保持图片媒体、关系、尺寸、裁剪和原有内容。
- 单元格水平对齐与垂直对齐是不同能力。
- 浮动图位置与所在段落对齐是不同能力。
- 单独的段落字体规则仍使用公共技能。
- 用户未要求新增题注时，不擅自添加文字或编号。
- 图片尺寸与任意浮动定位尚未接入AI工具；列宽可通过fitTableWidth调整，修复重叠可检查对象进入文档流能力。

## 输入与前置检查

- 读取 paragraphs 的嵌入图片及 floatingObjects；objectIndex 不能传给段落 indices，转换需 flowCapability.supported=true。
- 需要题注时读取 figures、captionTables 和 captions 视图。
- 题注目标 index 属于专用顶层索引空间，不是段落索引。
- 读取 cells 获取 paragraphIndices；读取 tables 获取 tableIndex、rowDetails 和 cellDetails 物理索引。
- 区分真实表格表头、布局表格与无边框签署区。
- 嵌套单元格结合 nestingLevel 识别，不仅靠行列号。
- 检查已有题注及编号，避免重复插入。
- 无法准确定位对象时说明缺口，不凭相邻文本猜索引。

## 识别与决策

- 嵌入图片可通过实际所在段落设置水平对齐。
- 一个段落包含文字和图片时，对齐会作用整段，先确认范围。
- 对象转文档流须单独提交 scope=floatingObjects、objectIndices、objectLayout；inlinePicture 保留尺寸，captionInFlow 只处理明确关联题注。
- 转换后必须重新读取段落及对象索引；新图片/题注段落使用自动行距，图片段落不套固定小行距。
- tables.layoutRecommendation为database-schema时，用户要求整理此类表格可执行其operations；短内容居中、长说明左对齐，用户明确格式优先，不推广到其他表。
- tables.layoutRecommendation为analysis-matrix或indicator-matrix且用户要求整理表格时，采用其operations；widthOperation须独立提交并回读列宽。明确行距优先于默认1.15倍，不把此策略用于其他三列表。
- 合并单元格的物理 cellIndex 不一定等于视觉列号。
- 不擅自拆分合并单元格、增加边框或重排表格内容。
- 表内全局 keepNext 可能导致跨页异常，不作为统一默认。

## 执行流程

- 用户要求整理全文/排版，连续等列Tab数据且明确表头、内容可无损保留时，可直接convert_tabular_text；用户禁止结构变化或结构含糊才询问。转换工具自动排版已识别的数据表，无须重复执行同一建议。
- 为明确图片段落使用 scope=indices 设置 alignment。
- 为指定单元格使用其 paragraphIndices 设置字体和段落属性。
- 垂直对齐用 tableCells、tableIndex、cellTargets 的物理 rowIndex/cellIndex，table.verticalAlignment 为 Top/Center/Bottom。
- 新增图题放图后，表题放表前，调用 add_captions。
- targetIndex 只能取对应 figures 或 captionTables 的 index。
- 章节编号仅在前置一级标题明确标出第几章时使用。
- 工具拒绝重复、浮动或不安全编号插入时停止，不自行拼文本替代。
- 重复表头/行不拆分用 tableRows、tableIndex、rowIndices 与 table.repeatHeader/cantSplit 布尔值；表头必须从首行连续；残留行高导致不一致时可用autoHeight=true恢复内容自适应，不保证所有行等高。

- 全表字体用tableText；纯嵌入图片段落居中用inlinePictures，避免只改读取样本。
- 表格版心适配用fitTableWidth与tableIndex独立提交，默认保留比例和合并；明确需要调整列宽可传columnRatios，数量须等于网格列数；复杂表格拒绝时单独报告。
- 已有题注布局用normalize_caption_layout，安全文本框移入正文后重读并格式化；受阻对象不删、不展开。

## 验收要求

- 回读图片所在段落的对齐，确认图像数量没有变化。
- 图片视觉位置、环绕和跨页仍需实际页面确认。
- 回读指定单元格的全部目标段落对齐。
- 检查邻近未选单元格，确认没有误做全表统一。
- 核对题注numbering与sequenceIdentifiers；manual-or-unverified不是自动题注，Caption样式或移入文档流不等于SEQ编号，图表关联仍须验证。
- 修改后重读结构与 revision；核对 tables.rowDetails/cellDetails，新增题注后不可继续使用旧段落索引。
- 任意浮动定位仍未支持；不一致兼容分支、关联不明的题注拒绝转换，必须逐项报告，不用成功数量代替全文完成。
- 段落属性通过不代表图表整体视觉通过。

## 案例与失败边界

- 案例：编号列居中、说明列左对齐，按实际单元格分别处理。
- 案例：图片独占一个段落，可设置该段 Centered 后回读。
- 已有手工题注可独立提交convertCaptions和indices：完整选择从1连续的同组编号；保留原文，复杂域或断档拒绝，章号前缀仍静态，需Word更新域验收。
- 反例：把 captionTables 的 index 直接传给段落格式 indices。
- 反例：将同一行的第三个物理单元格当成合并表格第三视觉列。
- 反例：将所有图表段落设固定 12pt 行距以统一正文。
- 缺少对象信息时只处理可确认的部分，并明确剩余对象。
- 图注与图同页、表题与表同页需单独视觉验收，不凭编号证明。
