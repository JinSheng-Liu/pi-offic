# 英文字体与首行缩进修正

## 查明的问题

毕业论文最新 v2 将大部分段落的 ascii/hAnsi/eastAsia/cs 字体槽都写为宋体，因此正文中的 SPI、D/C 等英文没有使用要求的 Times New Roman。英文摘要单独指定 Times New Roman 的操作不能覆盖正文中已经写死的西文字体。

同一 v2 中 16 个普通正文段落同时有一个前导 Tab 和 21 pt 首行缩进；截图文字首行相对续行多偏 21 pt。标题与列表中的 tab/左缩进可能承担结构含义，所以没有批量清除。

## 已实现

- `font.latinName` 单独写入 ascii/hAnsi，保留 eastAsia/cs；统一 `font.name` 仍能设置所有槽位。
- 文字范围格式支持单独指定西文字体，并保存重读校验。
- inspect 返回前导空白和缩进冲突提示；AI 修改前必须查看，不能仅凭段落属性值声称视觉缩进正确。
- 旧后端能力标识增加 `latin-font-slot`，前端发现旧版时阻止发起 AI 编辑并提示重启。

## 最新论文修正版验证

- 基于已校正缩进的验收副本，将全文西文槽写为 Times New Roman；99 段、326 个文本运行检查通过，中文东亚字体、段落格式、原文、7 个媒体文件不变。
- LibreOffice PDF 共8页，文本提取中的 `SPI` 字体为 TimesNewRomanPSMT（Times New Roman PostScript 名称）。正文截图对应两段首行测量为91.95 pt，续行70.95 pt，保留21 pt/两个汉字的首行缩进。
- 修正版：data/validation/latin-font/论文_英文与缩进修正版.docx；PDF: data/validation/latin-font/论文_英文与缩进修正版.pdf。
- 25项服务端回归、API回归、生产类型检查和构建通过。bundle约4.69 MB警告仍在。

用户工作区记录仍为 v2，没有写回 v3；用户5173服务未停止，检查确认其尚未提供 latin-font-slot 能力。需重启后端后新功能才会生效。PDF为LibreOffice渲染结果，不等于 Microsoft Word 保真验收。
