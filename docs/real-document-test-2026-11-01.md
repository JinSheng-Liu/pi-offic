# 用户公文真实模型测试

结论：部分通过，不能按全部排版需求验收。

原件：D:/xiangmu/agent-office/agent-office/测试文档/党政机关公文_无格式版.docx
原件 SHA-256：1ec3bab1cb3ad9cbac353de2ca16a781f801896d03b3c1794888e246a733cea3（重放前后相同）。

## 执行
1. 读取实际失败会话：颜色为 FF0000/000000；对齐为 Center/Justify，后端要求 #RRGGBB 和 Centered/Justified。
2. 修复引擎兼容明确别名，并把字体和段落工具 schema 从任意 Record 改为字段定义。
3. 用原件重放两组实际失败操作，均输出成功，30段文本一致。
4. 使用5174当前配置的真实模型，导入独立副本 d523414e-bdeb-4f54-b020-e354f356d204。发送原排版要求，并附带旧服务所需的规范颜色、对齐值。没有重启当前服务，因此本次不是新 schema 上线后的无辅助提示验收。
5. 真实请求完成，无 SSE error，副本 v1 -> v2；下载文件与原件30段文本逐段一致。
6. 直接检查输出 ZIP 内 document.xml：22pt/16pt、红黑颜色、center/both/right、正文 line=560 lineRule=exact、firstLine=640、段前后0，A4纵向及目标边距均已写入（twips取整）。

## 失败与限制
- 五处同段标题局部黑体未完成。
- 字体仅设置 ascii/hAnsi，缺少显式 eastAsia；不能认定中文字体正确。
- 未验证指定字体已安装、Word实际视觉排版、浏览器重新打开导出文件的完整交互。
- 新后端兼容与 schema 代码尚未加载到现有服务，不代表用户可直接无辅助重试。
- 工具验证 passed 只覆盖工具操作声明，不能代表全部用户需求完成。

## 产物
- data/validation/real-document/live-model-result.docx：真实模型输出。
- data/validation/real-document/xml-check.json：独立 XML 抽查。
- data/validation/real-document/live-inspection.json：完整结构读回。
- replay-1.docx / replay-2.docx：修复后重放输出，不是模型新生成结果。

回归：check:server 7/7；check:api 1/1。
