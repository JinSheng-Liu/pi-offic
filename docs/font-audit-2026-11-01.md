# 中文字体链路排查（2026-11-01）

## 证据与根因

用户记录 `4a4aee96-87ef-45f8-ba90-f66204739da1` 的 v2 只写了 rFonts.ascii/hAnsi，没有写 eastAsia。字体名称不是一个统一字段；西文字体写入成功不意味着中文字体改变。现有 inspect 的 font.name 同样只读取西文字体，容易造成错误验收。

5173 的 health 仅返回 ok/modelConfigured，缺少当前源码已有的 testMode；结合该服务新生成文件缺失当前引擎会强制验证的 eastAsia，判断为旧后端仍在运行。未停止用户进程。native 配置加载模式下不能假定后端模块更新会随前端热更新生效，当前需重启 npm run dev。旧历史版本不会因重启自动修复。

本轮为 inspect 增加 eastAsiaDirect/eastAsiaThemes，null 表示没有直接设置，不能据此推断继承后的最终字体。analyze_formatting 同时返回解释限制。既有写入修复写齐 ascii/hAnsi/eastAsia/cs 并清理主题覆盖，本轮对其做真实失败操作重放和浏览器验证。

## 已执行验证

- `node scripts/replay-font-regression.mjs`：旧输出 30 段失败；当前引擎重放后 30 段通过，原文一致。未覆盖用户原文档或历史。
- 新增 inspection 中文/西文字体区分测试：修复前失败，修复后通过。
- `npm run check:server`：11 项通过。
- `npm run check:api`：1 项集成测试通过（其中使用测试模型服务，不代表真实供应商压力测试）。
- `npm run build`：类型检查及构建通过，主包约 4.69 MB 的警告仍存在。
- 隔离服务 5175，真实模型请求“全文仿宋三号，不修改文字”：测试文档 a966c5c0-8952-45f6-84b8-be86904c3d84 v2，30 段/30 个文本运行 eastAsia 均为仿宋。
- 浏览器加载重放副本，DOM 文本字体包含仿宋、黑体、小标宋、方正小标宋简体，没有 MS Mincho。
- 真实模型执行后 DOM 文本字体改为仿宋；调用编辑器原生 save，再通过保存 API 创建 v3，30 个文本运行 eastAsia 仍为仿宋。

## 未解决或未验证

- 同一段中的标题前缀单独设黑体、其后正文保持仿宋：当前格式工具仅支持段落范围，真实旧请求因此遗漏。需要增加文本范围格式能力及结构保护回归，不应宣称完整公文排版完成。
- 未验证 Microsoft Word 中的实际排版、缺失字体替代及所有复杂 DOCX 元素。
- 本次验证 DOM 字体声明和导出文件，不等于确认小标宋等字体已安装或字形完全一致。
- 前端约 4.69 MB 主包存在加载性能风险。
- 尚未进行长时间模型稳定性或并发压力测试。

验收副本及 report.json 保存在 data/validation/font-regression；用于字体回归，不是所有公文要求均完成的交付文件。
