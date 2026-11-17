# 自动处理统一工作流整改

AI自动格式处理使用docx4j；人工编辑留在DOCX Editor。engine.mjs现在是兼容导出，处理实现位于server/document：package（包与结构保护）、structure（结构读取）、rules（参数规则）、docx4j-adapter（格式执行）、ooxml-sections与ooxml-captions（受控补充）、workflow（自动处理批次）。

DocumentStore持有Workflow。自动格式及题注都经版本锁、候选文件重读、取消检查后提交。每次工具调用是一个原子批次，已提交批次在取消后保留，通过版本历史恢复。成功磁盘提交产生document SSE事件，前端按版本而不是工具名刷新。

前端发送前自动保存人工修改，保存失败不发起AI并保留输入。AI期间只读，最终同步期间也不允许编辑。新增document-commit-events能力标记阻止新前端误用旧后端。

验证：31项服务端测试、API集成测试、vue-tsc和生产构建通过。API覆盖版本事件、任务锁和取消终止。构建仍有约4.69MB大包提示。

未完成验收：真实浏览器自动保存/重新加载交互、阅读位置恢复、Word视觉保真。本次没有停止5173服务，没有修改用户文档。题注原有能力限制没有因重构消失；本次也没有完成依赖许可证审计或全部格式功能扩展。
