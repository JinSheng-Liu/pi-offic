# 前端组件来源

- `src/components/ui/`：从 [shadcn-vue](https://github.com/unovue/shadcn-vue/tree/dev/apps/v4/registry/bases/reka/ui) 的 Reka 基础组件按需引入；MIT，见 `shadcn-vue.LICENSE`。本地修改包括导入路径和工作台主题样式。
- `src/components/ai-elements/task/`：从 [AI Elements Vue](https://github.com/vuepont/ai-elements-vue/tree/main/packages/elements/src/task) 按需引入；Apache-2.0，见 `ai-elements-vue.LICENSE`。本地修改包括导入路径及语义化列表项支持。
- Kokonut UI 仅用于执行记录折叠、状态反馈的交互参考，未复制其 React 代码或演示进度逻辑。

组件仅负责呈现。执行、取消、文档提交及核验继续由现有 Pi Runtime、服务端和编辑器完成。
