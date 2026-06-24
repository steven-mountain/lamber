# 采购甄选费写入效益分析表

- **Status:** Done
- **Objective:** 在 ICT 测算生成《效益分析表》时，将采购甄选费面板中的供应商报价和甄选上浮写入 `3-直接经济效益评估表` 的 `T26`、`T29`。

## Progress

1. [x] 将甄选测算面板的报价、上浮、实际测算成本、甄选费、最高限价和固定锚点作为 `IctInput` 可选元数据序列化。
2. [x] 项目/方案恢复时回填甄选面板状态，切换项目或无数据时清空为默认状态。
3. [x] 文档生成变量包在甄选面板已有报价、限价或测算结果时输出 `SELECTION_FEE_SUPPLIER_QUOTE` 和 `SELECTION_FEE_MARKUP`。
4. [x] Rust Excel 生成将供应商报价写入 `3-直接经济效益评估表!T26`，将甄选上浮写入 `3-直接经济效益评估表!T29`。
5. [x] 生命周期效益表导入解析同步回读 `T26/T29`，手动导入和后台自动导入保存到同一可选字段。
6. [x] 同步更新采购甄选费模块文档和 AI 变更日志。

## Validation

- `cargo test --manifest-path src-tauri/Cargo.toml selection_fee`：通过 2 个新增测试；保留仓库既有 Rust warnings。
- `npm run build --prefix src-ui`
- `npm run lint --prefix src-ui`：失败于既有 `src-ui/src/store/useAiContextStore.ts:107` 的 `@typescript-eslint/no-this-alias`；同时保留若干历史 warnings。

## Scope Boundary

- 未修改甄选费用、现金流、NPV、税额、科目金额、0 容差校验或反算公式。
- 新增甄选元数据只用于状态恢复、AI 只读上下文、Excel 生成回填和同类 Excel 导入回读。
- 未修改模板结构、行列布局或除 `T26/T29` 外的 Excel 坐标。
