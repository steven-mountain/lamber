# 智算金额来源自动同步与税率口径修复

- **Status:** Done
- **Objective:** 恢复智算金额变更后的自动 ICT 同步，统一智算与 ICT 的金额/税率口径，并完善金额来源命名、复制和基底选择流程。

## Progress

1. [x] 智算页面按业务指纹防抖自动保存当前金额来源，并调用 `sync_intelligent_compute_to_ict` 写入 ICT 正式状态。
2. [x] 自动同步成功后刷新 ICT 测算结果、顶部同步状态和项目汇总；手动“同步到 ICT / 查看同步明细”保留为预览、排查和重试入口。
3. [x] `ICT_SUBJECT_DEFINITIONS` 增加 `defaultTaxRate`，ICT 水合、智算导出和资金计划 finalizer 统一使用标准默认税率。
4. [x] 智算导出 payload 保留含税金额、不含税金额和税率；当前来源内同一 ICT 科目先汇总含税/不含税，再反算有效税率。
5. [x] 旧受控科目释放逻辑同时识别 `controlledSubjects` 和 `source: intelligent_compute` 导入痕迹，避免改映射或改来源后残留旧金额。
6. [x] 金额来源管理改为弹窗入口；新建来源必须命名，可选择空白、H200 标准、当前来源或任一已有来源作为基底。
7. [x] 金额来源改为单选同步来源；新建来源自动成为当前唯一 ICT 同步来源，后端保存兜底关闭其他来源。
8. [x] 智算未产出的 ICT 标准收入/成本科目全量写 0，冲突时提供重新加载后完全覆盖入口。
9. [x] 更新智算业务测试、模块文档和 AI 变更日志。

## Validation

- `npm run build`
- `npm run test:ai-compute-quote`
- `cargo test`
- `git diff --check`

## Scope Boundary

- 保留 ICT 正式 Rust 计算引擎、科目目录、资金计划结构和财务公式。
- 自动同步只读取当前选中的智算金额来源，并完全覆盖 ICT 标准收入/成本科目的金额、税率和资金计划。
- 同步金额口径固定为“元、含税”，万元仅用于页面展示。
- 不新增数据库表，不做 schema migration；继续使用现有 `intelligent_compute_amount_sources`。
