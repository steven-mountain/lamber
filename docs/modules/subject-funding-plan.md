# 科目级资金计划模块设计

本模块定义 ICT 生命周期测算的正式资金收付模型。发布口径下，科目级资金计划是唯一年度收付款事实来源；旧资金模型字段仅用于读取旧项目和一次性迁移兼容，不再作为用户入口或正式计算分支。

## 核心规则

1. 资金计划绑定到具体科目实例，稳定主键为 `side + groupId + key`，例如 `revenue:revIt:integration`。
2. 每个非零收入科目必须有启用中的 10 年收款计划；每个非零投入科目必须有启用中的 10 年付款计划。
3. 年度计划支持 `upfront`、`equal`、`custom` 三种模式，年度含税金额合计必须与科目含税金额一致。
4. 零金额科目不要求生成计划；若零金额科目存在非零年度计划，覆盖校验必须阻塞正式保存、文档生成和现金流刷新。
5. 金额变更统一由 `updateTaxItem` / `updateTaxItemsInclBatch` 汇入资金计划同步逻辑。正金额缺失计划时默认创建第一年一次性计划；已有计划按原年度比例缩放；金额清零时移除该科目计划并回到“未维护”。
6. 智能反算、差额承接、CT 产品/专线联动均通过同一金额更新入口同步资金计划，并写入 `lastChangeReason` 供 UI 解释。

## 智能反算边界

普通成本反算通过前端候选金额构造完整科目状态和同步后的科目资金计划，再调用 Rust 计算器评估指标。净现值率定义为 `NPV / 折现后现金流出`；当现金流出恰好为 0 时，后端为避免除零按既有约定返回 0，该值不能直接代表成本反算的最大可达净现值率。

因此净现值率成本反算必须同时探测：

- `0` 元，用于存在其他现金流出时的正常边界；
- `0.01` 元，即系统最小货币单位，用于总现金流出为 0 时进入可计算区间。

二分搜索从两者中指标更高的有效候选开始。毛利润率没有分母为零的同类问题，继续使用 `0` 元边界。该规则只修正反算搜索边界，不修改 NPV、NPVR、税率、年度现金流或科目资金计划公式。

## 旧项目迁移

迁移版本标记为：

```ts
subjectFundingPlanMigrationVersion = 1
```

加载项目时，若缺少该标记或旧数据未具备完整科目计划，系统执行一次性迁移：

1. 对每个含税金额大于 0 的收入/投入科目生成第一年一次性计划。
2. 迁移计划字段为 `mode: "upfront"`、`enabled: true`、`source: "migration"`、`lastChangeReason: "legacy_migration"`。
3. 零金额科目不生成非零年度计划。
4. 已存在有效 `custom` / `equal` / `upfront` 计划的科目不被覆盖。
5. 已存在异常计划的科目不被静默重置；系统仅补齐完全缺失的非零科目计划，随后由覆盖校验阻塞正式计算/保存。
6. 覆盖校验通过后写入迁移版本标记，后续打开不重复迁移。

本项目现有旧项目均按第一年一次性收付款，因此迁移后的年度收入、投入、项目整体现金流、IT 现金流、NPV、IRR 和回收期应与旧第一年一次性口径等价。

## 正式计算口径

`useIctCalculations` 始终构造 `cashflow_calculation_source: "subject_funding_plans"`。覆盖校验通过后，`buildAnnualCashflowFromSubjectFundingPlans()` 按科目税率将年度含税收付款换算为年度不含税数组，并传给 Rust 计算器：

- `rev_cashflow_excl`
- `cost_cashflow_excl`
- `it_rev_cashflow_excl`
- `it_cost_cashflow_excl`

覆盖校验不通过时，页面保留上一次有效现金流/指标，不回退旧模型计算，并阻塞效益指标保存和文档生成。

## CashflowSegment 职责

`CashflowSegment`、`cashflowModel`、`paymentModelJson`、`sectorCashflowJson` 中的旧资金模型字段可继续存储，以兼容旧项目读取、历史快照和迁移前数据留存；它们不再贡献正式年度收入、投入、NPV、IRR、回收期或 Excel 多年现金流。

分类展示如项目整体、IT 部分和年度穿透明细，均从具体科目计划汇总生成。Excel 投资效益分析模板的多年现金流变量读取正式 `metrics.cashflow`，与页面现金流一致。

## UI 与保存

页面不再展示“原资金模型 / 科目收付款计划”切换，也不再展示模型 A-E 或分板块资金模型配置入口。用户只维护具体科目的收款/付款计划、覆盖状态、批量初始化、定位异常、一键清空、智能反算、差额承接和 CT 联动。

保存路径会在 lifecycle input、cashflow assumptions 和 benefit snapshot 中携带 `subject_funding_plans`、`cashflow_calculation_source` 与 `subjectFundingPlanMigrationVersion`。旧字段保留但不再驱动正式计算。

## 测试基线

核心脚本：

- `scripts/test_subject_funding_plan.cjs`
- `scripts/test_subject_funding_cashflow.cjs`
- `scripts/test_subject_funding_sync.cjs`
- `scripts/test_subject_funding_phase4.cjs`
- `scripts/test_subject_funding_migration.cjs`
- `scripts/test_subject_funding_final.cjs`
- `scripts/test_ict_reverse_search.cjs`

构建验收：

- `npx tsc --noEmit`
- `npm run build`
- `cargo test`

最后更新：2026-06-10。
