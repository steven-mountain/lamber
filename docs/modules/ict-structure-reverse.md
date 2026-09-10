# 锁定总额结构反算

入口为 `useIctCalculations.performReverseCalculation` 的 `locked_total_structure` 分支，桌面操作调用原 `performLockedTotalStructureReverseCalculation`。AI 的 D 卡片经用户确认调用同一入口，独立验证记录见 `../verification/ai-structure-reverse-card.md`。

**2026-09-09 已修复候选与实际提交不一致。** 候选和批量入口共用 `ictTaxItemBatch.prepareIctTaxItemsInclBatch`，按有效归一金额、拆分清除、CT联动和科目计划求值；总额偏差由承接吸收并再归一，仍不闭合则拒绝。最终参数再次用于构造CT联动状态，确保不同税率的联动成本也与真实写入一致。[独立验证](../verification/structure-commit-normalization.md)。D仍须独立验收。

## 成功与写入契约

金额按分取整、年度计划离散分配，因此目标位于采样指标的最小值与最大值之间，并不证明目标可达到。搜索结束必须同时验证结果自洽和实际达标。

1. 沿用原采样、45 次二分上限、`increasing` 判向及金额舍入。中点命中目标时仍沿用原命中处理；仅区间塌缩时退出，不用较差的最后中点覆盖已追踪的最佳候选。
2. 只有指标差不超过 `METRIC_EPSILON = 0.0001` 的候选进入解集。再从达标候选中选择相对原目标科目金额改动最小者。
3. 所选候选重新求值后，依次验证候选/载荷/结果有效、同侧总额不变、金额非负、最终指标达到目标。保留 `MONEY_EPSILON = 0.004`。不得放宽容差或自动改写用户目标。
4. 所有验证、结果上下文及成功提示准备完毕，才进入金额提交阶段。拒绝和准备阶段异常不调用 `updateTaxItemsInclBatch`，也不清除原忽略尾差标志或发布新的结果上下文。准备阶段的异常处理不包围金额提交，以免把已经进入写入阶段的运行时异常误称为搜索未达标。
5. 成功仍一次批量提交目标/承接两科目，使用原 `reverse_calculation_sync` / `balance_allocation_sync` 原因，沿用原成功提示和结果字段。

## 不可达诊断

不敏感、超范围、未找到稳定解及最终复算未达标的提示都保留具体数字：用户目标、本次有效求值中最接近的指标、采样得到的最小值–最大值。明确采样范围不表示中间每个值都能达到；“最接近”只指本次搜索结果，不承诺未搜索区域的全局最优。

最终复算若与搜索时不一致，还必须单独显示最终复算值，不能把先前找到的值冒充当前复算结果。诊断仅供用户决定是否调整目标。

## 资金计划与验证

正式写入经 `updateTaxItemsInclBatch` 同步科目收付款计划；`modelEAmountMode` 三处仍为 `false`，不启用或删除遗留分板块同步代码。财务公式、税额、资金计划同步和零容差核验规则不变。

`npm run test:structure-reverse` 执行直接提取的生产入口/反算函数，候选与引擎为受控测试桩。每条拒绝用金额写入 spy 的调用次数断言 0，同时校验目标不变。成功路径与修复前冻结源码比较，另有原批量入口和资金计划回归。[独立验收记录](../verification/structure-reverse-success-check.md)。
