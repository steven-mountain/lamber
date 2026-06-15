# 智算统一控制 ICT 项目周期与折现率

- **Status:** Done
- **Objective:** 将项目周期和折现率统一为智算稳定参数，消除 ICT 页面与智算后台因输入来源不同产生的 NPV 偏差。

## Progress

1. [x] 新增稳定参数 `discount-rate / discount_rate`，智算按百分数输入，ICT 按小数接收。
2. [x] 参数区增加项目折现率醒目入口，与项目周期共同作为 ICT 核心参数。
3. [x] 蓝图升级为 Version 4；Version 1-3 项目从当前正式折现率强制初始化一次，避免默认值或热更新时序意外覆盖。
4. [x] 同步无条件覆盖 lifecycle 输入、lifecycle parameters、cashflow assumptions 和 Rust 正式计算输入中的折现率。
5. [x] SQLite 原子事务同步更新项目表 `discount_rate`，避免项目汇总保留旧值。
6. [x] 智算效益结论增加 ICT 项目净现值率。
7. [x] 回归测试覆盖百分数归一化、旧项目初始化、折现率覆盖和项目表原子更新。

## Validation

- `npm run test:ai-compute-quote`
- `npx tsc --noEmit`
- 定向 ESLint
- `npm run build`
- Rust 项目状态测试

## Scope Boundary

- 项目周期和折现率由智算优先控制。
- ICT 仍负责产权、现金流模型、科目人工覆盖及其他正式财务状态。
- 未修改 Rust 财务公式，只统一计算输入来源。
