# ICT 净现值率成本反算边界修复

- **Status:** Done
- **Objective:** 修复普通成本反算在零现金流出边界错误提示“无法达到目标值”的问题，不改变现有 NPV/NPVR 公式和科目资金计划口径。

## Completed

1. [x] 回退所有既有本地修改，并同步到 GitHub `origin/master` 的 `9d65b10`。
2. [x] 定位根因：当折现后现金流出为 0 时，Rust 计算器按既有约定返回 `npv_rate = 0`；前端误把该特殊值当作成本反算的最大可达净现值率。
3. [x] 净现值率成本反算现在同时探测 `0` 和最小货币单位 `0.01` 元，从指标更高的有效边界开始二分搜索。
4. [x] 毛利润率反算仍使用原有零成本边界；财务公式、税率换算、科目计划同步和 CT 收入投入联动均未修改。
5. [x] 新增前端纯函数回归脚本，并增加 Rust 测试覆盖 IT 集成收入 `78000`、CT 产品收入/投入 `500`、目标净现值率 `0.10` 的可达场景。

## Validation

- `node scripts/test_ict_reverse_search.cjs`: passed.
- All `scripts/test_subject_funding_*.cjs`: passed.
- `npx tsc --noEmit`: passed.
- `npm run build`: passed; existing Vite large-chunk warning remains.
- `cargo fmt -- --check`: passed after formatting.
- `cargo test benefit::calculator::tests`: passed, 9 tests.

## Remaining

- 使用 Tauri 桌面运行时和真实项目资金计划做一次交互点击验证，确认提示文案、反算结果写回和年度计划同步符合实际项目数据。
