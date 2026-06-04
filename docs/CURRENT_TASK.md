# ICT 科目级资金计划最终收官

- **Status:** Done
- **Objective:** 将 ICT 生命周期资金模型收敛为“按具体科目维护的 10 年收付款计划”单一正式口径，完成旧项目一次性迁移、旧模型 UI 移除、正式计算分支收敛与发布验收。

## Completed

1. [x] **旧项目一次性迁移**
   - 增加 `subjectFundingPlanMigrationVersion = 1` 标记。
   - 对缺失计划的非零收入/投入科目生成第一年一次性迁移计划。
   - 保留已有有效计划，不静默覆盖异常计划。
2. [x] **单一正式计算口径**
   - 新建、导入、加载项目均收敛为 `subject_funding_plans`。
   - 正式现金流、NPV、IRR、回收期和 Excel 多年现金流均来自科目年度计划汇总。
   - `CashflowSegment` 旧年度金额不再参与正式现金流覆盖。
3. [x] **旧模型用户入口移除**
   - 移除模型 A-E / 分板块资金模型配置入口。
   - 移除“原资金模型 / 科目收付款计划”切换控件和旧模型文案。
4. [x] **联动能力回归**
   - 智能反算、差额承接、CT 金额联动继续通过统一金额入口同步科目计划。
   - custom / equal / upfront 计划保存恢复、覆盖定位、一键清空、项目整体/IT 10 年现金流继续可用。

## Validation

- `for f in scripts/test_subject_funding*.cjs; do node "$f"; done` in `src-ui`: passed.
- `npx tsc --noEmit` in `src-ui`: passed.
- `npm run build` in `src-ui`: passed.
- `cargo test` in `src-tauri`: passed (11/11; existing warnings only).
- `cargo fmt -- --check` in `src-tauri`: passed.

## Remaining

- 页面级手工验收仍需在真实工作区中打开旧项目、执行保存刷新、文档导出并核对 Excel 多年现金流。

## Next

- 发布前全流程验收与安装包 smoke test。
