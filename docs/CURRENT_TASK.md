# 甄选后流程 · 第 2 阶段：ICT 测算表内"甄选前 / 甄选后"方案切换

- **Status:** Done（已用全新项目验证：两套方案工作副本物理隔离、切换互不影响）
- **Objective:** 让用户在 ICT 测算表页面内直接切换"甄选前 / 甄选后"两版效益分析，无需跳出到项目看板；无"甄选后"方案时可一键从当前方案派生创建，并保证左侧"一键生成全流程文档"始终使用当前选中方案数据。

## Progress（第 1 阶段：数据层打标，已完成）

1. [x] `BenefitAnalysisScheme` 增加可选 `stage` 字段（`pre_selection` / `post_selection` / `None`），serde 默认兼容历史数据。
2. [x] `benefit_schemes` 表新增 `stage TEXT` 列；schema_version 迁移 7 → 8（幂等 ALTER）。
3. [x] `save_benefit_analysis` 增加 `stage` 入参：新建按传入写入，更新用 `COALESCE(?, stage)` 保留原标签。
4. [x] 新增 `update_scheme_stage` 命令（独立改标签、不产生新快照），并在 `main.rs` 注册。
5. [x] 读取路径 + JSON/SQLite 仓储 + JSON→SQLite 迁移全部同步读写 `stage`。
6. [x] 前端：`lib/schemeStage.ts` 常量；`ProjectBoard` 方案 chip + 分段按钮设置阶段。

## Progress（第 2 阶段：测算表内方案切换，本次完成）

7. [x] `IctLifecycle` 顶部横幅由静态"当前方案"文本改为二段式切换控件 `[甄选前][甄选后] 更多方案 ▾`。
8. [x] 组件内维护 `schemes` 列表，按 `updated_at` 倒序推导每阶段主方案（`preScheme` / `postScheme`），其余（未标注或同阶段历史方案）收纳进"更多方案"下拉。
9. [x] 点击已存在阶段：先走 `confirmOrSave()` 未保存确认，再复用 `navigateTo("ict_lifecycle", pid, schemeId)` 加载路径（与项目下拉/返回一致，不新建并行状态）。
10. [x] 点击"甄选后（未生成）"：当前方案未打标签 → 就地标注；当前方案已属另一阶段 → 复用"另存为新方案"弹窗派生新方案（默认名 `${项目名}_甄选后`，`stage=post_selection`，以当前数据为起点）。
11. [x] 文档生成安全提示：点击含"签批"的模板且 `activeScheme.stage !== "post_selection"` 时，`handleTabSwitch` 给出非阻断 `confirm`，用户可继续或取消。
12. [x] `ProjectBoard` 方案列表保留阶段徽标与分段设置按钮（第 1 阶段已具备）。
13. [x] 修复：派生"甄选后"后点击"甄选前"无法切换（`activeScheme` 与导航 `activeSchemeId` 漂移导致 `navigateTo` 写同值不触发加载 effect）—— `switchToScheme` 同步 store 并在 store 已指向目标 id 时直接 `loadProjectContext` 兜底。
14. [x] 修复：阶段切换控件位置随方案名长短漂移 —— 控件移到当前方案行固定最左侧（`shrink-0`），方案名右置并 `truncate`。
15. [x] 修复：切换方案时顶部项目卡片跳动 —— `loadProjectContext` 同项目内切换不再 `setActiveProject(null)` 等整块清空（`isSameProject` 判定），避免卡片瞬间闪成"自由测算模式"再切回。
16. [x] 修复（严重数据串档，架构级）：改甄选后金额连甄选前也被改 —— 根因是工作副本 `project_lifecycle_states`/`project_cashflow_states` 为 `project_id UNIQUE` 的项目级单例，两方案共用一行。改为**按方案存储**：
    - schema 迁移 v8→v9：两表加 `scheme_id`，唯一键改 `(project_id, scheme_id)`，既有行回填到 `default_scheme_id`。
    - 后端 `save/get_lifecycle_state`、`save/get_cashflow_state` 新增 `scheme_id` 入参；智算导入/`get_project_full_state` 用默认方案桶。
    - 前端 `domainSaveService` 四方法加 `schemeId`；`IctLifecycle` 保存 handler / `persistLifecycleAndCashflowState` 带 `activeScheme.id`；`loadProjectContext` 按选中方案加载其独立草稿。
    - 取代早前基于 `default_scheme_id` 的临时判定（因金额编辑只落 cashflow scope 不成立）。

## Validation

- `cargo test`：29 passed（含 `v7_benefit_schemes_gains_stage_column_and_preserves_rows`、`fresh_database_uses_schema_v8_*`）。
- `npm run build --prefix src-ui`：通过。
- `npm run lint --prefix src-ui`：本次改动无新增告警/报错（仅 `useAiContextStore.ts` 一处历史 `no-this-alias` 报错，属既有、与本任务无关）。

## Scope Boundary

- 未改动 `calculator.rs` 测算引擎、NPV、现金流、税额、科目金额、甄选费、反算或 0 容差校验。
- 未改动快照（Snapshot）结构与版本号逻辑、Excel/Word 模板结构与 `T26/T29` 坐标映射。
- 复用既有"保存到当前项目""另存为新方案"能力，未重写其行为。

## Next (第 3 阶段，待用户提供模板)

- 《甄选结果签批表》docx：等用户提供模板 → `TemplateForms.tsx` 加专属 Tab + 变量映射，默认取"甄选后"方案 + 采购甄选费面板数据，走现有 `generate_lifecycle_docs`。


# 单科目含税金额拆分闭合

- **Status:** Done
- **Objective:** 对单笔含税金额无法按“不含税为锚点”双向闭合的科目，提供保持科目含税总额不变的两笔拆分能力。

## Progress

1. [x] 使用十进制半进位搜索两笔可独立闭合的含税金额，并保持两笔含税合计严格等于原科目金额。
2. [x] 拆分明细随工作副本序列化和还原；金额或税率再次编辑时自动失效，避免沿用过期拆分。
3. [x] 前端展示两笔含税/不含税明细并支持取消拆分；资金计划仍以原科目含税总额为锚点。
4. [x] `npm run test:tax-split --prefix src-ui` 与 `npm run build --prefix src-ui` 通过。

## Scope Boundary

- 本阶段只处理单科目自身的两笔闭合，不包含税率组整体尾差的候选科目推荐。
- 不改变原科目含税总额、税率或资金计划金额；拆分必须由用户主动确认。


# 税率组整体尾差拆分建议与应用

- **Status:** Done
- **Objective:** 对税率组 `[汇总误差-公式C1]` 提供可审计的单科目两笔拆分建议，并在用户确认后应用、保存和继续原流程。

## Progress

1. [x] 按税率组实际尾差反向计算所需不含税调整，只接受能让尾差精确归零的候选。
2. [x] 候选保持原科目含税总额和税率不变，两笔子金额分别通过双向闭合校验。
3. [x] 拦截弹窗最多展示三个候选科目及两笔含税/不含税金额、拆分前后尾差；没有精确候选时明确说明。
4. [x] 每个候选提供“应用并继续 / 应用此拆分”，通过统一状态入口再次校验子笔；唯一错误消除后继续原现金流或文档页。
5. [x] 科目变更同时标记 lifecycle、cashflow、benefit-analysis；工作副本合并和方案快照均保留 `split_parts`。
6. [x] Rust 计算引擎验证拆分后按子笔不含税之和计算，前后端口径一致。
7. [x] `npm run test:tax-split --prefix src-ui`、`npm run test:selection-batch --prefix src-ui`、`npm run build --prefix src-ui` 与 `cargo test` 通过。

## Scope Boundary

- 不经用户点击不应用拆分；应用只改变科目内部含税/不含税舍入分配，不改变科目含税总额、税率或资金计划。
- 仅处理一个科目拆成两笔即可归零的汇总尾差；不生成多科目组合或近似建议。
