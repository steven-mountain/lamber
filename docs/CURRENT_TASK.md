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
