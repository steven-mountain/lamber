# 测算方案甄选阶段（甄选前 / 甄选后）模块设计

同一个 ICT 项目下可以并列存在多套效益测算方案（`BenefitAnalysisScheme`），每套方案挂载若干只读快照（`BenefitAnalysisSnapshot`）。本模块在方案层引入"甄选阶段"标签，用于区分立项评估用的**甄选前**方案与拿到中标报价后重做的**甄选后**方案，并支持在 ICT 测算表页面内直接切换两版方案。

## 数据模型

- `BenefitAnalysisScheme.stage: Option<String>`，取值约定：
  - `"pre_selection"`（甄选前）
  - `"post_selection"`（甄选后）
  - `None`（未标注，历史数据或临时/敏感性分析方案的默认值）
- SQLite `benefit_schemes` 表新增 `stage TEXT` 列；`schema_version` 由 7 迁移到 8，迁移为幂等 `ALTER TABLE ... ADD COLUMN`，既有行 `stage` 默认 `NULL`。
- serde 对 `stage` 使用 `#[serde(default)]`，历史 JSON 无此字段时反序列化为 `None`，向后兼容。
- 阶段标签**仅用于方案的区分、展示与后续取数**，不参与 `calculator.rs` 的 NPV、现金流、税额、科目金额、甄选费或 0 容差校验，也不改变快照结构与版本号逻辑。

## 后端命令

- `save_benefit_analysis(..., stage: Option<String>)`：新建方案按传入 `stage` 写入；更新既有方案用 `stage = COALESCE(?, stage)`，即传 `None` 时保留原标签（普通保存不清空）。传入前统一经 `normalize_scheme_stage` 归一化，非法/空值视为未标注。
- `update_scheme_stage(project_id, scheme_id, stage)`：独立改写方案阶段标签，**不产生新的效益快照**，返回更新后的方案。空/非法值归一化为 `None`（取消标注）。
- 读取路径（`get_benefit_schemes`、`get_project_full_state`）、JSON/SQLite 双仓储与 JSON→SQLite 迁移全部同步读写 `stage`。

## 前端交互

统一常量与配色在 `src-ui/src/lib/schemeStage.ts`（`SCHEME_STAGE_OPTIONS`、`getSchemeStageOption/Label`、`normalizeSchemeStage`）。

### ICT 测算表（`IctLifecycle.tsx`）

顶部横幅"当前方案"处提供二段式切换控件 `[甄选前][甄选后] 更多方案 ▾`：

- 组件内维护 `schemes` 列表（在 `loadProjectContext` 中随项目加载/重置），按 `updated_at` 倒序推导每阶段主方案 `preScheme` / `postScheme`；同阶段多余方案与未标注方案收纳进"更多方案"下拉。
- **切换已存在方案**：先走 `confirmOrSave()` 未保存确认，再 `navigateTo("ict_lifecycle", projectId, schemeId)`，复用与项目下拉/返回按钮一致的加载路径，`activeScheme` 与各 Tab 数据随之刷新，不新建并行状态。
  - 注意：保存/派生方案走 `loadProjectContext` 直接加载、不刷新导航 store，`activeSchemeId` 可能陈旧地仍指向目标方案。此时 `navigateTo` 写入相同值不会触发 `[activeProjectId, activeSchemeId]` 加载 effect，故 `switchToScheme` 在 store 已指向目标 id 时会直接调用 `loadProjectContext` 兜底，避免"切换看似无效、需退出重进"。

控件位置：阶段切换控件固定在"当前方案"行的最左侧（`shrink-0`），方案名放其右侧并按需 `truncate`，切换控件位置不随方案名长短漂移。
- **"甄选后（未生成）"**：
  - 当前方案未打标签（`stage == null`，如首次测算生成的方案）→ 就地 `update_scheme_stage` 标注为该阶段。
  - 当前方案已属另一阶段 → 复用"另存为新方案"弹窗派生新方案，默认名 `${项目名}_甄选后`，`stage=post_selection`，以当前测算数据为起点，用户可按甄选实际报价调整后保存。
- **更多方案下拉**：选择后同样走 `confirmOrSave()` + `navigateTo` 切换，未标注方案按普通方案处理，不强制打阶段标签。

### 项目看板（`ProjectBoard.tsx`）

方案列表 chip 展示阶段徽标（甄选前 / 甄选后不同底色），快照面板提供"甄选前/甄选后"分段按钮设置/取消标注，保留原有列表交互。

## 多方案数据隔离（工作副本按方案存储，重要）

测算的"当前工作副本"存放在 `project_lifecycle_states` / `project_cashflow_states` 两张表。**自 schema v9 起，这两张表按 `(project_id, scheme_id)` 存储**（此前是 `project_id UNIQUE` 的项目级单例），从而每个测算方案拥有各自独立的工作副本，甄选前/甄选后互不影响。

- `scheme_id` 为空串 `''` 表示项目级默认草稿桶（智算导入的初始种子、legacy 兼容）。
- 迁移 v8 → v9：两表新增 `scheme_id`，唯一键由 `project_id` 改为 `(project_id, scheme_id)`；既有行按项目 `default_scheme_id` 回填（无默认方案则归入 `''`）。

读写路径：

- 后端命令 `save_lifecycle_state` / `save_cashflow_state` / `get_lifecycle_state` / `get_cashflow_state` 均带 `scheme_id` 入参，按 `(project_id, scheme_id)` upsert / 读取。
- 智算导入（`apply_ai_compute_quote_to_ict_locked`）与 `get_project_full_state` 使用项目默认方案桶（`resolve_default_scheme_bucket` = `default_scheme_id` 或 `''`），供智算视图/文档导出等"按项目取数"的消费方沿用原语义。
- 前端 `IctLifecycle`：保存 lifecycle/cashflow 的 handler 与 `persistLifecycleAndCashflowState` 都带当前 `activeScheme.id`；`loadProjectContext` 按选中方案精确加载**它自己的草稿**（`loadLifecycleState/loadCashflowState(project, schemeId)`）——无草稿则回落到该方案的快照（`benefit_snapshots.input_params`，`buildInputDataPayload` 已完整包含科目金额、资金计划、现金流分段、差额承接规则、甄选费等），无选中方案才回退到项目默认桶 / legacy 输入。

注意：科目金额编辑在前端标记的是 `cashflow` scope（保存写入该方案的 cashflow 草稿桶），加载时 `buildHydrationInput` 会用 cashflow 草稿的 `assumptionsJson` 覆盖各科目金额，因此仅保存 cashflow scope 也能在重载时完整还原编辑。

### 已知限制：v9 之前创建的旧项目

v9 之前工作副本是项目级单例，两个方案物理上共用同一行数据。v8 → v9 迁移只能把这唯一一行归属给项目 `default_scheme_id`，另一方案没有独立副本、会回落到自己的快照——若旧快照也带有当年共用状态的污染，就会出现"改一个方案连另一个也变"的纠缠现象（且方向取决于哪个方案在迁移时是 `default_scheme_id`）。这属于**历史数据纠缠**，非当前代码缺陷：全新项目下两套方案自创建起即为独立两行，完全隔离（已验证）。

处理旧项目的两种方式：分别切到每个方案、录入正确金额并「保存到当前项目」各存一次（存后即各自独立）；或删除其中一个方案按新流程重建。

## 文档生成安全提示

`IctLifecycle` 左侧"一键生成全流程文档"中，点击模板名包含"签批"的模板（对应《甄选结果签批表》等）且 `activeScheme.stage !== "post_selection"` 时，`handleTabSwitch` 弹出**非阻断** `confirm`："当前方案为甄选前，建议切换到甄选后方案再生成"。用户可选择继续或取消，不强制拦截，以兼容甄选前也可能需要预生成草稿的场景。

## 集成边界

- 甄选前与甄选后是**同一项目下的不同方案**，靠 `stage` 打标而非新建项目，便于对比与后续《甄选结果签批表》取数。
- 复用既有"保存到当前项目""另存为新方案"能力，未重写其行为；未改动 Excel/Word 模板结构与 `T26/T29` 坐标映射。

最后更新：2026-07-01。
