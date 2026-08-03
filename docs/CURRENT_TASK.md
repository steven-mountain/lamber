# 多项目甄选结果签批表合并

- **Status:** Done
- **Objective:** 将多个已完成甄选的 ICT 项目按业务数据汇总为一份《ICT项目甄选结果签批表（50万以下）》，而不是拼接多个 Word 文件。

## Progress

1. [x] 在签批表专属配置中增加“当前项目 / 多项目合并”模式，可选择、排序项目并编辑默认批次名称。
2. [x] 候选项目必须同时存在已保存的甄选前与甄选后方案；A 表限价优先取手填甄选限价，否则取甄选前 IT 投入，B-E 表取甄选后数据。
3. [x] 立项金额按不含税口径重算为“全部 IT 投入 + CT 专线建设/维护/带宽”；“专线/其他产品续签成本”要求用户逐项目确认后决定是否计入。
4. [x] 5 张明细表保持统一项目顺序并用 Decimal 重算合计；项目收入合计与项目投入合计各自独立，禁止复用错误总数。
5. [x] 阻断中选合作伙伴、甄选方式、甄选规则和收付款方式冲突；甄选范围、行业/场景、标准方案差异经用户确认后可由批次字段覆盖。
6. [x] 生成前执行项目级财务 0 容差、效益指标完整性和批次立项金额 `< 500000` 校验。
7. [x] 新增前端汇总规则测试与后端 DOCX 五表增行回归；完成 3 页样张渲染检查。

## Validation

- `npm run test:selection-batch --prefix src-ui`：通过。
- `npm run test:tax-split --prefix src-ui`：通过。
- `npm run build --prefix src-ui`：通过。
- `cargo test selection_result_docx_fills_batch_rows_and_approval_amount -- --nocapture`：通过。
- `npm run lint --prefix src-ui`：本次文件无新增错误；仓库仍有既有 `useAiContextStore.ts` `no-this-alias` 错误及历史 warnings。

## Scope Boundary

- 合并只读取已保存方案和模板资料，不写入其他项目的核心数据，也不改变财务公式、现金流、NPV 或税额规则。
- 不修改原 Word 模板结构；继续复用 `TABLE_*` 行克隆和变量替换引擎。
- `>= 50 万元` 的批次不允许使用该模板，不自动切换到其他审批模板。


# 甄选结果签批表 · 常用资料字段接入

- **Status:** Done
- **Objective:** 为《ICT项目甄选结果签批表》补齐与其他模板一致的“常用 / 存为常用”字段操作。

## Progress

1. [x] 复用 `CommonPresetFieldHeader` 和既有常用资料服务；未新增第二套弹窗、持久化或表单保存路径。
2. [x] 项目背景、收入侧收款方式、支出侧付款方式接入既有稳定 FieldKey。
3. [x] 新增甄选结果专属 FieldKey：中选合作伙伴、甄选内容说明、甄选范围、行业/场景、甄选方式、甄选规则、标准方案说明；每项字段绑定自身表单状态。
4. [x] `npm run build --prefix src-ui` 通过；`npm run lint --prefix src-ui` 仅保留既有的 `useAiContextStore.ts` `no-this-alias` 错误及项目历史 warnings，本次未新增 lint 问题。
5. [x] 发布前脱敏：项目索引中的本机绝对路径已改为仓库相对链接；`.claude/` 本地工具目录已忽略，避免未来误提交本机配置。

## Scope Boundary

- 未修改模板变量、文档生成、项目数据保存、财务计算、甄选限价或 0 容差校验。
- 常用内容只在用户点击替换或保存时通过既有组件生效；不会自动写入项目核心数据。
- 发布前须扫描已跟踪文件中的本机绝对路径与个人标识；项目文档只使用仓库相对链接。


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


# 立项决策汇报 PPT · 拆分科目展示方式

- **Status:** Done
- **Objective:** 允许用户在 PPT 投资收益页选择按科目合并展示或按测算拆分子笔展开，同一科目可生成多行且汇总口径不变。

## Progress

1. [x] 仅在当前 PPT 明细范围检测到有效拆分科目时展示选择，默认保持“合并展示”。
2. [x] 选择通过既有模板表单状态按项目、模板持久化，不写入或修改测算核心数据。
3. [x] 新增共享 PPT 拆分行规则；展开前重新校验子笔，损坏数据回退为汇总行。
4. [x] IT/CT 收入与投入四组明细均支持同科目多行，并以备注标注拆分笔次。
5. [x] 汇总、小计、现金流与效益指标继续使用原科目聚合值，不受展示选择影响。
6. [x] PPT 动态明细在模板原行高预算内等分；成本表新增行时不缩小字号，只下移评分标题及合同期限/动态回收期说明，底部指标保持原位。
7. [x] 税额专项测试、前端生产构建、真实 PPTX 行克隆测试与两页渲染检查通过。

## Scope Boundary

- 只处理现有投资收益页中的 IT收入、CT收入、IT投入、CT投入；非IT/CT与综合类科目不在当前 PPT 四张明细表范围。
- 本功能是文档展示偏好，不改变拆分状态、税额、资金计划、NPV 或其他财务计算。
