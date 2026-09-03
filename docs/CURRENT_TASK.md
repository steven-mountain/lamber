# 闭环 B 收尾：前端接通 / 审批持久化 / 挂起槽位清理

- **Status:** 2 项完成并自动验证；第 1 项已补齐代码与契约用例，**真实鼠标点击未由 AI 验证**（本机未授予 osascript 辅助功能与屏幕录制权限），需人工按 README 步骤确认。
- **模块文档:** [agent-bridge/README.md](../agent-bridge/README.md)

## 1. 前端弹窗接通

- [x] **发现真实缺口**：上一轮虽有 `AgentApprovalDialog.tsx` 并已挂载，但前端**无任何地方调用 `ai_send_prompt`**（`AiChatPanel` 仍走旧 `AiRuntime`），真实应用里根本无法触发 Agent，弹窗永远不出现。
- [x] 新增 `AgentLabView.tsx`（路由 `#/agent-lab`）作为触发入口：发指令、看会话事件、看审批审计日志。
- [x] `LAMBER_AGENT_LAB=1` 启动后自动跳转该路由（窗口无地址栏，否则不可达）；正常启动不受影响。
- [x] 弹窗在主界面与联调台两条渲染路径上都挂载。
- [x] 契约用例 2 个：审批事件的 6 个 camelCase 字段与弹窗读取的名字一致；弹窗挂载点与 `ai_send_prompt`/`ai_list_approval_log` 调用存在。
- [x] 应用能真实启动并停在联调台路由，进程稳定、stderr 无报错。
- [ ] **真实鼠标点击确认/拒绝两条路径** —— 未完成，需人工执行（README「人工验证步骤」）。

## 2. 审批决定持久化

- [x] 工作区 SQLite 新增 `agent_approval_log` 表；schema v9 → v10（纯建表，无数据迁移）。
- [x] `decided_by` 区分 `user` / `timeout` / `shutdown` / `internal`。
- [x] `ApprovalGate` 只持 `ApprovalRecorder` 回调，不碰数据库；生产由 `workspace_recorder` 注入，连接在决定时才解析。
- [x] 审计写失败只警告不改变决定。
- [x] Tauri 命令 `ai_list_approval_log(limit)`。
- [x] 用例：`approval_decisions_survive_a_process_restart`（丢弃全部内存持有者后重开磁盘库仍可查）、`audit_log_distinguishes_rejection_from_timeout`、`an_audit_write_failure_does_not_alter_the_decision`、`v9_database_gains_agent_approval_log_and_preserves_rows`。

## 3. 挂起槽位清理

- [x] `gate.shutdown()`：拒绝所有挂起审批并唤醒；`reopen()` 供下次启动复用。
- [x] `ai_agent_stop` 与 `main.rs` 的 `RunEvent::Exit` 钩子都会排空。
- [x] 关闭后新请求当场拒绝、不建槽位。
- [x] 崩溃路径：套接字随进程消失 → answerer `fetch` 报错 → `rejected`；最坏有 180 秒兜底。**未做优雅通知**（崩溃时没有机会通知），靠断连与超时兜底，符合失败关闭原则。
- [x] 用例：`shutdown_denies_parked_approvals_immediately`（网关超时设 300 秒，只有关闭能救）、`approvals_after_shutdown_are_denied_without_parking`、`answerer_fails_closed_when_the_bridge_dies_mid_request`。

## Validation

- `cargo test`：50 passed（新增 7 个默认用例 + 2 个迁移用例），无回归。
- `cargo test agent_bridge -- --ignored`（带真实 key）：9 passed，含闭环 A 两个回归检查点与闭环 B 完整链路。
- `npm run build --prefix src-ui` 通过；lint 新文件零告警。

## Scope Boundary

- 未改动已验证的审批核心机制：超时分层（90s / 180s）、失败关闭、令牌鉴权。
- 未做精美 UI；联调台与弹窗都是能演示流程的最简实现。
- 仍未加任何真实写操作工具。

---

# Agent 人工审批通道（deepseek-harness）· 闭环 B

- **Status:** Done（审批链路全程已实测；仅「模型真正发出被拦工具的 tool_call」这一步待有 API key 时验证）
- **Objective:** 打通 `dsh → answerer 插件 → Rust → React 弹窗 → 用户确认 → dsh`，让写操作类工具必须经用户确认才执行。
- **模块文档:** [agent-bridge/README.md](../agent-bridge/README.md)

## Progress（闭环 B，本次完成）

1. [x] 读源码确认审批协议（`@deepseek-ai/dsh-user-approval@0.1.2-alpha.5`）：事件字段、`ApprovalOutcome` 四值、无 answerer 时的失败关闭行为、`tools/pre-execute` 瀑布终止默认值为 `allow`。
2. [x] `write_test_marker(note?)` 无害测试工具：只写系统临时目录，不碰 lamber 任何数据。
3. [x] 审批守卫 + 答复器（`approval.ts`）与参数关联表（`pendingCalls.ts`）。
4. [x] Rust `approval.rs`：`POST /lamber-bridge/approval` + `ApprovalGate` 挂起/唤醒/超时/槽位回收，复用既有令牌鉴权，未开无鉴权端口。
5. [x] Tauri 命令 `ai_resolve_approval`、前端事件 `ai://approval-request`。
6. [x] `AgentApprovalDialog.tsx` 最简弹窗，挂 App 根节点。
7. [x] `bridge_server.rs` 改每请求一线程（上限 16）——修复闭环 A 遗留的单线程 accept 循环缺陷。
8. [x] 分层验证 6 个默认用例 + 8 个 `#[ignore]` 用例。

## Validation

- `cargo test`：42 passed，无回归。
- `cargo test agent_bridge -- --ignored`：8 passed，含闭环 A 两个回归检查点。
- `only_the_write_tool_is_gated_behind_approval`：`run_benefit_calculation=false` / `write_test_marker=true`，只读工具未被误伤。
- `approval_times_out_as_an_explicit_rejection`：超时 1 秒内返回明确拒绝，不挂死，槽位已回收。
- `a_parked_approval_does_not_block_the_calculation_route`：挂起的审批不阻塞其它路由。
- `npm run build --prefix src-ui` 通过；lint 新文件零告警。
- 完整闭环用例需 `DEEPSEEK_API_KEY`，本机未设置，按设计跳过。

## Scope Boundary

- 未加任何真实写操作工具（硬约束）；`write_test_marker` 只写 `os.tmpdir()`。
- 未改动 `calculator.rs` / `docfill.rs` / 测算引擎 / `AiRuntime.ts` / `AiChatPanel.tsx`。
- 未做 SEA 打包，未做精美 UI。

## Next

- 把真实写操作工具逐个加进 `GATED_TOOLS`，并补桥接路由与端到端用例。
- 对话流里可视化工具调用过程（订阅 `ai://session-event`）。
- 审批弹窗产品化（批量授权、历史记录、按 DESIGN.md 精修）。
- API key 改由前端凭证传入；SEA 打包瘦身。

---

# Agent 工具执行能力接入（deepseek-harness）· 闭环 A

- **Status:** Done（Rust → dsh → 工具 → calculator 全链路已实测；仅「模型真正发出 tool_call」这一步待有 API key 时验证）
- **Objective:** 让 lamber 的 AI 顾问具备真正的工具执行能力。由 dsh 承担 agent loop / 工具编排 / 审批，lamber 保留全部 Rust 业务逻辑，两者以本地回环 HTTP 桥接相连。
- **模块文档:** [agent-bridge/README.md](../agent-bridge/README.md)

## Progress（闭环 A，本次完成）

1. [x] `agent-bridge/dsh-tool-lamber/`：dsh 自定义工具插件包，`defineTool` 注册 `run_benefit_calculation(projectId, scenario?)`，`npx tsc` 零报错。
2. [x] `agent-bridge/patch.yml` + `scripts/provision-profile.mjs`：插件挂载进 `sdk` profile（必须先 `dsh plugin add` 链接进 `$DSH_HOME/profiles/sdk/`，只写 `--patch` 加载不到）。
3. [x] `bridge_server.rs`：仅监听 `127.0.0.1`、临时端口的 tiny_http 桥接服务，带一次性令牌鉴权（定长比较）与请求体大小上限。
4. [x] `calculation.rs`：`POST /lamber-bridge/calculate` → 项目 → 方案（`pre_selection` / `post_selection` / 方案 id / 方案名，缺省用 `default_scheme_id`）→ 最新快照 → `calculate_ict_benefit`。严格只读，未命中即报错不静默回退。
5. [x] `dsh_session.rs`：dsh 子进程管理 + 手写 newline-delimited JSON-RPC 2.0 客户端（id 匹配请求/响应，无 id 即通知；读线程 + 条件变量唤醒，子进程退出时唤醒阻塞调用方）。
6. [x] `mod.rs`：`AgentRuntime` 懒启动与生命周期，Tauri 命令 `ai_send_prompt` / `ai_agent_status` / `ai_agent_stop`，通知按 `{method, params}` 统一 emit 到 `ai://session-event`。
7. [x] `main.rs` 注册模块、命令与 `AgentRuntime` 状态。
8. [x] `Cargo.toml` 新增 `tiny_http 0.12`（`default-features = false`）。
9. [x] 分层测试：6 个默认用例（无需 Node/网络）+ 3 个 `#[ignore]` 集成用例。
10. [x] `.gitignore` 排除 `agent-bridge/.dsh-home/` 与插件构建产物。

## Validation

- `cargo test`：36 passed，既有 29 个用例无回归。
- `cargo test agent_bridge -- --ignored`：
  - `plugin_tool_body_reaches_the_calculator_over_the_bridge` 通过——插件 `execute()` 经桥接拿回引擎直算一致的 NPV / 利润率。
  - `dsh_advertises_the_lamber_tool_in_its_request_header` 通过——真实 dsh 子进程启动、`initialize` 成功、`request/header` 中出现 `run_benefit_calculation`。
  - `dsh_tool_call_reaches_the_calculator_and_returns_real_numbers` 需 `DEEPSEEK_API_KEY`，本机未设置，按设计跳过。
- `npx tsc`（插件包）：零报错。

## Scope Boundary

- 未改动 `calculator.rs` / `docfill.rs` / 测算引擎 / NPV / 现金流 / 税额 / 甄选费 / 反算 / 0 容差校验。
- 未改动 `AiRuntime.ts` / `AiChatPanel.tsx`——本轮验收标准是 Rust 侧能拿到通知，前端展示不在范围内。
- 插件包内只有 `run_benefit_calculation` 一个工具，且为只读。

## Next（闭环 B 及后续）

- **审批通道**：dsh 的 `approval/request` 是进程内 Cordis 事件，**不会**经 SDK JSON-RPC 转发给外部客户端。需在 dsh 侧新写一个 answerer 插件（建议 `agent-bridge/dsh-answerer-lamber/`，`patch.yml` 再加一条 `insert`）把请求转给 Rust → React 弹窗 → 用户确认 → 回传 dsh。**在此之前不要往插件里加任何会写数据的工具。**
- 前端 `AiChatPanel.tsx` 订阅 `ai://session-event` 展示工具调用过程。
- API key 改由前端凭证传入（现从环境变量读；前端的 key 在 `localStorage.lamber_ai_api_key`）。
- 单文件 SEA 打包 / 依赖瘦身（现依赖开发机上的 `agent-bridge/node_modules`）。

---

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
