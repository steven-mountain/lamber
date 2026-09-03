# CHANGELOG_AI.md

> [!NOTE]
> **历史兼容性说明**：本文件记录 AI 代理所做的结构修改、业务规则和上下文更改，不再作为 AI 每次任务的默认必读文件。
> 只有在追溯历史回归或用户明确要求时，才需要加载和查看此变更日志。

This changelog records structural modifications, business rules, and context changes made by AI agents to maintain a reliable project state mapping.

## 2026-09-03（三）

### 闭环 B 收尾：前端接通 / 审批持久化 / 挂起槽位清理

Added:
- `src-ui/src/components/ai/AgentLabView.tsx` + 路由 `#/agent-lab`：真实应用里驱动 Agent 与审批的入口（发指令 / 会话事件流 / 审批审计日志）。`LAMBER_AGENT_LAB=1` 启动自动跳转。
- `src-tauri/src/agent_bridge/approval_log.rs`：审批审计的 SQLite 落库与查询；Tauri 命令 `ai_list_approval_log`。
- `agent_approval_log` 表（schema v9 → v10，纯建表无数据迁移），`decided_by` 区分 user/timeout/shutdown/internal。
- `ApprovalGate::shutdown()` / `reopen()` / `set_recorder()`；`main.rs` 增加 `RunEvent::Exit` 排空钩子。

Fixed:
- **真实缺口**：前端此前无任何地方调用 `ai_send_prompt`（`AiChatPanel` 仍走旧 `AiRuntime`），导致审批弹窗在真实应用里永远无法出现。补 `AgentLabView` 作为触发入口。

Decisions:
- **网关不碰数据库**，只持 `ApprovalRecorder` 回调；连接在决定时解析而非启动时捕获，使 Agent 起来后才打开的工作区也能记账。
- **审计写失败只警告不改变决定**：让存储故障阻塞 Agent 是错的，失败放行更错。
- **崩溃路径不做优雅通知**：lamber 崩溃时没有机会通知 dsh，靠套接字断开 + 180 秒兜底，与既有失败关闭原则一致。
- **联调台路由靠环境变量开启**：窗口无地址栏，否则不可达；正常启动完全不受影响。

Validation:
- `cargo test`：50 passed，无回归。
- `cargo test agent_bridge -- --ignored`（带真实 key）：9 passed。
- 新增用例覆盖：跨进程重启后审批记录仍可查、拒绝与超时可区分、审计写失败不改变决定、关闭立即释放挂起审批、关闭后不建槽位、桥接中途断开时 answerer 失败关闭、前后端字段契约、弹窗挂载点。

Known limitation:
- **真实鼠标点击未由 AI 验证**：本机未授予 `osascript` 辅助功能权限与 `screencapture` 屏幕录制权限，无法发出真实点击也无法截图。已用两个契约用例覆盖"点击本来能发现的那类问题"（跨进程名字漂移），但弹窗视觉呈现与可点性仍需人工确认，步骤见 `agent-bridge/README.md`。
- 审计表位于工作区数据库；审批发生时若未打开工作区，该条记录会丢（只留 stderr 警告）。

## 2026-09-03（二）

### Agent 人工审批通道（deepseek-harness / 闭环 B）

Added:
- `dsh-tool-lamber/src/writeTestMarker.ts`：无害测试工具 `write_test_marker(note?)`，只往 `os.tmpdir()` 下新建的临时目录写一个带时间戳的文本文件，**不碰 lamber 工作区/数据库/项目文件**。
- `dsh-tool-lamber/src/approval.ts`：审批守卫（`tools/pre-execute` 对被拦工具返回 `{kind:'ask'}`）+ 答复器（`approval/request` 转发给 Rust，阻塞等待）。导出 `isGatedTool` 作为拦截策略的唯一真相源。
- `dsh-tool-lamber/src/pendingCalls.ts`：按 `callId` 关联被拦调用的参数（有上限、自清理）。
- `agent-bridge/scripts/check-approval.mjs` / `check-gating.mjs`：分别单跑「答复器→桥接→网关→回传」和「拦截策略」两跳的排错脚本。
- `src-tauri/src/agent_bridge/approval.rs`：`POST /lamber-bridge/approval` 路由与 `ApprovalGate`（条件变量挂起/唤醒、超时、槽位回收）。
- Tauri 命令 `ai_resolve_approval(requestId, approved)`；前端事件 `ai://approval-request`。
- `src-ui/src/components/ai/AgentApprovalDialog.tsx`：最简审批弹窗，挂在 App 根节点。

Changed:
- `bridge_server.rs` 由单线程 accept 循环改为**每请求一个线程**（上限 16）。
- `workspace_handler` 增加 `gate` 与 `announce` 两个入参；`AgentRuntime` 持有跨 dsh 启动周期的 `ApprovalGate`。

Decisions:
- **守卫与答复器和工具同包，不拆成 `dsh-answerer-lamber/`**（推翻上一轮 README 里的预留）。根因：`ApprovalRequestEvent` 不带工具参数，必须靠进程内 map 按 `callId` 关联；两个 npm 包在 pnpm link 下可能拿到两份模块实例、两张表。源码上仍分文件保持接缝清晰。
- **全程失败关闭**：桥接报错、返回体异常、超时、锁中毒一律 `'rejected'`，不挂起也不默认放行。
- **超时分层**：Rust 网关 90 秒（`LAMBER_APPROVAL_TIMEOUT_SECS` 可覆盖），答复器 180 秒。答复器故意更长，让正常路径是网关给出明确 `rejected`。
- **网关超时时长做成 `ApprovalGate` 字段而非 wait 时读环境变量**：否则并行测试改同一个进程全局变量会互相干扰。
- **审批弹窗挂在 App 根节点而非 AI 面板内**：后端挂着一个 dsh 工具调用等这个答复，监听器随面板关闭而消失会把每次审批都变成超时。
- **仍不加任何真实写操作工具**，符合本轮硬约束与 AGENTS.md。

Validation:
- `cargo test`：42 passed（新增 6 个默认审批用例），既有用例无回归。
- `cargo test agent_bridge -- --ignored`：8 passed。含闭环 A 的两个回归检查点（`dsh_advertises_the_lamber_tool_in_its_request_header`、`plugin_tool_body_reaches_the_calculator_over_the_bridge`）仍通过。
- `only_the_write_tool_is_gated_behind_approval` 直接断言 `run_benefit_calculation=false` / `write_test_marker=true`，证明只读工具没被审批误伤。
- `npm run build --prefix src-ui` 通过；`npm run lint` 新文件零告警（仅既有 `useAiContextStore.ts` 的 `no-this-alias` 报错）。
- 带 `DEEPSEEK_API_KEY` 的完整闭环用例 `dsh_gated_tool_runs_only_after_the_user_confirms` 已写好（确认/拒绝两路都测），本机无 key，按设计跳过。

读源码确认的协议事实（与任务书预设有出入的部分，详见 `agent-bridge/README.md`）：
- `defineTool` **没有**"审批等级"字段，只能靠 `tools/pre-execute` 返回 `{kind:'ask'}`；瀑布终止默认值是 `allow`，所以未点名的工具原样放行。
- `ApprovalRequestEvent` 只有 `{agent, toolName, callId?, reason?, signal?}`，**不含参数**。
- `ApprovalOutcome` 四值：`allowed-once` / `rejected` / `cancelled` / `unavailable`，只有 `allowed-once` 放行且仅此一次。
- 无 answerer 时两种默认都是失败关闭；会话策略 `never` 时根本不问人，直接 `rejected`。
- `approval/asked` / `approval/decided` 是 session 事件，会流到 Rust 侧，适合做断言。

## 2026-09-03

### Agent 工具执行能力接入（deepseek-harness / 闭环 A）

Added:
- `agent-bridge/dsh-tool-lamber/`：独立 npm 包形态的 dsh 自定义工具插件，用 `defineTool` 注册唯一工具 `run_benefit_calculation(projectId, scenario?)`。工具体只做一件事——把参数 POST 给 lamber 的回环桥接服务，不含任何业务数学。`npx tsc` 零报错。
- `agent-bridge/patch.yml` + `scripts/provision-profile.mjs`：把插件挂进 dsh 的 `sdk` profile。dsh 解析插件包名是相对 `$DSH_HOME/profiles/<profile>/`，因此必须先 `dsh plugin add` 链接，只写 `--patch` 不生效。
- `agent-bridge/scripts/check-bridge.mjs`：只跑「插件工具体 → 桥接 → calculator」一跳的排错脚本，不启 dsh、不消耗 LLM。
- `src-tauri/src/agent_bridge/bridge_server.rs`：仅监听 `127.0.0.1`、临时端口的 HTTP 桥接服务（tiny_http，独立线程，阻塞式）。
- `src-tauri/src/agent_bridge/calculation.rs`：`POST /lamber-bridge/calculate` 路由，项目 → 方案 → 最新快照 → `benefit::calculator::calculate_ict_benefit`。严格只读。
- `src-tauri/src/agent_bridge/dsh_session.rs`：dsh 子进程管理与手写 JSON-RPC 2.0（newline-delimited over stdio）客户端。
- `src-tauri/src/agent_bridge/mod.rs`：`AgentRuntime` 生命周期 + Tauri 命令 `ai_send_prompt` / `ai_agent_status` / `ai_agent_stop`，通知统一按 `{method, params}` emit 到前端事件 `ai://session-event`。
- 新增直接依赖 `tiny_http 0.12`（`default-features = false`，4 个小传递依赖，无 TLS）。

Decisions:
- **业务逻辑边界**：dsh 只负责 agent loop / 工具编排 / 审批，`calculator.rs`、`docfill.rs` 等一律不动。桥接是唯一接缝，本轮只开一条只读路由。
- **不引入官方 TS/Python SDK 客户端**：协议是一行一个 JSON 对象的 JSON-RPC 2.0，Rust 侧手写读写即可，避免为拼 JSON 而嵌入一个 Node SDK。
- **选 tiny_http 而非 hyper/axum**：lamber 后端整体同步（`rusqlite` + `Arc<Mutex<Connection>>`，`calculate_ict_benefit` 是同步函数），在异步 handler 里同步加锁是反模式；桥接只有一个回环路由、并发极低。也不手写 HTTP 解析——那比引入一个小依赖债更大。
- **桥接令牌鉴权**：只绑回环不等于鉴权，同机任意进程都能读到客户项目财务数据。每次启动随机生成令牌经环境变量交给子进程，请求头 `x-lamber-bridge-token` 定长比较校验，不落盘。
- **遥测强制关闭**：`sdk` profile 默认把遥测发往外部主机，Rust 侧硬编码 `DSH_TELEMETRY_MODE=DISABLED`。
- **`scenario` 未命中即报错，不静默回退**：报错好过让 agent 引用错方案的财务数字。选择器支持 `pre_selection` / `post_selection` / 方案 id / 方案名，缺省用项目 `default_scheme_id`。
- **懒启动**：dsh 子进程在首次 prompt 时才拉起，不占用未使用 AI 面板用户的启动时间。
- **本轮不加任何会写数据的工具**，等审批通道（闭环 B）就绪后再说，避免绕过 AGENTS.md 的用户确认要求。

Validation:
- `cargo test`：36 passed（新增 6 个默认用例：默认方案 / 按阶段 / 按方案名与 id 选择、未知 scenario 拒绝、令牌鉴权、桥接端到端返回引擎数字），既有 29 个用例无回归。
- `cargo test agent_bridge -- --ignored`：`plugin_tool_body_reaches_the_calculator_over_the_bridge` 与 `dsh_advertises_the_lamber_tool_in_its_request_header` 均通过——已实测 dsh 真实子进程启动、`initialize` 握手成功、`request/header` 通知中出现 `run_benefit_calculation` schema。
- 带 `DEEPSEEK_API_KEY` 的完整闭环用例已写好，本机无 key，运行时按设计跳过。

与任务书预设的差异（详见 `agent-bridge/README.md`）：包版本为 `0.1.2-alpha.5`；provider 是 `deepseek-official` 而非 `deepseek`，模型 `deepseek-v4-flash`；`output.schema` 的对象属性除 `additionalProperties` 外还需逐个 `required: true`；`DEEPSEEK_API_KEY` 环境变量确实生效；`pnpm` 无需全局安装。

Not done (下一轮):
- 闭环 B 审批通道（需 dsh 侧 answerer 插件转发进程内 Cordis 事件）、前端 `AiChatPanel.tsx` 展示、API key 改由前端凭证传入、SEA 打包瘦身。

## 2026-07-01

### ICT 测算表内"甄选前 / 甄选后"方案切换（第 2 阶段）

Modified:
- `IctLifecycle.tsx` 顶部横幅由静态"当前方案"文本改为二段式切换控件 `[甄选前][甄选后] 更多方案 ▾`：
  - 组件内新增 `schemes` 状态（在 `loadProjectContext` 中随项目一并加载/重置），按 `updated_at` 倒序推导每阶段主方案 `preScheme` / `postScheme`；未标注方案及同阶段历史方案统一进入"更多方案"下拉。
  - 点击已存在阶段按钮：先走 `confirmOrSave()` 未保存确认，再 `navigateTo("ict_lifecycle", projectId, schemeId)` 复用既有加载路径切换方案（不新增并行状态）。
  - 点击"甄选后（未生成）"：当前方案未打标签时就地 `update_scheme_stage`；当前方案已属另一阶段时复用"另存为新方案"弹窗派生新方案（默认名 `${项目名}_甄选后`，`stage=post_selection`，以当前测算数据为起点）。
  - `domainSaveService.saveBenefitAnalysis` / `handleSaveAsNew` 透传 `stage`，另存为弹窗支持"派生阶段方案"标题与说明。
- 文档生成安全提示：`handleTabSwitch` 中，点击模板名含"签批"且 `activeScheme.stage !== "post_selection"` 时给出非阻断 `confirm`（继续 / 取消），不强制拦截，兼容甄选前预生成草稿。
- 顺带修复 `IctLifecycle.tsx` 两处历史全角空格（U+3000）触发的 `no-irregular-whitespace` lint 报错（改为 `{"　"}` 表达式，渲染不变）。

Decisions:
- 方案切换复用项目下拉/返回按钮相同的 `confirmOrSave` + `navigateTo` 路径，避免引入第二套 activeScheme 加载逻辑。
- "甄选后未生成"分两种语义：当前方案未标注 → 就地打标（对应"首次测算生成甄选前方案"）；当前方案已标注 → 派生新方案（对应"基于甄选前创建甄选后"），避免误建重复方案。
- 未改动测算引擎、快照结构/版本号、Excel/Word 模板与坐标映射。

Fixed:
- 派生"甄选后"方案后点击"甄选前"无法切换（需退出重进才生效）：根因是 `activeScheme`（本地态）与导航 store `activeSchemeId` 两个真相源漂移——保存/派生走 `loadProjectContext` 直接加载不刷新 store，导致 store 陈旧地仍指向目标方案，`navigateTo` 写入相同值不触发 `[activeProjectId, activeSchemeId]` 加载 effect。`switchToScheme` 改为：始终 `navigateTo` 同步 store，并在 store 已指向目标 id 时直接 `loadProjectContext` 兜底，保证切换必然生效。
- 甄选前/甄选后切换控件位置随方案名长短漂移：把阶段切换控件移到当前方案行的固定最左侧（`shrink-0`），方案名放其右侧并 `truncate`，单行不换行，位置稳定。
- 切换方案时顶部项目卡片（红框）跳动/闪一下：根因是 `loadProjectContext` 开头同步 `setActiveProject(null)` 等清空，异步取数期间 `activeProject === null` 令卡片瞬间切成"自由测算模式"布局再切回。改为**同项目内切换只保留项目级/方案级展示状态、不整块清空**（`isSameProject` 判定，比较导航目标与 `useProjectStore` 当前项目），仅跨项目时才清空；同项目下这些状态在数据就绪后从旧值平滑更新到新值。
- **数据串档（严重，架构级修复）**：修改"甄选后"金额后切到"甄选前"，甄选前数据也被改（反之亦然）。
  - 根因：测算"工作副本"存放在 `project_lifecycle_states` / `project_cashflow_states` 两张表，且是 `project_id UNIQUE` 的**项目级单例**（每项目仅一行）；同时科目金额编辑标记的是 `cashflow` scope，保存时只写这份共享副本、不落各方案快照，加载又优先读它。于是甄选前/甄选后自始至终读写同一行数据。
  - 修复：把工作副本改为**按方案存储**。
    - 后端 schema 迁移 v8 → v9：两表新增 `scheme_id` 列，唯一键由 `project_id` 改为 `(project_id, scheme_id)`（SQLite 建新表→拷贝→删旧→改名重建）；既有行归属到项目 `default_scheme_id`（无默认方案则归入 `''` 桶）。
    - `save_lifecycle_state` / `save_cashflow_state` / `get_lifecycle_state` / `get_cashflow_state` 命令新增 `scheme_id` 入参，按 `(project_id, scheme_id)` upsert / 读取；智算导入路径写入项目默认方案桶（`resolve_default_scheme_bucket`）；`get_project_full_state` 返回默认方案桶（供智算视图/文档导出等按项目取的消费方沿用）。
    - 前端：`domainSaveService` 的四个状态读写方法新增 `schemeId`；`IctLifecycle` 的 lifecycle/cashflow 保存 handler 与 `persistLifecycleAndCashflowState` 均带上当前 `activeScheme.id`；`loadProjectContext` 改为按选中方案精确加载**该方案自己的草稿**（`loadLifecycleState/loadCashflowState(project, schemeId)`），无草稿再回落到该方案快照，无选中方案才回退到项目默认桶/legacy。
  - 效果：甄选前/甄选后是两份完全独立的工作副本，编辑与切换互不影响。此修复取代当日早前基于 `default_scheme_id` 的临时判定（`preferCurrentState` 门控），后者因金额编辑只落 `cashflow` scope、既不建快照也不改 default 而无法成立。
  - 迁移测试：`v8_lifecycle_cashflow_states_gain_scheme_id_and_allow_multiple_per_project`（含既有行回填、去除 project_id 唯一约束、(project_id, scheme_id) 仍唯一）；`fresh_database_uses_schema_v9_*` 校验全新库落到 v9 且两表含 `scheme_id`。

### 测算方案甄选阶段标签（甄选前 / 甄选后）

Modified:
- `BenefitAnalysisScheme` 新增可选 `stage` 字段（`pre_selection` / `post_selection` / `None`），随 serde 默认值兼容历史数据。
- 数据库 `benefit_schemes` 表新增 `stage TEXT` 列；`schema_version` 迁移 7 → 8（幂等 `ALTER TABLE ... ADD COLUMN`，既有行默认未标注）。
- `save_benefit_analysis` 命令新增 `stage` 入参：新建方案按传入 stage 写入；更新既有方案用 `COALESCE(?, stage)` 保留原标签（普通保存不清空）。
- 新增 `update_scheme_stage(project_id, scheme_id, stage)` 命令：独立改写方案阶段标签，不产生新的效益快照；空/非法值归一化为未标注。
- 读取路径（`get_benefit_schemes`、`get_project_full_state`）与 JSON/SQLite 仓储、JSON→SQLite 迁移全部同步读写 `stage`。
- 前端：`ProjectBoard` 方案 chip 展示阶段标签，并在快照面板提供“甄选前/甄选后”分段按钮设置/取消标注；`IctLifecycle` 当前方案横幅展示阶段 chip；新增 `lib/schemeStage.ts` 统一标签、配色与归一化。

Decisions:
- 阶段标签仅用于区分与展示两版效益分析（甄选前用预算/最高限价成本、甄选后用中标实际成本），不改动 `calculator.rs` 测算引擎、NPV、现金流、税额、科目金额或 0 容差校验。
- 甄选前/甄选后作为**同一项目下的不同方案**并列存在，靠 stage 打标而非新建项目，便于对比与后续《甄选结果签批表》取数（第 2 阶段）。

## 2026-06-17

### 采购甄选费 Excel 回填

Modified:
- 甄选测算面板新增可选 `IctInput` 元数据保存：供应商报价、上浮、实际测算成本、甄选费、最高限价和固定锚点。
- 生成《效益分析表》时，将供应商报价回填到 `3-直接经济效益评估表!T26`，将甄选上浮回填到 `3-直接经济效益评估表!T29`。
- 手动/后台导入生命周期效益分析表时回读 `T26/T29`，保存到同一甄选元数据字段，支持再次生成时保留。
- 甄选面板字段变化会标记效益方案 dirty，避免仅修改报价/上浮后退出时没有保存提示。

Decisions:
- 新字段只服务模板回填、状态恢复和 AI 只读上下文，不参与 NPV、现金流、税额、科目金额或 0 容差校验。
- 只有面板已有报价、限价或测算结果时才输出 T26/T29，避免默认上浮值在未使用甄选测算时污染空模板。

### 需求导入表 Tab 过度拆分优化

Modified:
- 《ICT项目需求导入表》共享模板骨架的导航从“需求基础信息 / 需求内容 / 商务与预算 / 实施与风险 / 生成确认”收敛为“需求信息 / 生成确认”。
- 原需求单位、客户确认、服务内容、设备清单、技术方案清单、业务模式、公示/招标材料、部署环境、信息安全和客户确认截图合并到“需求信息”页，并使用低对比子模块分组承载。
- 模板切换时需求导入表默认进入“需求信息”页。

Decisions:
- 本次只调整前端展示组织，不修改 `gen_*` 字段、图片上传/粘贴、AI 图片分析、常用内容入口、保存 payload、模板变量、文档生成逻辑或后端接口。
- 内容较少的模板不应机械拆成多个 Tab；需求导入表与立项签批表保持两段式结构，正文内容优先在单页内分组扫描。

### 立项签批表 Tab 过度拆分优化

Modified:
- 《ICT项目立项签批表》共享模板骨架的导航从“项目基础信息 / 服务内容 / 商务条款 / 立项与采购 / 生成确认”收敛为“签批信息 / 生成确认”。
- 原项目背景、IT/CT 服务内容、收付款方式、垫资和立项后甄选选项合并到“签批信息”页，并使用低对比子模块分组承载。
- 模板切换时立项签批表默认进入“签批信息”页。

Decisions:
- 本次只调整前端展示组织，不修改 `gen_*` 字段、常用内容入口、保存 payload、模板变量、文档生成逻辑或后端接口。
- 内容较少的模板不应机械拆成多个 Tab；保留“生成确认”独立页，正文内容优先在单页内分组扫描。

### AI 助手悬浮入口拖拽性能优化

Modified:
- `AiFloatingLauncher` 拖动定位从连续更新 `left/top` 改为外层 `translate3d(...)`。
- 拖动过程中使用 `requestAnimationFrame` 合并 DOM transform 更新，结束后再同步 React 状态和 `lamber_ai_launcher_position`。
- 将 hover/active 视觉动画移入内部视觉层，外层定位元素不再使用 `transition-all`，避免 transform 或位置变化被 CSS 过渡拖慢。

Decisions:
- 拖拽性能问题不是页面刷新导致，而是高频 React 状态更新叠加位置 CSS 过渡导致的交互延迟。
- 本次只优化 AI 入口按钮拖拽链路，不改变 AI 窗口、AI 上下文、项目数据或显示/隐藏偏好策略。

### AI 助手悬浮入口隐藏设置

Modified:
- 外观偏好新增 `aiLauncherVisible`，默认隐藏主窗口右下角 AI 助手悬浮入口，并通过 `lamber_appearance_settings` 持久化。
- 外观 store 对旧偏好和跨窗口同步 payload 统一归一化，缺失字段时补默认隐藏值。
- 主应用只在 `aiLauncherVisible` 为 `true` 时渲染 `AiFloatingLauncher`；设置中心新增“AI 助手入口”开关用于重新显示或隐藏。

Decisions:
- 该偏好只控制主窗口入口按钮，不改变 AI 对话窗口、AI 上下文构建、模板图片分析、项目测算或任何业务数据写入路径。
- 偏好继续属于应用 UI 设置，仅写入 localStorage，不写入 workspace SQLite 或项目文件。

## 2026-06-16

### 智算金额来源另存与导入导出

Created:
- `amountSourceExchange.ts`：智算金额来源交换包构建、导入归一化、同步状态净化和项目周期/折现率参数重建 helper。
- 后端金额来源交换包导出和读取 IPC，复用 Tauri 文件对话框，并在 Rust 侧校验 `kind/schemaVersion/source/projectSettings` 结构。

Modified:
- 智算页面新增“另存为来源”“导出当前来源”“导入来源”；另存使用当前内存画布创建新普通来源，不先覆盖原来源。
- 导出包仅包含当前来源的参数、公式、收入/成本项、资金计划、映射、计算快照摘要和项目周期/折现率，不携带原项目 ID、来源版本、H200 保护角色、同步 revision 或 ICT 正式结果。
- 导入包生成新的来源 ID 并设为当前同步来源，用户可在导入预览中确认是否使用文件内项目周期/折现率覆盖当前项目级参数。
- 自动保存定时器在执行时重新读取 store 状态，避免另存或导入期间的排队保存覆盖旧来源。
- ICT 同步明细弹窗使用固定列宽和横向滚动，给“来源项”“状态”列保留稳定宽度，避免状态列单字换行拉高行高。

Decision:
- 金额来源导入导出是来源级交换，不是完整项目迁移；导入后仍必须通过既有智算同步和 0 容差校验才能写入 ICT 正式测算。

Validation:
- 前端专项测试覆盖导出包净化、导入参数重建和覆盖/保留项目参数两类路径；Rust 测试覆盖非法 kind/schema、缺少核心数组字段和写出读回。

### 智算 ICT 同步冲突与 H200 来源删除保护

Modified:
- `sync_intelligent_compute_to_ict` 成功响应现在返回最新智算项目状态，前端同步成功后立即刷新 `syncRevision/stateVersion/controlledSubjects`，避免覆盖 ICT 后再次点击仍使用旧乐观锁基线并误报冲突。
- 金额来源切换、新建和删除会清空旧同步预览，后续同步明细必须从最新 ICT 状态重新生成。
- 金额来源管理弹窗新增来源列表，可将普通来源设为当前或删除，并显示当前来源与 H200 基准状态。
- 默认 H200 来源写入 `metadata.sourceRole = "h200_baseline"`；后端兼容旧数据中最早的 H200 默认描述来源并拒绝删除该基准，前端同步禁用删除入口。
- 新建来源默认以系统内置 H200 标准模板为基底，保留空白、当前来源和已有来源复制选项。

Decision:
- H200 标准基准可重命名和复制，但不能删除；不新增表或迁移，保护规则使用现有 `metadata_json` 与旧描述兼容。

Validation:
- 前端专项测试覆盖 H200 默认基底、删除规则和同步锁 revision 构建；Rust 测试覆盖最后来源、H200 基准和普通来源删除后的活动来源切换。

### 智算金额来源自动同步与税率口径修复

Modified:
- 智算金额来源编辑后会按业务指纹防抖保存并自动同步到 ICT 正式状态；手动同步和同步明细仍保留用于预览、排查和重试。
- ICT 标准科目目录增加 `defaultTaxRate`，智算导出、ICT 水合和资金计划 finalizer 统一使用该目录默认税率。
- 智算到 ICT 的同步 payload 同时写入含税金额、不含税金额和税率；当前同步来源内同一科目按含税/不含税分别汇总后反算有效税率。
- 金额来源管理改为弹窗入口，新建来源必须命名，并可从空白、H200 标准、当前来源或任一已有来源创建；新建后自动成为当前唯一 ICT 同步来源。
- 同步来源从“多个启用来源汇总”收敛为“当前选中来源完全覆盖 ICT”，后端保存金额来源时兜底关闭同项目其他来源的启用标记。
- 当前来源未产出的 ICT 标准收入/成本科目会写入金额 0、默认税率和 10 年全 0 资金计划；同步明细新增“智算无输出，写 0”状态。
- 并发版本冲突不使用旧预览强写，前端提供“重新加载冲突并覆盖”，重载最新状态后再执行智算完全覆盖 ICT。
- `source: intelligent_compute` 导入痕迹可作为释放旧受控科目的兜底依据，防止改映射或改来源后 ICT 残留旧金额。

Decision:
- 智算自动同步会写入 ICT 正式状态，并完全覆盖 ICT 标准收入/成本科目的金额、税率和资金计划；产权、现金流模型等非智算金额参数保持 ICT 当前状态。
- 同步金额口径固定为“元、含税”，万元仅用于展示。

Validation:
- 智算专项测试覆盖 H200 大额同步、空科目默认税率、单来源选择、全量科目归零和 `intelligent_compute` 痕迹释放。
- TypeScript 生产构建与 Rust 测试通过。

## 2026-06-13

### 智算参数类别与自由排序

Created:
- `parameterLayout.ts`：参数类别默认配置、旧蓝图迁移、类别/参数排序和删除规则纯函数。

Modified:
- 蓝图持久化升级到 Version 3，增加 `parameterGroups`，参数增加 `groupId` 与 `isKey`。
- 参数区支持新增、重命名、排序类别，以及参数跨类别移动、同类排序和类别内新增。
- 内置类别不可删除；自定义类别只有清空参数后才可删除。筛选或搜索时禁用拖拽，菜单操作保留为键盘和触屏备用路径。
- 移除固定参数说明侧栏和高级参数折叠，参数说明改为卡片内按需展开。
- 同步指纹忽略类别归属、排序和关键标记，布局调整不触发 ICT 正式计算。

Validation:
- 专项测试覆盖旧蓝图迁移、H200 参数归属、类别持久化往返、类别删除规则、参数跨类别与同类排序，以及同步指纹稳定性。
- 浏览器验证新增类别、类别内新增参数、菜单跨类别移动、类别菜单排序、卡片说明、搜索和 `1440px / 760px` 无横向溢出。

### 智算参数区视觉回调

Modified:
- 参数区调整为参考稿的分组折叠结构：规模、定价、投入默认展开，运营、财务显示单行摘要。
- 关键参数集调整为 12 个主要业务参数，其余参数由“更多参数 / 高级参数”承载。
- 参数卡移除多类型彩色左边框和常驻字段 key/操作区，只保留名称、值、单位、状态与影响提示。
- 字段 key、敏感性开关及复制/删除移入次级设置菜单；参数说明改为右侧窄卡。

Validation:
- 浏览器验证分组展开、字段 key 搜索及 `1440px / 760px` 页面无横向溢出。

### 智算测算工作台布局优化

Modified:
- `AiComputeQuoteView.tsx`：重组为 compact header、四 Tab 主编辑画布和可收起效益结论抽屉。
- 参数按规模、定价、投入、运营、财务分组，默认只展示关键参数，并支持已修改、敏感性、全部筛选及中文名/字段 key/单位搜索。
- 参数卡改为轻量决策卡，增加状态与影响提示；普通参数进入高级折叠区，参数说明可随卡片交互切换。
- 收入和成本项增加占比与轻量进度条；敏感性分析独立成分析工具 Tab。

Decision:
- 页面筛选、搜索、Tab、说明选择、抽屉和折叠状态均为本地 UI 状态，不修改蓝图持久化结构或计算输入。
- 保持报价公式、资金计划、ICT 映射、自动同步、保存、输出包和正式 Rust 效益计算逻辑不变。

Validation:
- 在 `1440px` 与 `760px` 浏览器视口检查页面级横向溢出、关键参数默认态、筛选、搜索、四 Tab 和抽屉展开/收起。
- 通过智算专项测试、TypeScript 检查和生产构建。

## 2026-06-12

### 智算与 ICT 实时双向联动

Created:
- `ictCalculationInput.ts`: shared formal ICT funding-plan validation and cashflow input builder.
- `ictSync.ts`: revision fingerprints, subject snapshots, ICT override reconciliation, merge conflicts, and formula-control restoration.
- `sync_ai_compute_quote_to_ict`: Rust command that calculates through the existing ICT engine and atomically persists quote/ICT/formal metrics.

Modified:
- 智算参数、公式、映射、税率和计划停止编辑 `500ms` 后自动同步；ICT 折现率和正式财务假设保持原值，智算年份覆盖项目周期。
- 智算右侧效益指标和敏感性分析改为直接使用 Rust ICT 计算结果，不再展示本地近似正式指标。
- 异常、停用或关闭输出的已映射项目按零同步；其他有效科目继续提交。
- ICT 保存单来源人工修改时反写为 `ict_override`，多来源人工修改标记 `merge_conflict`；恢复公式控制后重新自动同步。
- Revision 冲突会合并最新持久化联动状态并重试，旧请求不能覆盖新编辑。
- 智算同步适配器会补齐旧版/不完整 lifecycle 快照缺失的项目身份、产权、分布数组和标准科目字段，避免 `IctInput` 在 Tauri IPC 反序列化阶段失败，并在成功同步后自修复持久化输入。

Decision:
- ICT 是正式财务参数、现金流和效益指标的唯一数据源；智算只拥有业务公式、参数、业务项金额和业务项计划。

Tests:
- 前端专项测试覆盖折现率保留、共享现金流输入、异常清零、单来源覆盖、恢复控制和多来源冲突。
- Rust 测试覆盖 revision 拦截及蓝图、ICT 状态、正式指标的原子提交。

## 2026-06-12

### 智算与 ICT 测算口径对齐

Modified:
- `calculations.ts`: quote totals remain tax-inclusive, while gross profit, gross margin, and per-device monthly cost now use tax-exclusive values, matching the ICT calculator.
- `AiComputeQuoteView.tsx`: metric labels explicitly distinguish tax-inclusive quote totals from ICT tax-exclusive benefit metrics.
- `ictExport.ts`: confirmed ICT exports now carry the quote blueprint project cycle and overwrite stale ICT cycle values in lifecycle and cashflow payloads.
- `project_state/mod.rs`: the import transaction now synchronizes project aggregate revenue/cost and project years, and clears stale project summary metrics.
- Frontend and Rust tests cover the H200 `21.13%` ICT margin, cycle synchronization, project metadata, and stale-cache invalidation.

Decision:
- ICT financial formulas remain authoritative and unchanged. The adapter aligns quote presentation and imported state with that existing tax-exclusive calculation contract.

## 2026-06-11

### 智算公式 Token 光标定位

Created:
- `formulaTokenEditing.ts`: pure helpers for cursor clamping, insertion at an arbitrary Token boundary, token deletion, and backspace-before-cursor behavior.

Modified:
- `QuoteFormulaCalculator.tsx`: added `N+1` clickable insertion positions, active caret presentation, Token-click positioning, cursor-relative insertion, deletion cursor correction, and cursor-relative backspace.
- Token delete controls are visually hidden by default and appear only on Token hover or keyboard focus.
- Formula and result preview now render at the top of the expanded calculator before the Token editing area.
- Formula preview references no longer use brace characters; parameters, calculated results, and constants use inline-code-style semantic emphasis.
- `test_ai_compute_quote.cjs`: added regression coverage for middle insertion and cursor movement after deletion.

Decision:
- Cursor position remains local UI state and is not persisted as part of the quote formula or project blueprint.

## 2026-06-11

### 智算报价折叠式 DIY 计算器

Modified:
- `types.ts`, `formulaEngine.ts`, and `calculations.ts`: upgraded quote formulas to ID-backed Version 2 tokens, added safe parsing for arithmetic/parentheses/`SUM`, line-item dependency calculation, disabled-reference warnings, missing-reference errors, and circular-reference isolation.
- `AiComputeQuoteView.tsx` and `QuoteFormulaCalculator.tsx`: changed every revenue/cost calculation process to a default-collapsed editor with parameter/result/fixed-value insertion, operators, parentheses, `SUM`, comma, clear, undo, token removal, formula preview, and result/error preview.
- `presets.ts`: expressed the H200 preset with Version 2 formulas and changed capital cost to depend on the calculated machine, maintenance, and networking costs.
- `test_ai_compute_quote.cjs`: added coverage for dependency updates, `SUM`, cycles, divide-by-zero, missing references, disabled references, legacy compatibility, and expansion-state isolation.

Decisions:
- Formulas persist stable business IDs; display names are resolved at render time.
- No formula path uses JavaScript `eval`.
- Formula expansion is local component state and never becomes a second business-state source.
- Cycles fail only the involved items; unrelated quote calculations remain available.

## 2026-06-11

### 智算报价测算 Phase 1

Created:
- `src-ui/src/features/ai-compute-quote/`: added data types, H200 preset, pure calculation/output/sensitivity functions, Zustand draft/persistence store, and the independent quote blueprint page.
- `src-ui/scripts/test_ai_compute_quote.cjs`: covers formulas, percentage scaling, invalid inputs, H200 totals, output merging/filtering, and non-mutating sensitivity analysis.
- `docs/modules/ai-compute-quote.md`: records module ownership, financial units, persistence, ICT safety boundary, and future adapter rules.

Modified:
- `App.tsx` and `useNavigationStore.ts`: added the Hub card and `ai_compute_quote` route with Workspace gating.
- `package.json`: added the focused `test:ai-compute-quote` command.

Decisions:
- Project persistence reuses `project_settings` under `ai_compute_quote::active`; no schema or Rust change is needed.
- The first-stage ICT action is output-package preview only. It does not write formal ICT subjects, subject funding plans, cashflow, or benefit state.
- NPV, NPV rate, IRR, and payback are not reimplemented in the quote module. They remain pending a read-only ICT preview adapter.

## 2026-06-10

### macOS 1.0.1 Release Packaging

Modified:
- Root npm package, Tauri app configuration, and Rust package versions were aligned to `1.0.1`.
- The Apple Silicon macOS release is packaged as a DMG through the Tauri release build.

Distribution:
- The local build environment has no Apple Developer signing identity. The generated application uses ad hoc signing and is not Apple-notarized.
- The full `.app` bundle was signed with the explicit ad hoc identity before DMG creation, and strict deep signature verification passed.
- Artifact: `云数中心工具集_1.0.1_aarch64.dmg`.
- SHA-256: `8c036dda1111d43a967fd3e3820820f699c9f17bf37f07eca17996a9ae47f8c9`.

## 2026-06-10

### ICT NPV-Rate Cost Reverse Boundary Fix

Modified:
- [useIctCalculations.ts](../src-ui/src/hooks/useIctCalculations.ts): Replaced the single zero-cost reachability check with boundary probing that evaluates both `0` and `0.01` yuan for NPV-rate cost reverse calculation, then starts binary search from the higher-metric valid boundary.
- [ictReverseSearch.ts](../src-ui/src/lib/ictReverseSearch.ts): Added pure helpers for reverse boundary probe selection.
- [calculator.rs](../src-tauri/src/benefit/calculator.rs): Added regression tests for the zero-outflow NPV-rate convention and the `78000` IT integration revenue plus `500` CT product revenue/cost scenario targeting `0.10`.
- [subject-funding-plan.md](../docs/modules/subject-funding-plan.md): Recorded the reverse boundary rule.

Decision:
- The backend financial definition remains unchanged: NPV rate is `NPV / discounted cash outflow`, and a zero denominator returns 0. The frontend must not treat that sentinel value as the mathematical maximum for cost reverse calculation. The minimum positive currency probe enters the valid ratio domain without changing formulas or persisted business data.

Tests:
- `node scripts/test_ict_reverse_search.cjs`: passed.
- All `scripts/test_subject_funding_*.cjs`: passed.
- `npx tsc --noEmit`: passed.
- `npm run build`: passed with the existing Vite chunk-size warning.
- `cargo test benefit::calculator::tests`: passed, 9 tests.

## 2026-06-05

### Common Preset Quick-Fill Field Header Alignment

Modified:
- [CommonPresetQuickFill.tsx](../src-ui/src/components/common-presets/CommonPresetQuickFill.tsx): Added `CommonPresetFieldHeader`, a reusable form header that keeps the visible label and preset actions in one responsive row.
- [IctBasicInfo.tsx](../src-ui/src/components/IctBasicInfo.tsx) and [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Replaced split label/action rows with `CommonPresetFieldHeader` across the first-phase preset-connected fields.
- [common-presets.md](../docs/modules/common-presets.md) and [CURRENT_TASK.md](../docs/CURRENT_TASK.md): Recorded the form-side layout rule.
- [CommonPresetQuickFill.tsx](../src-ui/src/components/common-presets/CommonPresetQuickFill.tsx): Added `CommonPresetLabelHeader` for plain fields that need to align with preset-enabled fields, and compacted the quick-fill action buttons.
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Removed the detached sign-off payment preset band and attached the revenue/expenditure preset actions to the actual payment input labels.

Decision:
- Preset buttons should not be aligned with fixed offsets or separate `justify-end` rows. They now use a shared flex header so normal desktop fields keep the label and actions on the same line, while narrow containers can wrap naturally.
- Forms that mix preset and non-preset fields should use the shared header components for both variants so neighboring inputs align vertically.

## 2026-06-04

### Common Materials Quick-Fill Field Coverage Expansion

Modified:
- [presetFieldKeys.ts](../src-ui/src/lib/presetFieldKeys.ts): Added stable keys for demand-import fields, meeting-review fields, sign-off IT/CT service fields, and shared revenue/expenditure payment method fields.
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Added `CommonPresetQuickFill` controls for project demand unit, demand service content, customer confirmation, deployment environment requirement, onsite support staff, IT/CT construction content, revenue collection method, expenditure payment method, time requirement, and sign-off IT/CT service content.
- [common-presets.md](../docs/modules/common-presets.md), [CURRENT_TASK.md](../docs/CURRENT_TASK.md), and [test_common_presets.cjs](../src-ui/scripts/test_common_presets.cjs): Recorded and tested the expanded field coverage.

Decision:
- Revenue collection and expenditure payment methods use shared `payment.*` field keys because meeting-review and sign-off sections write the same formal `revCollection` / `expPayment` state. This lets one saved common material appear in both places without duplicating records.

### Common Materials & Project Presets Phase 1

Created:
- [common_presets.rs](../src-tauri/src/common_presets.rs): Added workspace-scoped reusable material commands for listing, saving, enabling/disabling, soft deletion, and usage tracking.
- [PresetCenterView.tsx](../src-ui/src/views/PresetCenterView.tsx): Added the independent "常用资料与项目预设" management page with short-field and long-text tabs.
- [presetFieldKeys.ts](../src-ui/src/lib/presetFieldKeys.ts): Added stable field-key definitions for reusable content and future project preset binding.
- [CommonPresetQuickFill.tsx](../src-ui/src/components/common-presets/CommonPresetQuickFill.tsx): Added reusable field-side picker/save control.
- [commonPresetService.ts](../src-ui/src/services/commonPresetService.ts): Added frontend IPC wrapper for common preset commands.
- [test_common_presets.cjs](../src-ui/scripts/test_common_presets.cjs): Added field-key catalog tests.

Modified:
- [db.rs](../src-tauri/src/db.rs): Added `common_presets` table initialization and schema version 6 migration.
- [main.rs](../src-tauri/src/main.rs): Registered common preset Tauri commands.
- [App.tsx](../src-ui/src/App.tsx), [useNavigationStore.ts](../src-ui/src/store/useNavigationStore.ts), and [iconMap.ts](../src-ui/src/components/icons/iconMap.ts): Added the first-level Hub entry and route for `preset_center`.
- [IctBasicInfo.tsx](../src-ui/src/components/IctBasicInfo.tsx): Connected quick fill to customer name and project background.
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Connected quick fill to project background, technical solution, meeting reviewers, branch/department name, and risk/project owner fields.
- Project context docs were updated to record the module boundary, SQLite structure, fieldKey mechanism, first connected fields, and later-phase exclusions.

Tests:
- Ran `npx tsc --noEmit` in `src-ui`: passed.
- Ran `node scripts/test_common_presets.cjs` in `src-ui`: passed.
- Ran all `scripts/test_subject_funding*.cjs` in `src-ui`: passed.
- Ran `npm run build` in `src-ui`: passed with the existing Vite chunk-size warning.
- Ran `cargo fmt -- --check` in `src-tauri`: passed after formatting.
- Ran `cargo test common_presets::tests` in `src-tauri`: passed.
- Ran `cargo test benefit::calculator::tests` in `src-tauri`: passed.
- Ran full `cargo test` in `src-tauri`: 10 passed, 2 failed because the local docfill test template `项目全生命周期文件模版/效益分析表 .xlsx` is missing.

Decision:
- Presets are only reusable fill sources. After a user applies content, the owning form state is updated through existing setters and save domains; document generation and AI context continue to read official project/template state.
- Phase 1 deliberately does not implement full project preset templates, project-creation preset selection, one-click multi-field application, automatic history extraction, AI recommendation, or AI auto-fill.

### ICT Selection Fee: Fixed Quote/Limit Anchor

Modified:
- [useIctCalculations.ts](../src-ui/src/hooks/useIctCalculations.ts): Added `selectionFeeAnchor` for mutually exclusive quote/limit anchoring. Markup changes now call the forward selection-fee command when quote is fixed and the reverse command when limit is fixed. Added request-sequence protection so stale async invoke results cannot overwrite newer user input.
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Added compact fixed-state dot buttons beside supplier quote and selection limit labels, wired with `aria-pressed`, and added accessible labels to the selection-fee inputs.
- [selection-fee.md](../docs/modules/selection-fee.md), [PROJECT_INDEX.md](../docs/PROJECT_INDEX.md), and [CURRENT_TASK.md](../docs/CURRENT_TASK.md): Recorded the anchor model, UI behavior, and validation status.

Tests:
- Ran `npx tsc --noEmit` in `src-ui`: passed.
- Ran `npm run build` in `src-ui`: passed.
- Ran `cargo test` in `src-tauri`: passed (11/11; existing warnings only).
- Browser smoke-tested the local Vite page: fixed-dot default, mutual exclusion, and edit-to-anchor state passed. Numeric invoke calculation requires the Tauri runtime and was not executed in the plain browser page.

Decision:
- The root cause was that markup edits always used supplier quote as the implicit source of truth. Adding another local conditional would preserve the hidden coupling, so the bounded fix makes the quote/limit source of truth explicit and reuses the existing Rust forward/reverse commands.

### ICT Subject Funding Plans Final: Migration and Single Official Source

Modified:
- [ictSubjectFundingPlan.ts](../src-ui/src/lib/ictSubjectFundingPlan.ts): Added migration version `SUBJECT_FUNDING_PLAN_MIGRATION_VERSION = 1`, `legacy_migration` audit reason, migration-plan creation, and `migrateLegacySubjectFundingPlans()` to fill only missing non-zero subject plans while preserving existing valid/custom/equal plans and surfacing invalid plans through coverage validation.
- [useIctState.ts](../src-ui/src/hooks/useIctState.ts) and [useIctCalculations.ts](../src-ui/src/hooks/useIctCalculations.ts): Defaulted state and payloads to `subject_funding_plans`; official annual cashflow overrides now always come from subject funding plans when coverage is valid, and `CashflowSegment` no longer contributes formal annual cashflow.
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Migrates old project loads after tax-item hydration, writes `subjectFundingPlanMigrationVersion`, removes the calculation-source switch, and keeps coverage summary, batch init, locate, clear-all, balance allocation, CT linkage, and smart reverse flows on the new source.
- [IctBasicInfo.tsx](../src-ui/src/components/IctBasicInfo.tsx), [IctCashflowTable.tsx](../src-ui/src/components/IctCashflowTable.tsx), and [IctSubjectFundingPlanEditor.tsx](../src-ui/src/components/IctSubjectFundingPlanEditor.tsx): Removed old model A-E / segment funding UI and old-source explanatory text; cashflow preview and drill-down now present only the subject-plan source.
- [ProjectFilesTab.tsx](../src-ui/src/components/project/ProjectFilesTab.tsx), [projectService.ts](../src-ui/src/utils/projectService.ts), [models.rs](../src-tauri/src/benefit/models.rs), [excel.rs](../src-tauri/src/benefit/excel.rs), and [calculator.rs](../src-tauri/src/benefit/calculator.rs): Added migration-version compatibility and made Excel import payloads use the subject funding source.
- Added [test_subject_funding_migration.cjs](../src-ui/scripts/test_subject_funding_migration.cjs) and [test_subject_funding_final.cjs](../src-ui/scripts/test_subject_funding_final.cjs).

Tests:
- Ran `for f in scripts/test_subject_funding*.cjs; do node "$f"; done` in `src-ui`: passed.
- Ran `npx tsc --noEmit` in `src-ui`: passed.
- Ran `npm run build` in `src-ui`: passed.
- Ran `cargo test` and `cargo fmt -- --check` in `src-tauri`: passed; existing warnings remain.

Decision:
- The root issue was not just a visible source selector. Keeping `legacy_model` as a runnable branch meant old segment schedules could still compete with subject plans as annual cashflow sources. The bounded fix removes the user switch and formal legacy branch while retaining old fields for read/migration compatibility.
- Existing abnormal subject plans are intentionally not overwritten by migration. The app fills missing non-zero subjects, then lets the canonical coverage validator block formal calculation and saving until the user resolves the invalid rows.

### ICT Funding Plans: Default Activation and Cashflow Source Repair

Modified:
- [ictSubjectFundingPlan.ts](../src-ui/src/lib/ictSubjectFundingPlan.ts): Changed zero-amount sync from "disable and keep record" to removing the subject plan so the UI returns to "未维护"; plan mode switches and custom annual edits now force `enabled: true`; normalization preserves sync audit fields.
- [useIctState.ts](../src-ui/src/hooks/useIctState.ts): Removed the old gate that only synced plans when `cashflowCalculationSource === "subject_funding_plans"`. Amount edits now always sync subject plans, create missing upfront plans, backfill other positive missing subjects, and switch the formal cashflow source to `subject_funding_plans` on positive amount input.
- [IctSubjectFundingPlanEditor.tsx](../src-ui/src/components/IctSubjectFundingPlanEditor.tsx): Opening a zero-amount subject no longer creates a persisted 0 yuan plan; missing plans show an unchecked enable state until explicitly created.
- [subject-funding-plan.md](../docs/modules/subject-funding-plan.md) and [CURRENT_TASK.md](../docs/CURRENT_TASK.md): Recorded the new default activation and zero-reset behavior.

Tests:
- Ran `npm run build` in `src-ui`: passed.
- Ran `for f in scripts/test_subject_funding*.cjs; do node "$f"; done` in `src-ui`: passed.

Decision:
- The root cause was split ownership between subject amount input and cashflow source selection. Users could maintain a subject annual plan while the official cashflow table still read the legacy global model, so custom year-2 amounts appeared in the editor but not in the 10-year result.
- The stable bounded fix makes subject amount input the convergence point: positive amounts create/enable default subject plans and activate the subject-plan source; clearing an amount removes that subject's plan and returns it to "未维护".

## 2026-06-03

### ICT Lifecycle Phase 3.5/4 Regression Repair: Legacy Restore, Clear-All, Coverage Locate, Excel Cashflow

Modified:
- [models.rs](../src-tauri/src/benefit/models.rs): Added backward-compatible serde defaults for IT cashflow fields and `IctResult.cashflow` alias/default support so legacy snapshots without Phase 4 fields can still deserialize.
- [calculator.rs](../src-tauri/src/benefit/calculator.rs): Updated test `IctInput` builders for the optional `subject_funding_plans` field.
- [useIctState.ts](../src-ui/src/hooks/useIctState.ts): Added a unified `clearFinancialSubjects()` state action that clears all revenue/cost amounts, subject names, subject funding plans, Model E segment amount schedules, balance allocation, tail-difference state, and reconciliation prompts.
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Restored tax items from both `incl_tax/tax_rate` and `incl/tax/excl` shapes; added confirmed "一键清空全部收入和支出"; added coverage issue drill-down that reuses `subjectFundingCoverage.issues[0]` and auto-opens the target funding plan editor.
- [IctSubjectFundingPlanEditor.tsx](../src-ui/src/components/IctSubjectFundingPlanEditor.tsx): Added `forceOpenToken` so coverage locate can expand a collapsed plan editor.
- [IctCashflowTable.tsx](../src-ui/src/components/IctCashflowTable.tsx): Completed the project-wide vs IT-only 10-year view with IT cumulative net and PV values, avoiding invalid `NaN` display.
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Fixed Excel cashflow variable source from the incorrect `metrics.cashflows` to the formal `metrics.cashflow`; only emits `CASH_IN_Y1..10` and `CASH_OUT_Y1..10` when official cashflow rows exist so template fallback formulas remain intact.
- [subject-funding-plan.md](../docs/modules/subject-funding-plan.md) and [CURRENT_TASK.md](../docs/CURRENT_TASK.md): Recorded repair rules and current validation status.

Tests:
- Ran `npm run build` in `src-ui`: passed.
- Ran `node scripts/test_subject_funding_plan.cjs`, `node scripts/test_subject_funding_cashflow.cjs`, and `node scripts/test_subject_funding_sync.cjs` in `src-ui`: passed.
- Ran `cargo test` in `src-tauri`: 11/11 passed; existing warnings remain.
- Ran `cargo fmt -- --check` in `src-tauri`: passed after formatting.

Decision:
- The old-project zero-amount symptom was addressed at the deserialization boundary rather than masking it in the UI. Legacy output metrics missing new IT fields now load with zero IT defaults, preserving existing project-wide values.
- Clear-all treats custom subject naming as data on fixed standard subjects because the current catalog has no separate dynamic subject-instance structure. The safe behavior is to restore fixed subject rows to blank initial state and remove all user-entered financial schedules.
- Coverage locate reuses the canonical validation result so issue ordering and classification cannot drift from the blocking calculation-source logic.
- Investment-benefit Excel multi-year cashflow uses the official calculator result, not a parallel reconstruction, preserving old-model and subject-plan compatibility.

### ICT Subject-Level Funding Plans Phase 3.5 & 4: Status Semantics, Batch Init, & Drill-Down

Created:
- [test_subject_funding_phase4.cjs](../src-ui/scripts/test_subject_funding_phase4.cjs): Comprehensive tests for zero-recovery (proportional preservation), exact reason tracking, batch initialization skipping existing plans, and annual cashflow drill-down generation.
- [docs/modules/subject-funding-plan.md](../docs/modules/subject-funding-plan.md): Modular design doc covering the entire subject funding plan system, rules, and integration boundaries.

Modified:
- [ictSubjectFundingPlan.ts](../src-ui/src/lib/ictSubjectFundingPlan.ts): Added `lastValidAnnualInclValues`, `lastChangeReason`, and `lastChangedAt` to `SubjectFundingPlan` state for zero-recovery and transparency. Added `initializeMissingSubjectFundingPlans` helper and `buildAnnualCashflowSubjectContributions` drill-down generator. Appended `revenueSubjects` and `costSubjects` arrays to `FundingPlanCoverageResult`.
- [useIctState.ts](../src-ui/src/hooks/useIctState.ts): Updated `updateTaxItem` and `updateTaxItemsInclBatch` to accept optional `reason` arguments from caller layers, propagating them to the sync algorithm to replace "manual_amount_sync" with accurate business contexts like "reverse_calculation_sync" or "ct_linkage_sync".
- [useIctCalculations.ts](../src-ui/src/hooks/useIctCalculations.ts): Injected specific update reasons (`reverse_calculation_sync` and `balance_allocation_sync`) into amount update dispatches.
- [IctSubjectFundingPlanEditor.tsx](../src-ui/src/components/IctSubjectFundingPlanEditor.tsx): Added a subtle, friendly UI status pill displaying the `lastChangeReason` (in Chinese) for auto-adjusted plans.
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Added batch initialization buttons ("一键生成...") visible when `subject_funding_plans` calculation source is active. Re-enabled the smart reverse button while in subject plans mode (removed accidental phase 3 blocking logic).
- [IctCashflowTable.tsx](../src-ui/src/components/IctCashflowTable.tsx): Implemented an interactive drill-down state `expandedYear` that renders nested per-subject revenue/cost contribution breakdowns natively below each cashflow year row when expanded.

Tests:
- Ran `node scripts/test_subject_funding_phase4.cjs` in `src-ui`: 4/4 passed.
- Ran `cargo test` in `src-tauri`: 10/10 passed.
- Ran `npm run build` in `src-ui`: Build succeeded.

Decision:
- Zero Recovery: The UI requires subjects to seamlessly toggle on and off. Instead of deleting zeroed plans, we preserve their scale in `lastValidAnnualInclValues`. Re-adding an amount restores the prior timeline proportions precisely.
- Explicit Sync Reason: Differentiating manual edits from reverse calcs or linkages directly aids user comprehension. Storing `lastChangeReason` at the domain layer ensures the UI displays an accurate origin narrative, increasing trust.
- Drill-Down Readability: Aggregated NPV values hide composition. Exposing the exact `buildAnnualCashflowSubjectContributions` mappings within the cashflow table removes the "black box" feeling, letting users trace final outputs directly back to input subjects.

### ICT Subject-Level Funding Plans Phase 3: Amount-Change Synchronization

Created:
- [test_subject_funding_sync.cjs](../src-ui/scripts/test_subject_funding_sync.cjs): 12 pure-function tests covering proportional scaling, tail-difference correction, auto-create upfront, clear-to-zero, zero-total fallback, batch sync, CT linkage, mode preservation, scale-down, negative amounts, and immutability.

Modified:
- [ictSubjectFundingPlan.ts](../src-ui/src/lib/ictSubjectFundingPlan.ts): Added `syncSubjectFundingPlanToAmount` (single-subject proportional scale with tail correction, auto-create, and zero-clear) and `syncSubjectFundingPlansToAmounts` (batch variant). Uses integer-cents arithmetic internally to avoid floating-point drift.
- [useIctState.ts](../src-ui/src/hooks/useIctState.ts): Injected funding plan sync into `updateTaxItem` and `updateTaxItemsInclBatch`. When `cashflowCalculationSource === "subject_funding_plans"`, any `incl` or `excl` field change (including CT linkage side-effects) triggers `syncSubjectFundingPlansToAmounts` in the same React state batch. Tax-rate-only changes (`field === "tax"`) do not trigger sync.
- [useIctCalculations.ts](../src-ui/src/hooks/useIctCalculations.ts): Added `fundingPlansOverride` option to `buildInputDataPayload` so candidate evaluations use simulated-synced plans. Added `buildCandidateSyncUpdates` helper for CT-aware sync update construction. Modified `buildReverseCandidate` and `buildLockedTotalStructureCandidate` to compute and inject synced plans during candidate evaluation.
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Removed the `subject_funding_plans` reverse calculation block — smart reverse now works under both calculation sources.

Tests:
- Ran `node scripts/test_subject_funding_sync.cjs` in `src-ui`: 12/12 passed.
- Ran `node scripts/test_subject_funding_plan.cjs` in `src-ui`: passed (existing tests unaffected).
- Ran `node scripts/test_subject_funding_cashflow.cjs` in `src-ui`: passed (existing tests unaffected).
- Ran `npx tsc --noEmit` in `src-ui`: zero errors.

Decision:
- Sync logic centralized in `updateTaxItem` / `updateTaxItemsInclBatch` (the convergence points for all amount writes), rather than scattered across each caller. This ensures all 5 write paths (manual edit, normal reverse, locked-total reverse, balance allocation, CT linkage) are automatically covered.
- Proportional scaling preserves the user's custom year-by-year distribution shape. Mode (`upfront`, `equal`, `custom`) is preserved across syncs.
- `legacy_model` mode is completely unaffected; sync is gated by `cashflowCalculationSource`.

### ICT Subject-Level Funding Plans Phase 2

Created:
- [test_subject_funding_cashflow.cjs](../src-ui/scripts/test_subject_funding_cashflow.cjs): Added pure Node tests for subject-plan coverage validation, cashflow source normalization, and annual cashflow aggregation with per-subject tax conversion.

Modified:
- [ictSubjectFundingPlan.ts](../src-ui/src/lib/ictSubjectFundingPlan.ts): Added `CashflowCalculationSource`, coverage issue/result types, full coverage validation, and subject-plan annual cashflow generation helpers. The annual cashflow helper converts each subject's annual tax-inclusive plan values to tax-exclusive values using that subject's tax rate before summing.
- [useIctState.ts](../src-ui/src/hooks/useIctState.ts): Added `cashflowCalculationSource`, normalized setter, and legacy default behavior.
- [useIctCalculations.ts](../src-ui/src/hooks/useIctCalculations.ts): Added coverage and annual cashflow derivation, serialized `cashflow_calculation_source`, and routed valid subject-plan annual arrays through existing Rust direct cashflow override fields. Invalid active subject-plan coverage blocks recalculation and keeps the previous result.
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Restores, saves, and clears the calculation source; renders the cashflow source selector and coverage summary; blocks switching into subject-plan mode when coverage is incomplete; blocks official benefit saves and document generation when active subject-plan coverage is invalid; disables smart reverse while subject-plan mode is active.
- [IctSubjectFundingPlanEditor.tsx](../src-ui/src/components/IctSubjectFundingPlanEditor.tsx) and [IctCashflowTable.tsx](../src-ui/src/components/IctCashflowTable.tsx): Updated UI text and cashflow previews so the active source is visible and stale subject-plan states are explicit.
- [projectService.ts](../src-ui/src/utils/projectService.ts), [ProjectFilesTab.tsx](../src-ui/src/components/project/ProjectFilesTab.tsx), [models.rs](../src-tauri/src/benefit/models.rs), [calculator.rs](../src-tauri/src/benefit/calculator.rs), and [excel.rs](../src-tauri/src/benefit/excel.rs): Added payload/source compatibility fields and legacy defaults.

Tests:
- Ran `node scripts/test_subject_funding_plan.cjs` in `src-ui`: passed.
- Ran `node scripts/test_subject_funding_cashflow.cjs` in `src-ui`: passed.
- Ran `npm run build` in `src-ui`: TypeScript and Vite build passed.
- Ran `cargo test` in `src-tauri`: 10 tests passed; only existing compiler warnings were emitted.

Decision:
- Subject-level funding plans affect official cashflow only after an explicit user switch to `subject_funding_plans`; old projects and projects with plans remain on `legacy_model` unless switched.
- New-source coverage failures do not fall back to legacy. The app keeps prior cashflow/metrics visible with an invalid/stale warning and blocks benefit saves/doc generation until coverage is repaired or legacy mode is selected.
- Smart reverse remains a legacy-source feature for now because the current reverse solvers do not update every affected subject-level annual plan.

### ICT Subject-Level Funding Plans Phase 1

Created:
- [ictSubjectFundingPlan.ts](../src-ui/src/lib/ictSubjectFundingPlan.ts): Added subject-level funding plan types and pure helpers for `side + groupId + key` IDs, default upfront plans, 1-10 year equal splits, custom annual value updates, normalization, advisory validation, upsert, and deleted-subject cleanup.
- [IctSubjectFundingPlanEditor.tsx](../src-ui/src/components/IctSubjectFundingPlanEditor.tsx): Added inline revenue/cost subject funding plan editor with "收款计划" / "付款计划" wording, three modes (`upfront`, `equal`, `custom`), 10-year tax-inclusive annual inputs, enabled toggling, and difference/status messaging.
- [test_subject_funding_plan.cjs](../src-ui/scripts/test_subject_funding_plan.cjs): Added a lightweight Node logic test for funding-plan helpers.

Modified:
- [useIctState.ts](../src-ui/src/hooks/useIctState.ts): Added `subjectFundingPlans` current state plus normalized setter, upsert, and subject-ref cleanup methods.
- [useIctCalculations.ts](../src-ui/src/hooks/useIctCalculations.ts): Serializes `subject_funding_plans` into the existing lifecycle/AI payload for persisted context. Calculation formulas and reverse solvers are unchanged.
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Restores subject funding plans from lifecycle payloads or cashflow assumptions, persists them under cashflow assumptions, marks cashflow dirty when they change, clears stale plans when entering free/empty contexts, and renders the inline editor below each subject row.
- [PROJECT_STATUS.md](../docs/PROJECT_STATUS.md), [ARCHITECTURE_MAP.md](../docs/ARCHITECTURE_MAP.md), [AI_CONTEXT.md](../docs/AI_CONTEXT.md): Updated long-term project status and architecture context.

Tests:
- Ran `node scripts/test_subject_funding_plan.cjs` in `src-ui`: passed.
- Ran `npm run build` in `src-ui`: TypeScript and Vite build passed.
- Ran `cargo test` in `src-tauri`: 10 tests passed; only existing compiler warnings were emitted.

Decision:
- Subject-level funding plans are current business state, not a calculation source in Phase 1.
- Plans bind to concrete subject instances using `side + groupId + key`, not just `subjectCode`, so future duplicate standard subjects can maintain independent schedules.
- Existing `cashflowModel`, `CashflowSegment`, NPV/IRR/payback, balance allocation, and smart reverse behavior remain unchanged.
- Existing projects without `subject_funding_plans` / `subjectFundingPlans` normalize to `{}`. No migration from old funding models or segment schedules is performed in this phase.

### Custom Accent Color Real-Time Preview Fix

Modified:
- [SettingsView.tsx](../src-ui/src/components/settings/SettingsView.tsx): Resolved a bug where custom accent colors (especially low-contrast inputs) did not update the real-time preview or warning banners correctly. Implemented a `useEffect` that continuously translates user-typed hex values (including HSL self-adjusted safe versions) and applies them to the DOM for instant preview. Supported flexible hex formats (3/6 digits, with or without leading '#').

### Front-End Advanced Appearance Customization & Accessibility Safeguards (Phase 3)

Created:
- [colorUtils.ts](../src-ui/src/theme/colorUtils.ts): Custom RGB/HSL conversions, relative luminance calculations, and WCAG contrast ratio checkers.
- [deriveAccentTokens.ts](../src-ui/src/theme/deriveAccentTokens.ts): Generates WCAG-compliant derived HSL accent tokens (`primary`, `primary-foreground`, `primary-soft`, `ring`, `accent`, `accent-foreground`) from custom colors using HSL lightness shifting.
- [test_color_math.js](../scratch/test_color_math.js): Verification script to test HSL conversion, WCAG contrast calculations, and automatic lightness adjustments for custom accent colors.

Modified:
- [appearance.ts](../src-ui/src/theme/appearance.ts): Extended `AppearanceSettings` type to support `contrastPreference`, `customAccent` settings, and version fields. Upgraded standard default settings to version 3.
- [presets.ts](../src-ui/src/theme/presets.ts): Refined HSL dark theme preset mappings (`DARK_THEMES`) for all 5 presets to establish distinct visual styles in dark modes. Added HSL contrast override variables.
- [applyAppearance.ts](../src-ui/src/theme/applyAppearance.ts): Implemented sequence-based color tokens resolution: standard presets are modified by custom accent calculations and then overlaid with high-contrast rules. Applied `data-contrast` attribute on document element.
- [useAppearanceStore.ts](../src-ui/src/store/useAppearanceStore.ts): Added preference setters and state for Custom Accent and Contrast. Added Version 3 localStorage settings parsing and configuration migration paths.
- [SettingsView.tsx](../src-ui/src/components/settings/SettingsView.tsx): Extended layouts with Accent color palette recommended selections, Custom HTML5 Color Picker, Standard vs High Contrast selectors, and warning banners for auto-adjusted low contrast accent inputs. Enhanced real-time preview panel with focused input ring and simulated AI chatbot bubble elements.
- [index.ts](../src-ui/src/theme/index.ts): Exported `colorUtils` and `deriveAccentTokens`.
- [DESIGN.md](../DESIGN.md): Appended Phase 3 visual specs, contrast boundaries, custom accent derivation rules, and dark theme refinement notes.
- [PROJECT_STATUS.md](../docs/PROJECT_STATUS.md), [ARCHITECTURE_MAP.md](../docs/ARCHITECTURE_MAP.md): Updated project status history, directory maps, and theme components definitions.

Decision:
- Accent color adjustments are validated dynamically against WCAG standards (4.5:1 standard minimum, 7.0:1 high contrast minimum) against light/dark backgrounds. If low contrast is detected, lightness is shifted in HSL, warning the user and applying a safe version.
- High Contrast preference applies pure white/black canvas background blocks, highly visible text, and thick distinct border lines.
- SQLite project database structures, business cashflow calculations, NPV metrics, and AI context composer schemas remain completely untouched.

### Front-End Appearance Settings Center & Theme Runtime Switching (Phase 2)


Created:
- [appearance.ts](../src-ui/src/theme/appearance.ts): Formulates brightness modes, color themes, font scaling levels, interface density configurations, default presets, and font scaling ratio factors.
- [presets.ts](../src-ui/src/theme/presets.ts): Formulates full HSL variables mapping for 5 corporate light themes (`lamber`, `graphite`, `navy`, `forest`, `warmStone`) and system dark base configuration with light preset primary adaptation.
- [applyAppearance.ts](../src-ui/src/theme/applyAppearance.ts): Writes color tokens, font scaling weights, and interface densities to `document.documentElement` attributes and styles dynamically.
- [useAppearanceStore.ts](../src-ui/src/store/useAppearanceStore.ts): Persists user styling configuration to `localStorage`, observes brightness system media query settings, applies styles to root, and manages cross-window event synchronization.
- [SettingsView.tsx](../src-ui/src/components/settings/SettingsView.tsx): Renders custom selectors for all style choices with a real-time mini business preview card and default restoration command.

Modified:
- [index.ts](../src-ui/src/theme/index.ts): Exports all theme, presets, configuration types, and DOM applier helper utilities.
- [index.css](../src-ui/src/index.css): Dynamically scales typographic line heights by `--font-scale` to prevent overlap, and implements custom overrides for card padding, button/input heights, table cell vertical paddings, and page spacing gaps.
- [card.tsx](../src-ui/src/components/ui/card.tsx), [button.tsx](../src-ui/src/components/ui/button.tsx), [input.tsx](../src-ui/src/components/ui/input.tsx): Rewrote primitive sizes and paddings to read custom CSS density variables with standard defaults.
- [main.tsx](../src-ui/src/main.tsx): Runs synchronous appearance store hydration prior to React rendering to ensure color and layouts render without startup flashing.
- [useNavigationStore.ts](../src-ui/src/store/useNavigationStore.ts): Integrated `"settings"` view type and track previous view context to return safely on settings panel close.
- [App.tsx](../src-ui/src/App.tsx): Routes settings views and places Settings button in `HubView` header toolbar.
- [ProjectBoard.tsx](../src-ui/src/views/ProjectBoard.tsx): Places Settings button on the kanban board header toolbar next to the GlobalSaveButton.
- [WorkspaceHeader.tsx](../src-ui/src/components/WorkspaceHeader.tsx): Places Settings button on the header toolbar of all workspaces next to the GlobalSaveButton.
- [DESIGN.md](../DESIGN.md): Documented specifications for presets, mode observers, scaling bounds, density levels, and application mechanisms.
- [PROJECT_STATUS.md](../docs/PROJECT_STATUS.md), [ARCHITECTURE_MAP.md](../docs/ARCHITECTURE_MAP.md), [AI_CONTEXT.md](../docs/AI_CONTEXT.md): Updated project milestones, directory indexes, and visual token guidelines.

Decision:
- Appearance preferences are strictly application-level settings and must never be written to project databases or workspaces.
- Color presets are dynamically set via raw space-separated numbers on root styles to respect Tailwind's HSL wrapper.
- Cross-window visual synchronizations are dispatched via Tauri window event listener bridges.

### Front-End Global Visual Foundation Refactoring (Phase 1)

Created:
- [tokens.ts](../src-ui/src/theme/tokens.ts): Define design system color schemes, border radius (`lg: var(--radius)`, `md`, `sm`), and sizing.
- [typography.ts](../src-ui/src/theme/typography.ts): Declare base line-heights, weight bindings, and typographic roles mapped to `--font-scale`.
- [index.ts](../src-ui/src/theme/index.ts): Centralized theme exports.

Modified:
- [index.css](../src-ui/src/index.css): Added new HSL color variables and typography scale calculations. Integrated `.numeric-value` class.
- [tailwind.config.js](../src-ui/tailwind.config.js): Extended Tailwind configuration with semantic color names and typographic roles.
- [button.tsx](../src-ui/src/components/ui/button.tsx), [input.tsx](../src-ui/src/components/ui/input.tsx), [card.tsx](../src-ui/src/components/ui/card.tsx), [label.tsx](../src-ui/src/components/ui/label.tsx): Refactored primitive components to adhere to visual standards.
- [App.tsx](../src-ui/src/App.tsx), [ProjectBoard.tsx](../src-ui/src/views/ProjectBoard.tsx), [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx), [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Migrated views to utilize semantic styles, typography variables, and tabular numbers.
- [DataManagement.tsx](../src-ui/src/views/DataManagement.tsx): Refactored path lists, workspace cards, status labels, and relocation tables to use semantic HSL tokens.
- [AiChatPanel.tsx](../src-ui/src/components/ai/AiChatPanel.tsx): Refactored connection status badge and quick actions.
- [MessageBubble.tsx](../src-ui/src/components/MessageBubble.tsx): Refactored inline badges, user/assistant chat bubbles, blockquotes, and code block components.
- [DESIGN.md](../DESIGN.md): Updated document with the Lamber Global Visual Specification v1 details.

Tests:
- Run `npm run build` inside `src-ui`: TypeScript compilation and Vite build succeeded.
- Run `cargo test` in `src-tauri`: All unit tests passed.

Decision:
- Visual elements must follow the "No-Line" rule using background HSL shifts instead of thick solid borders.
- Sizing and typography must support dynamic resizing based on `--font-scale` variable.
- Financial numerical representation enforces tabular-nums across the UI.

### ICT Lifecycle Balance Control UI Layout Alignment

Modified:
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Changed the balance control columns container to top-align (using `md:items-start` instead of `md:items-center`), and wrapped the input container and summary card/text elements in a matched `h-[38px]` height container. This ensures Row 1 (labels), Row 2 (inputs/summaries), and Row 3 (1% own-product prompts) align row-by-row across columns.

Tests:
- Ran `npm run build` in `src-ui`: TypeScript compilation and Vite build succeeded.
- Ran `cargo test benefit::calculator::tests` in `src-tauri`: All unit tests passed.

## 2026-06-02

### ICT Lifecycle Subject Role Configuration Optimization

Created:
- [IctSubjectRoleComponents.tsx](../src-ui/src/components/IctSubjectRoleComponents.tsx): Added modular UI components `SubjectRoleActions` (row-level menu for setting/clearing balancing and reverse roles on each subject) and `SelectedSubjectRoleSummary` (card summarizing active role, supporting locate and clear actions), along with helper utilities `scrollToSubject` and `highlightSubjectElement` for smooth-scroll tab switching and visual outline feedback.

Modified:
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Integrates role assignment directly in subject table rows, removing the legacy select dropdowns from the top control areas and right panel. Simplified the right reverse panel to automatically infer target details and reverse side from the selected target, disabling the execute button if no target is set. Removed unused imports and methods.

Tests:
- Ran `npm run build` in `src-ui`: TypeScript check and Vite build passed.
- Ran `cargo test` in `src-tauri`: 8 passed, 2 expected failures due to missing template file.
- Ran `cargo test benefit::calculator::tests`: 7 passed.
- Ran `cargo test ai_context::service::tests`: 1 passed.

Decision:
- Entry point for balancing and reverse target roles is shifted to the concrete subject rows for direct visual context.
- Mutual exclusions (balancing subjects cannot be reverse targets) are enforced with warnings on hover/click.
- Switching balancing subjects prompts for confirmation and preserves the old balancing subject's current amount value.
- Right reverse panel automatically adapts its state based on the side and reverse mode of the active target.

### ICT Model E Structure Reverse Segment Sync

Modified:
- [useIctCalculations.ts](../src-ui/src/hooks/useIctCalculations.ts): Removed the `model_e` amount-mode structure reverse entrance block and added a shared candidate/final segment synchronization path for locked-total structure reverse. The sync maps stable reverse subjects to `CashflowSegment` side/scope buckets, applies target and balancing deltas before `calculate_ict_benefit`, rejects invalid candidates, and writes the accepted synced segment array back only after the final calculation succeeds.
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Extended the structure reverse hint so `model_e` amount mode tells users that segmented cashflow amount plans will be synchronized.

Tests:
- Ran `npm run build` in `src-ui`.
- Ran `cargo test` in `src-tauri`: 8 passed, 2 failed because the local test template `项目全生命周期文件模版/效益分析表 .xlsx` is missing for the two `docfill::tests` lifecycle Excel cases.
- Ran `cargo test benefit::calculator::tests` in `src-tauri`: 7 passed.
- Ran `cargo test ai_context::service::tests` in `src-tauri`: 1 passed.

Decision:
- `model_e` amount-mode structure reverse is supported only where the selected subject and balancing subject map to existing segment buckets. Revenue IT/CT/non-IT-CT subjects map to `revenueScope`; cost IT/CT/non-IT-CT/mixed subjects map to `costScope`.
- Same-bucket transfers preserve the aggregate bucket total; cross-bucket transfers move equal and opposite deltas between buckets. Existing custom annual amount plans are scaled through the amount-mode annual adjustment helper.
- Candidates are excluded from reachability and final write-back if they require a missing bucket or would create a negative segment, bucket, or annual amount.
- CT product and line revenue still mirror to their paired CT cost subjects and now synchronize the linked cost-side segment bucket. If that cross-side linkage collides with an active locked-total investment balancing rule, the solver blocks conservatively instead of attempting a four-variable cross-side reverse.
- The revenue own-product 1% prompt remains display-only.

### ICT Locked-Total Structure Reverse Calculation

Modified:
- [ictReverseCalculation.ts](../src-ui/src/lib/ictReverseCalculation.ts): Added reverse mode resolution for `normal`, `locked_total_structure`, and `blocked`, plus locked-total structure context construction, dual-subject candidate application, and bounded sample-point generation.
- [useIctCalculations.ts](../src-ui/src/hooks/useIctCalculations.ts): Added the locked-total structure solver. It samples the finite reallocatable pool, detects metric-insensitive and unreachable targets, chooses the crossing solution closest to the current target amount, validates total preservation/non-negative amounts, and writes target plus balancing subject results together.
- [useIctState.ts](../src-ui/src/hooks/useIctState.ts): Added `updateTaxItemsInclBatch` for inclusive-amount batch writes, preserving tax-exclusive recomputation and CT revenue-to-cost amount pairing without same-group stale-state overwrite risk.
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Replaced same-side balance blocking with automatic reverse mode detection, added the structure-mode hint in the smart reverse panel, kept balancing subjects disabled, and passed the resolved reverse context into the calculation hook.

Tests:
- Ran `npm run build` in `src-ui`.

Decision:
- Same-side reverse under a valid total balance rule is now handled as structure reverse: the selected target subject changes and the balancing subject moves inversely so the side's inclusive total stays unchanged.
- `model_e` amount mode is blocked for structure reverse because the current segment data model cannot reliably map two changed subjects to distinct cashflow destinations. Normal selected-subject reverse remains supported in amount mode.
- The revenue own-product 1% prompt is still display-only in the current UI; this phase did not add a blocking validation rule for it.

### ICT Dynamic Reverse Subject Calculation

Created:
- [ictReverseCalculation.ts](../src-ui/src/lib/ictReverseCalculation.ts): Added shared helpers for reverse subject stable references, eligible subject options, selected-subject display names, candidate tax-inclusive amount application, CT paired amount mirroring, and balance-allocation conflict validation.

Modified:
- [useIctCalculations.ts](../src-ui/src/hooks/useIctCalculations.ts): Replaced fixed frontend reverse writes to `rev_it_integration` / `cost_it_integration` with a dynamic selected-subject solver. Candidate payloads now support all revenue and cost subject groups, reuse `calculate_ict_benefit`, preserve the existing margin/NPV-rate target metrics, and write final results through `updateTaxItem`.
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Added the smart reverse subject selector, clears invalid subject selections when switching reverse side, disables active balancing subjects, and blocks same-side locked-total reverse conflicts with a clear message.

Tests:
- Ran `npm run build` in `src-ui`.

Decision:
- The current balance allocation implementation is present on both revenue and investment sides, so same-side reverse conflict handling is applied to both sides.
- The reverse subject is identified by stable catalog fields (`side`, `subjectCode`, `groupId`, `key`) and display names are presentation-only.
- Locked-total structure reverse is intentionally not implemented in this phase. When a same-side balance allocation rule is valid, same-side reverse is blocked; cross-side reverse remains allowed.
- The old Rust fixed reverse commands remain registered for compatibility, but the ICT frontend smart reverse panel now evaluates selected subjects through the general `calculate_ict_benefit` path.

### ICT Balance Allocation Rules

Created:
- [ictBalanceAllocation.ts](../src-ui/src/lib/ictBalanceAllocation.ts): Added shared frontend helpers for revenue/investment balance rule normalization, serialization, subject reference matching, inclusive-amount difference calculation, and validation status reporting.

Modified:
- [useIctState.ts](../src-ui/src/hooks/useIctState.ts): Added `balanceAllocation` state with independent `revenue` and `investment` rules.
- [useIctCalculations.ts](../src-ui/src/hooks/useIctCalculations.ts): Serializes `revenue_balance_rule` and `investment_balance_rule` into lifecycle input payloads and AI context payloads.
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Replaced separate quick split panels with revenue/investment total balance controls, removes one-click quick fill and integration-service preview from the control area, shows the revenue own-product minimum prompt as 1% of the revenue total, applies valid balancing amounts through `updateTaxItem`, makes active balancing subject amount fields read-only while preserving tax-rate editing, persists rules in lifecycle/cashflow saves, restores rules from current state, and blocks cashflow/document-generation navigation on negative balance validation.
- [projectService.ts](../src-ui/src/utils/projectService.ts): Added typed balance rule payload fields to `IctInput`.
- [models.rs](../src-tauri/src/benefit/models.rs): Added optional `revenue_balance_rule` and `investment_balance_rule` to `IctInput` so benefit snapshots preserve the new rule configuration.
- [calculator.rs](../src-tauri/src/benefit/calculator.rs) and [excel.rs](../src-tauri/src/benefit/excel.rs): Updated test/import input constructors with disabled balance rules.

Tests:
- Ran `npm run build` in `src-ui`.
- Ran `cargo test` in `src-tauri`.

Decision:
- Balance differences are based on inclusive amounts and are written back only through the existing tax item update path, preserving existing inclusive/tax/exclusive linkage and CT paired update behavior.
- Switching a balancing subject clears the previous balancing subject's inclusive amount to `0` before the new subject receives the balancing difference, so the same balancing amount does not remain on both subjects.
- Negative balance differences are validation errors, not formal amounts. They block cashflow and document generation before the 0-tolerance reconciliation modal can be bypassed.
- Smart reverse calculation remains fixed to the existing current targets and algorithms in this phase. Arbitrary-subject reverse calculation and total-locked reverse solving remain a follow-up phase.

## 2026-06-01

### ICT Sign-off Project Situation Itemization

Modified:
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Reordered the sign-off-only configuration section to Project Background, IT/CT Service Content, advance-payment and post-approval-selection checkboxes, then revenue collection and expenditure payment methods.
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Removed manual sign-off billing-subject override inputs and now generates `PROJECT_INVESTMENT_SITUATION` / `PROJECT_REVENUE_SITUATION` by enumerating non-zero measurement-table subjects with billing-subject-name first, standard-subject fallback, and fixed category prefixes.
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Meeting-review project overall investment wording now reuses the same generated investment sentence as the sign-off project situation investment line.
- [docfill.rs](../src-tauri/src/docfill.rs): Added generation-time normalization for older sign-off templates that still contain the hardcoded IT/CT-only project situation wording.
- [【2025版】ICT项目立项签批表（仅适用50万以下项目）模板.docx](../项目全生命周期文件模版/【2025版】ICT项目立项签批表（仅适用50万以下项目）模板.docx): Replaced the hardcoded IT/CT-only project situation wording with full-line placeholders for generated investment and revenue situation text.

Tests:
- Ran `npm run build` in `src-ui`.

Decision:
- The optional "申请立项后甄选" wording is controlled by a sign-off checkbox and is no longer hardcoded in the template.
- Sign-off project situation wording is presentation-only and continues to use existing tax-exclusive amount fields; no financial calculation paths changed.

### ICT Billing Subject Name Extension

Modified:
- [ictSubjectCatalog.ts](../src-ui/src/lib/ictSubjectCatalog.ts): Added `billingSubjectName` / `billing_subject_name` support and the shared `resolveBillingSubjectPresentation` resolver for Excel display names, document business names, and document dedup keys.
- [useIctState.ts](../src-ui/src/hooks/useIctState.ts), [useIctCalculations.ts](../src-ui/src/hooks/useIctCalculations.ts), and [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Added the "计费科目名称（文书/计费口径）" input for every existing revenue/cost subject, synchronized it across CT paired subjects, and persisted it through lifecycle/cashflow/benefit payloads without changing amount behavior.
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Continued using catalog helpers for Excel variables, sign-off wording, meeting-review wording, and business-composition deduplication so billing subject names now take priority over product/business names.
- [models.rs](../src-tauri/src/benefit/models.rs), [excel.rs](../src-tauri/src/benefit/excel.rs), [ProjectFilesTab.tsx](../src-ui/src/components/project/ProjectFilesTab.tsx), [projectFileService.ts](../src-ui/src/services/projectFileService.ts), and [projectService.ts](../src-ui/src/utils/projectService.ts): Extended the serialized item DTOs with optional `billing_subject_name` while preserving old data without the field.

Tests:
- Ran `npm run build` in `src-ui`.
- Ran `cargo test` in `src-tauri`; existing Excel subject-row tests still verify name/G/Q writes and blank amount clearing.

Decision:
- `customSubjectName` remains the product/business name field; the new billing subject name is stored independently and only affects presentation/export/document wording.
- The resolver priority is `计费科目名称 > 具体业务/产品名称 > 标准科目名称` for Excel/page display and `计费科目名称 > 具体业务/产品名称 > existing fallback` for documents.
- No standard subject rows, Excel formulas, tax calculations, cashflow math, selection fee logic, or reverse-calculation behavior were changed.

### ICT Subject Custom Business Name Extension

Created:
- [ictSubjectCatalog.ts](../src-ui/src/lib/ictSubjectCatalog.ts): Added a fixed subject catalog mapping stable subject codes, UI groups, standard subject names, Excel variable prefixes, and document business prefixes.

Modified:
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx) and [useIctState.ts](../src-ui/src/hooks/useIctState.ts): Added custom business/product name inputs for every existing revenue and cost subject, storing the optional custom name separately from amount/tax fields and restoring old data safely.
- [useIctCalculations.ts](../src-ui/src/hooks/useIctCalculations.ts), [models.rs](../src-tauri/src/benefit/models.rs), [excel.rs](../src-tauri/src/benefit/excel.rs), [ProjectFilesTab.tsx](../src-ui/src/components/project/ProjectFilesTab.tsx), and [projectFileService.ts](../src-ui/src/services/projectFileService.ts): Persist custom subject names through lifecycle/snapshot payloads and preserve them when parsing/importing exported lifecycle Excel files.
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Generates Excel subject display names, document business names, and deduplicated business composition values from the subject catalog.
- [docfill.rs](../src-tauri/src/docfill.rs): Replaced scattered Excel cell writes with a unified `3-直接经济效益评估表` subject-row mapping that writes the subject name, `G` tax-exclusive amount, and `Q` tax-inclusive amount for every standard subject row.

Tests:
- Added a Rust unit test verifying that CT product revenue and CT other product cost write `产品收入（视频监控）` / `其他产品成本（视频监控）`, `G=41.51`, and `Q=44` into the lifecycle Excel sheet.
- Added a Rust unit test verifying that blank amount variables clear `G/Q` cells instead of writing numeric `0`.

Decision:
- Custom business/product names are metadata on fixed standard billing subjects, not new subject rows. The implementation does not support adding/deleting subjects or inserting Excel rows.
- Financial calculations remain keyed by the existing standard subjects and continue to use only amount/tax fields; custom names affect UI labels, Excel output/import, and document wording only.
- Business composition wording deduplicates repeated document business names across revenue and cost, while amount detail fields keep each side separate.
- Empty/zero frontend amount inputs remain blank in generated Excel amount cells; paired CT subject custom names follow existing amount pass-through relationships.

### Meeting Investment Subject Alignment Fix

Modified:
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Meeting-review `PROJECT_TOTAL_INVESTMENT_DETAIL` now uses the same resolved IT/CT cost subjects as the sign-off form variables instead of deriving CT subject wording from the mid-platform capability name.

Decision:
- The presales meeting-review "项目整体投入金额" wording should follow the sign-off form's investment billing subjects. This changes only document subject wording and keeps the existing tax-exclusive IT, CT, mixed-cost, and total investment calculations intact.

### ICT Cashflow Price Persistence Hydration Fix

Modified:
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Added current-state hydration that overlays `project_cashflow_states.assumptions_json`, payment model, and cashflow segments onto lifecycle/snapshot input data before filling calculator fields.

Decision:
- Price fields belong to the cashflow domain during ordinary edits and can be saved without rewriting lifecycle input payloads.
- Reopening an ICT project must treat cashflow assumptions as the latest current editor state for IT/CT revenue and cost inputs, otherwise stale lifecycle payloads can restore old prices.

### Inquiry Vendor Image State Preservation Fix

Modified:
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): One-click three-vendor quote generation now merges existing vendor screenshots into regenerated quote rows instead of clearing `images`.
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Vendor screenshot upload now uses functional state updates so async file reads cannot overwrite newer vendor quote state.

Decision:
- Vendor quote screenshots are part of the vendor row state and must survive quote amount regeneration.
- The merge prefers vendor-name matches and falls back to row index to support both regenerated default vendors and manually edited rows.

### Template Image Document Embedding Fix

Modified:
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Document-generation image payloads now use `assetId` as the primary `data` value and include `assetId` explicitly, preventing frontend preview URLs from being serialized into Word variables.
- [docfill.rs](../src-tauri/src/docfill.rs): JSON image parsing now prefers `assetId`, embeds only asset IDs or legacy `data:image` values, and suppresses unresolved image JSON so raw payload text does not appear in generated Word documents.

Decision:
- Frontend `asset://localhost/...` URLs are preview-only and must not be used as document image source data.
- The backend remains the authority for resolving image binaries through Workspace SQLite asset ownership and file resolution.

### Lifecycle Document Workspace Output Path Fix

Modified:
- [docfill.rs](../src-tauri/src/docfill.rs): Added Workspace-aware lifecycle document output resolution. Relative output folders are expanded against the active Workspace root, and `projectId` can be used to derive the project directory from the Workspace SQLite `projects` record when no explicit output directory is provided.
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Passes the active `projectId` to `generate_lifecycle_docs`.

Decision:
- ICT lifecycle generated documents belong in the bound project directory under the active Workspace, not under the Tauri backend working directory.
- The previous narrow symptom was a relative path such as `都市花园254号` being interpreted as `src-tauri/都市花园254号`, which also caused Tauri dev watch rebuild/restart behavior. The fix centralizes output path resolution in the backend instead of adding frontend-only path string workarounds.

### AI Workspace Specified Project Context Routing

Created:
- [workspaceProjectRouter.ts](../src-ui/src/ai/context/workspaceProjectRouter.ts): Added deterministic Workspace project-name and template-name routing helpers for AI chat context composition.

Modified:
- [ai_context/dto.rs](../src-tauri/src/ai_context/dto.rs), [ai_context/service.rs](../src-tauri/src/ai_context/service.rs), and [ai_context/commands.rs](../src-tauri/src/ai_context/commands.rs): Added the read-only `list_ai_workspace_projects` command, returning lightweight current-Workspace project identity, saved-state existence flags, and saved template names without paths, file contents, image bytes, or full template JSON.
- [main.rs](../src-tauri/src/main.rs): Registered `list_ai_workspace_projects`.
- [aiProjectContextService.ts](../src-ui/src/services/aiProjectContextService.ts): Added typed frontend DTOs and invoke wrapper for the Workspace project index.
- [buildAiChatContext.ts](../src-ui/src/ai/context/buildAiChatContext.ts) and [types.ts](../src-ui/src/ai/context/types.ts): Extended the composer to route explicitly named projects to real `projectId` reads, load at most two specified project contexts per turn, reuse `template_detail` for one uniquely resolved specified template, inject lightweight project index context for Workspace-level list questions, and keep current-page draft overlay scoped to its bound project.
- [AiChatPanel.tsx](../src-ui/src/components/ai/AiChatPanel.tsx): Added system rules for specified project context, ambiguity handling, projectId-based official reads, multi-project separation, and draft isolation.

Tests:
- Added a Rust unit test covering the lightweight Workspace project index flags and saved template-name metadata.

Decision:
- Project names are only current-turn routing hints inside the active Workspace; official data reads continue to use real `projectId` values through `build_ai_project_context`.
- The composer does not guess when project names are duplicate, ambiguous, or absent. It emits warnings and asks the model to request clarification rather than falling back to another project.
- Broad Workspace questions use the lightweight index and do not trigger full project/template reads across all projects.
- This phase does not add AI writes, RAG, embedding search, file full-text reads, automatic image loading, financial logic changes, or cross-Workspace queries.

## 2026-05-31

### AI Template Detail Context and Controlled Vision Assets

Created:
- [templateAssetSelection.ts](../src-ui/src/ai/templateAssetSelection.ts): Added a lightweight cross-window bridge for explicit template image analysis selections, carrying only `projectId`, `templateId`, `assetId`, and field metadata.

Modified:
- [ai_context/dto.rs](../src-tauri/src/ai_context/dto.rs), [ai_context/service.rs](../src-tauri/src/ai_context/service.rs), and [ai_context/commands.rs](../src-tauri/src/ai_context/commands.rs): Added `template_detail` support to the read-only AI Project Context Service, including saved template field sanitization, asset metadata summaries, and `load_ai_template_asset` for controlled vision reads.
- [main.rs](../src-tauri/src/main.rs): Registered the `load_ai_template_asset` command.
- [aiProjectContextService.ts](../src-ui/src/services/aiProjectContextService.ts): Added frontend DTOs for template detail and controlled template asset image loading.
- [buildAiChatContext.ts](../src-ui/src/ai/context/buildAiChatContext.ts): Requests `template_detail` only when the active ICT AI context is a template editing context, and keeps template saved detail separate from draft overlay.
- [AiChatPanel.tsx](../src-ui/src/components/ai/AiChatPanel.tsx), [AiRuntime.ts](../src-ui/src/ai/AiRuntime.ts), [AiInputBox.tsx](../src-ui/src/components/ai/AiInputBox.tsx), [ImageAttachmentPreview.tsx](../src-ui/src/components/ai/ImageAttachmentPreview.tsx), and [MessageBubble.tsx](../src-ui/src/components/MessageBubble.tsx): Reused the existing `image_url` multimodal path for explicitly selected template assets, resolving asset bytes through the backend only at send time.
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Publishes current `projectId` and `selectedTemplate` into the template AI context payload and adds explicit AI analysis actions to template image thumbnails.

Decision:
- Template detail context is read on demand for the current specified template only; ordinary project questions continue to use template summaries.
- Saved template content comes from Workspace SQLite (`project_template_states`, with `project_settings` fallback). Unsaved template edits remain a separate frontend draft overlay.
- Template image contents are never auto-loaded. Only explicitly selected images become current-turn vision attachments after backend project/asset ownership validation.
- The implementation does not add AI writes, auto-fill, RAG, embeddings, document full-text reads, unselected image reads, or financial calculation changes.

### AI Project Context Chat Integration

Created:
- [src-ui/src/ai/context/buildAiChatContext.ts](../src-ui/src/ai/context/buildAiChatContext.ts), [buildDraftOverlay.ts](../src-ui/src/ai/context/buildDraftOverlay.ts), and [types.ts](../src-ui/src/ai/context/types.ts): Added a lightweight frontend context composer that loads saved official project context from the backend AI Project Context Service and conditionally builds an unsaved frontend draft overlay from dirty page state.

Modified:
- [AiChatPanel.tsx](../src-ui/src/components/ai/AiChatPanel.tsx): Calls the context composer on every message send and passes layered saved/draft/warning context nodes to the existing `PromptAST`/`AiRuntime` streaming path.
- [AiChatPanel.tsx](../src-ui/src/components/ai/AiChatPanel.tsx): Resets the streaming parser before inserting a new assistant placeholder and overwrites the active placeholder from the current parser output, preventing the previous AI reply from appearing while a new response is still thinking.
- [useWorkspaceStore.ts](../src-ui/src/store/useWorkspaceStore.ts): Clears workspace-scoped active project/scheme identity from project, navigation, and legacy ICT local storage state whenever the active workspace is cleared or changed.
- [ProjectBoard.tsx](../src-ui/src/views/ProjectBoard.tsx) and [useAiContextStore.ts](../src-ui/src/store/useAiContextStore.ts): Project Board now replaces its AI context snapshot with the active workspace ID, project count, lightweight project card list, and selected project summary so AI can answer workspace-level board questions after switching workspaces.
- [buildAiChatContext.ts](../src-ui/src/ai/context/buildAiChatContext.ts), [buildDraftOverlay.ts](../src-ui/src/ai/context/buildDraftOverlay.ts), [useNavigationStore.ts](../src-ui/src/store/useNavigationStore.ts), and [useProjectStore.ts](../src-ui/src/store/useProjectStore.ts): Refresh the latest persisted navigation/current-project/ICT active project identity at message-send time so a floating AI window opened before project selection can still call `build_ai_project_context` for the current ICT-bound project.
- [PromptRenderer.ts](../src-ui/src/ai/PromptRenderer.ts): Renamed rendered context layers to distinguish Workspace SQLite official state from current unsaved draft overlay and loading notes.
- [aiContextKeys.ts](../src-ui/src/utils/aiContextKeys.ts), [App.tsx](../src-ui/src/App.tsx), and [ProjectBoard.tsx](../src-ui/src/views/ProjectBoard.tsx): Added Project Board AI context key support so project-detail dirty edits can be treated as page-level draft state.
- [useIctCalculations.ts](../src-ui/src/hooks/useIctCalculations.ts) and [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Added active `projectId` markers to frontend AI context payloads for draft/project consistency checks.
- [calculator.rs](../src-tauri/src/benefit/calculator.rs): Updated existing Rust unit-test fixture construction with `project_background: None` so tests compile after the earlier model field addition. This does not change calculation logic.

Decision:
- AI chat now treats Workspace SQLite data as the only saved official project state. Zustand/localStorage page snapshots are used only as unsaved draft overlays when current dirty scopes match the current project/page.
- Each chat turn owns a fresh streaming parser lifecycle; prior assistant text must remain only in its completed message and must not seed the new pending assistant bubble.
- Workspace-level Project Board summaries are allowed as current page context for list/count questions, but they remain read-only prompt context and do not replace project-level official SQLite detail retrieval.
- Active project identity may be refreshed from persisted frontend navigation/current-project selection at send time, but that identity is used only to request read-only Workspace SQLite context; local draft state is not promoted to saved official data.
- Context loading failures degrade into prompt warnings and must not block chat streaming.
- Draft overlays are sanitized to omit base64/data URL previews and absolute paths, truncate large content, and avoid reading any image/document binaries.
- This phase remains read-only and does not implement AI writes, saves, patch application, RAG, embeddings, file full-text summaries, image analysis, scans, repairs, or financial recalculation.

### AI Project Context Service

Created:
- [ai_context/mod.rs](../src-tauri/src/ai_context/mod.rs), [ai_context/dto.rs](../src-tauri/src/ai_context/dto.rs), [ai_context/service.rs](../src-tauri/src/ai_context/service.rs), and [ai_context/commands.rs](../src-tauri/src/ai_context/commands.rs): Added a read-only backend service and Tauri command `build_ai_project_context` for structured project-level AI context retrieval from the active Workspace SQLite database.
- [aiProjectContextService.ts](../src-ui/src/services/aiProjectContextService.ts): Added typed frontend invoke wrapper for later AI integration.

Modified:
- [main.rs](../src-tauri/src/main.rs): Registered the new `build_ai_project_context` command.

Decision:
- AI project context is sourced only from persisted `.lamber.sqlite` state through `WorkspaceRuntime`, not from frontend draft state or user-supplied paths.
- The service is strictly read-only: no database writes, folder scans, repairs, document full-text reads, image binary/base64 reads, or prompt injection are performed in this phase.
- Template and file contexts are summaries/metadata only. Template image assets expose counts, not absolute paths or binary content.

### Project Background, Collection/Payment and IT/CT Content Sync in Template Forms

Modified:
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Added "IT服务内容" and "CT服务内容" textareas to the 《立项签批表》 (Project Approval/Sign-off Form) section. Handled variables resolution in `handleGenerate` to prioritize these overrides with fallback to original values.

Decision:
- Allow users to override IT and CT service contents in the Sign-off Form template configuration.
- The input boxes display original default values when empty/unmodified and correctly fallback to defaults if cleared.

### Project Background and Collection/Payment Methods Synchronization in Template Forms

Modified:
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Added "项目背景" (project background), "收入侧收款方式" (revenue collection method) and "支出侧付款方式" (expenditure payment method) fields to the 《立项签批表》 (Project Approval/Sign-off Form) section. They are bound directly to the shared states.

Decision:
- Enable immediate, reactive dual-direction synchronization of the template form configuration parameters (Project Background, Collection & Payment methods) between the template configuration view and the global variables.
- Changes are fully tracked and persisted in the database state upon saving.

### Project Background Synchronization in Template Forms

Modified:
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Added a "项目背景" (project background) textarea configuration field to the 《立项签批表》 (Project Approval/Sign-off Form) section. It is bound directly to the shared `projectBackground` state.

Decision:
- Enable immediate, reactive dual-direction synchronization of the Project Background content between the Project Parameters form page and the Sign-off Form template configuration. Modifying either updates the other in real time.
- Changes are fully tracked and persisted in the database lifecycle state upon saving.

### Workspace Management Card Interaction Fix

Modified:
- [DataManagement.tsx](../src-ui/src/views/DataManagement.tsx): Made associated Workspace cards locally selectable without opening the Workspace, isolated inline open/reveal/unlink button events, and routes the explicit open action into the Project Board after the Workspace is active.

Decision:
- The Workspace Management tab should not reorder cards on plain selection. Card clicks only highlight the chosen record; opening a Workspace remains an explicit button action.
- The Workspace Management open action is an entry action, not just a background switch; successful opening should land in the Workspace's project board.

### Workspace Management UI Separation

Modified:
- [DataManagement.tsx](../src-ui/src/views/DataManagement.tsx): Moved associated Workspace records into a dedicated "Workspace Management" tab and restyled them as card-based workspace selectors matching the Project Board workspace picker.

Decision:
- Global Workspace association management should be visually and logically separate from current Workspace backup/health maintenance. The management tab handles local association records; the maintenance tab handles operations on the active Workspace.

## 2026-05-30

### Global Workspace Management

Modified:
- [workspace.rs](../src-tauri/src/workspace.rs): Added commands to unlink a remembered Workspace from local config and to close the current Workspace while clearing `lastOpenedWorkspacePath`.
- [main.rs](../src-tauri/src/main.rs): Registered the new Workspace management commands.
- [workspaceService.ts](../src-ui/src/utils/workspaceService.ts) and [useWorkspaceStore.ts](../src-ui/src/store/useWorkspaceStore.ts): Added frontend wrappers/store actions for unlinking and closing Workspaces.
- [DataManagement.tsx](../src-ui/src/views/DataManagement.tsx): Added an associated Workspace list with open, reveal, and unlink actions.

Decision:
- Workspace unlinking is a local association operation only. It removes the app's recent-workspace reference but never deletes the physical Workspace folder, database, or project files. If unlinking the active Workspace, the UI must run the dirty guard first and the backend must release the active runtime/database connection.

### Workspace Import Argument Guard

Modified:
- [workspaceMaintenanceService.ts](../src-ui/src/services/workspaceMaintenanceService.ts): Changed `importWorkspace` to accept explicit arguments instead of an options object.
- [WorkspaceGate.tsx](../src-ui/src/components/workspace/WorkspaceGate.tsx) and [DataManagement.tsx](../src-ui/src/views/DataManagement.tsx): Updated import calls to pass `openAfterImport` and `conflictStrategy` directly.
- [workspace_maintenance.rs](../src-tauri/src/workspace_maintenance.rs): Parses `openAfterImport` as JSON and normalizes boolean or nested-object payloads before building import options. Missing or malformed nested booleans now default to `false`.

Decision:
- Import should not fail at the Tauri argument decoding layer when a stale or malformed frontend payload passes an object where a boolean is expected. The command now reaches backend validation and returns controlled errors.

### Workspace Export Reveal Target

Modified:
- [DataManagement.tsx](../src-ui/src/views/DataManagement.tsx): After Workspace export, opens the generated archive's containing directory rather than passing the archive file path as the folder target.
- [workspace_maintenance.rs](../src-tauri/src/workspace_maintenance.rs): Adjusted Windows Explorer file reveal arguments to use `/select,PATH` without embedded quote characters.

Decision:
- Export completion should take users directly to the `.lamber.zip` output directory. File selection remains a backend capability, but the export action itself only needs to open the containing folder.

### Windows Workspace System Hidden Attributes

Modified:
- [workspace.rs](../src-tauri/src/workspace.rs): Added Windows hidden-attribute helpers and applies them to Workspace system entries when workspaces are inspected, opened, created, initialized, or imported.
- [workspace_maintenance.rs](../src-tauri/src/workspace_maintenance.rs): Marks `.backups`, `.exports`, and `.projects` hidden when maintenance flows create or repair those directories.
- [assets.rs](../src-tauri/src/project_files/assets.rs): Marks `.projects` hidden when template asset storage creates the internal asset sandbox.

Decision:
- Dot-prefixed names remain the portable workspace format, but Windows Explorer needs the Hidden attribute to hide those files when hidden items are disabled.

### Workspace Initialization Nonblocking Scan

Modified:
- [workspace.rs](../src-tauri/src/workspace.rs): Existing-folder Workspace initialization now drops the SQLite transaction guard and spawns project folder scanning / automatic Excel import as a background task after the Workspace is opened.
- [commands.rs](../src-tauri/src/project_files/commands.rs): Added a workspace-root/database-context helper for automatic Excel import so background initialization scans do not need to re-read the current runtime workspace.

Decision:
- Initialization must not block the modal on file scanning or Excel parsing. The Workspace opens after project records are created; scan/import remains best-effort and runs separately.

### Workspace Import Flat Arguments

Modified:
- [workspace_maintenance.rs](../src-tauri/src/workspace_maintenance.rs): `import_workspace` now accepts flat Tauri IPC arguments (`openAfterImport`, `conflictStrategy`, `destinationName`) and no longer parses a nested `options` value.
- [workspaceMaintenanceService.ts](../src-ui/src/services/workspaceMaintenanceService.ts): Flattens import options before invoking the backend command instead of sending an `options` object.
- [workspace_maintenance.rs](../src-tauri/src/workspace_maintenance.rs): Import destination resolution now treats the selected folder as the parent directory and creates `{selectedFolder}/{workspaceName}` by default.

Decision:
- The target folder selected during import is not the workspace root itself unless explicitly named through `destinationName`; it is the directory into which the imported workspace folder is created.
- Because Lamber has not been publicly released, old import IPC compatibility is intentionally removed to keep the command contract simple.

### Workspace Backup List Cleanup UI

Modified:
- [DataManagement.tsx](../src-ui/src/views/DataManagement.tsx): Added single-backup deletion and full-list clear actions to the Workspace backup list, backed by the existing `delete_workspace_backup` command.

Decision:
- Backup cleanup is a file-level maintenance action for `.backups` entries only. It does not modify or compact the active Workspace database.

### Standalone Hub Module Cleanup

Removed:
- [BenefitTool.tsx](../src-ui/src/views/BenefitTool.tsx): Removed the standalone Investment Benefit Analysis view.
- [DocfillTool.tsx](../src-ui/src/views/DocfillTool.tsx): Removed the standalone Document Material Production view.

Modified:
- [App.tsx](../src-ui/src/App.tsx): Removed Hub cards and routes for the retired standalone modules.
- [useNavigationStore.ts](../src-ui/src/store/useNavigationStore.ts): Reduced `ViewType` to `hub`, `project_board`, `ict_lifecycle`, and `data_management`.
- [aiContextKeys.ts](../src-ui/src/utils/aiContextKeys.ts), [aiContextSerializer.ts](../src-ui/src/utils/aiContextSerializer.ts), and [AiChatPanel.tsx](../src-ui/src/components/ai/AiChatPanel.tsx): Removed active AI scopes and quick actions for the retired modules while filtering legacy stored scopes.
- [main.rs](../src-tauri/src/main.rs): Unregistered standalone benefit batch/template and document-generation commands.
- [calculator.rs](../src-tauri/src/benefit/calculator.rs), [excel.rs](../src-tauri/src/benefit/excel.rs), [models.rs](../src-tauri/src/benefit/models.rs), and [docfill.rs](../src-tauri/src/docfill.rs): Removed code used only by the standalone modules while keeping shared ICT lifecycle calculation, Excel import, template listing, and lifecycle document generation code.
- [workspace_maintenance.rs](../src-tauri/src/workspace_maintenance.rs): Ignores stale app-level module paths for retired `benefit_tool` and `docfill_tool` when listing external workspace paths.
- [Cargo.toml](../src-tauri/Cargo.toml): Removed the no-longer-used `rust_xlsxwriter` dependency.

Decision:
- Retire only the two Hub modules requested by the user. Project Board, ICT Lifecycle, Data Management, template forms, Excel import, benefit schemes/snapshots, and shared document generation remain in scope and active.

### Workspace Health Repair Path Fallback Fix

Modified:
- [workspace_maintenance.rs](../src-tauri/src/workspace_maintenance.rs): Regenerating missing `project.json` now resolves the project directory using `relative_path`, `linked_folder_relative_path`, `folder_path`, then `folder_name`, matching the health-check path resolution used to report the missing manifest.
- [workspace_maintenance.rs](../src-tauri/src/workspace_maintenance.rs): External `module_path:*` health items are now repairable by resetting the module base path to `.projects/modules/{moduleId}` inside the current Workspace and updating the app-level module config.

Decision:
- Older or imported flat workspace projects that only have `folder_path` should remain repairable without requiring a separate path-conversion step first.
- Repairing a module path does not move or delete files from the old external path; it only changes where the module reads/writes templates and output going forward.

## 2026-05-29

### Workspace Refactoring Phase 4: Local Portability, Backup, Restore, and Health

Created:
- [workspace_maintenance.rs](../src-tauri/src/workspace_maintenance.rs): Workspace maintenance commands for daily/manual SQLite backups, backup listing/deletion/restore, `.lamber.zip` export/import/validation, read-only workspace health checks, repairable issue execution, external path listing, dry-run internal absolute path conversion, and file-manager reveal.
- [workspaceMaintenanceService.ts](../src-ui/src/services/workspaceMaintenanceService.ts): Frontend service wrapper for workspace maintenance IPC commands.

Modified:
- [workspace.rs](../src-tauri/src/workspace.rs): Exposed workspace manifest/database path helpers for maintenance flows, added reserved workspace entry detection, added safe database connection closing for restore, changed recent workspace updates to support import-without-open, and triggers a best-effort daily SQLite backup when a workspace opens.
- [commands.rs](../src-tauri/src/benefit/commands.rs): Rejects new flat workspace project folders that collide with reserved workspace entries and skips reserved entries during root-level workspace inspection.
- [assets.rs](../src-tauri/src/project_files/assets.rs): Resolves project folders relative to the active workspace for template asset storage and retrieval, writes new internal assets inside the current workspace, and keeps AppData lookup only as a legacy read fallback.
- [DataManagement.tsx](../src-ui/src/views/DataManagement.tsx): Added a Workspace Maintenance tab for current workspace metadata, manual backup, backup restore, export/import, health check results, repair buttons, external path listing, and internal path conversion.
- [WorkspaceGate.tsx](../src-ui/src/components/workspace/WorkspaceGate.tsx): Added `.lamber.zip` import entry when no workspace is active.
- [main.rs](../src-tauri/src/main.rs): Registered workspace maintenance commands.

Decisions:
- Direct folder copy remains the primary workspace migration path. Export/import is a convenience layer and preserves `workspaceId` by default.
- `.lamber.zip` archives use the workspace root as the archive root and include `export-manifest.json`; no random top-level folder is added. Import still accepts archives with one wrapper folder for compatibility.
- `.backups` and `.exports` are excluded from export by default. The pre-export database backup protects the live workspace and is not itself included unless backups are explicitly selected.
- `run_workspace_health_check` is read-only. Repairs and internal absolute path conversion are explicit user actions and create a database backup before modification.
- Backup restore releases the active SQLite connection before replacing `.lamber.sqlite`, then reopens the workspace and attempts rollback/reopen if replacement fails.
- `project_roots` representing external roots are only checked and reported; automatic path conversion does not rewrite them.

## 2026-05-28

### Workspace Refactoring Phase 3: Domain Save Boundaries and Dirty State

Created:
- [project_state/mod.rs](../src-tauri/src/project_state/mod.rs): Workspace-scoped project state commands for project detail, lifecycle state, cashflow state, benefit analysis, template states, template assets listing, and full project state loading.
- [useSaveStore.ts](../src-ui/src/store/useSaveStore.ts): Global dirty scope store with registered save handlers, context checks, partial failure handling, and Ctrl/Command+S integration.
- [domainSaveService.ts](../src-ui/src/services/domainSaveService.ts): Frontend domain service wrapping project detail, lifecycle, cashflow, benefit analysis, template state, and full-state commands.
- [GlobalSaveButton.tsx](../src-ui/src/components/GlobalSaveButton.tsx), [useGlobalSaveShortcut.ts](../src-ui/src/hooks/useGlobalSaveShortcut.ts), and [useUnsavedChangesGuard.ts](../src-ui/src/hooks/useUnsavedChangesGuard.ts).

Modified:
- [db.rs](../src-tauri/src/db.rs): Added schema version 5 with `project_lifecycle_states`, `project_cashflow_states`, and `project_template_states`; normalized `project_template_assets` creation for fresh databases and added `template_id` compatibility.
- [main.rs](../src-tauri/src/main.rs): Registered project-state Tauri commands.
- [service.rs](../src-tauri/src/benefit/service.rs): Prevented `update_project` from inserting a missing project into the current workspace database.
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Registered lifecycle and cashflow save handlers, loaded new project full state with legacy snapshot fallback, displayed unsaved status, and guarded project/template navigation.
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Routed template form persistence through the template domain state while preserving legacy fallback and asset references.
- [ProjectBoard.tsx](../src-ui/src/views/ProjectBoard.tsx), [WorkspaceGate.tsx](../src-ui/src/components/workspace/WorkspaceGate.tsx), and [WorkspaceHeader.tsx](../src-ui/src/components/WorkspaceHeader.tsx): Integrated global save, project/workspace switching guards, and project-detail dirty handling.

Decisions:
- Current ICT editing state no longer depends on `benefit_snapshots` as its only persistence boundary.
- Benefit方案 buttons still save schemes and snapshots, but global save / Ctrl+S save current editor state through domain handlers.
- Template field values and template assets remain separate from lifecycle/cashflow state; legacy `project_settings` template payloads remain readable and are kept as compatibility mirrors during saves.

Phase 3B updates:
- `useSaveStore` save handlers now return explicit `savedScopes`; missing handlers, partial failures, and workspace/project switches no longer clear unrelated dirty state.
- `TemplateForms` propagates template save failures to the global save handler so `template-forms` cannot be marked saved unless `saveTemplateState` succeeds.
- ICT lifecycle header shows project status and save status globally, while project, workspace, and template switches are guarded before navigation.
- ICT lifecycle project loading now prefers `project_lifecycle_states` current editor state over benefit scheme snapshots even when a default scheme id is present; snapshots remain a fallback for old projects.

### Workspace Structure Flattening, Hidden System Files, and Auto Migration

Modified:
- [workspace.rs](../src-tauri/src/workspace.rs): Changed manifest and database filenames to `.lamber.workspace.json` and `.lamber.sqlite` (prefixed with dot to make them hidden files on macOS/Linux). Changed backups and exports folder names to `.backups` and `.exports`. Removed the projects subfolder layer to place all project subdirectories directly at the workspace root. Added `migrate_legacy_workspace_files` to automatically rename existing visible files/folders (`lamber.workspace.json`, `lamber.sqlite`, `backups`, `exports`) to dot-prefixed hidden ones on workspace inspection and loading.
- [commands.rs](../src-tauri/src/benefit/commands.rs): Adjusted `create_project_in_workspace` and `inspect_workspace_projects` to create and scan folders directly under the workspace root, and to ignore hidden dot-prefixed system folders.
- [service.rs](../src-tauri/src/project_files/service.rs): Updated files creation and relative path references in `add_project_file` to skip the `projects/` subfolder.
- [assets.rs](../src-tauri/src/project_files/assets.rs): Changed sandboxed fallback files folder to `.projects/` (hidden) to keep the workspace root clean of database project IDs.

Decisions:
- Flattered workspace layout so that only the actual, user-visible project directories (e.g. `项目A`, `项目B`) are visible directly under the workspace root.
- Hidden all metadata, database, and cache files/directories behind Unix dot-prefixes (`.backups`, `.exports`, `.projects`, `.lamber.sqlite`, `.lamber.workspace.json`).
- Automatically detect and rename legacy visible workspace configuration files and directories when opened to ensure backward compatibility and seamless updates.

### Automatic Excel Calculations Import on Scan and Initialization

Modified:
- [commands.rs](../src-tauri/src/project_files/commands.rs): Implemented `auto_import_excel_if_needed` to search for files starting with "效益分析表" and ending with ".xlsx" / ".xls", sorting by modification date descending to pick the newest, and parsing/importing using the `ProjectService`. Updated `scan_project_folder` and `bind_project_folder` commands to trigger this check.
- [workspace.rs](../src-tauri/src/workspace.rs): Updated `initialize_workspace_from_existing_directory` to keep track of successfully imported project IDs, run directory scans, and call the auto-import logic for all projects after committing.

Decisions:
- Enforced a safety boundary: only trigger auto-import when the target project currently has 0 schemes to protect existing user work.
- Decoupled files service scanning from the benefit service calculations: `auto_import_excel_if_needed` resolves relative workspace paths against the active workspace root and passes them to the parsing module.

### Excel Importing Workspace Relative Path Resolution Fix

Modified:
- [excel.rs](../src-tauri/src/benefit/excel.rs): Updated the `parse_benefit_excel` command to accept the active `WorkspaceRuntime` state and resolve workspace-relative paths (e.g. scanned files in workspace directories) against the active workspace root path prior to performing existence checks and opening the workbook.

Decisions:
- Passed `State<'_, Arc<WorkspaceRuntime>>` to `parse_benefit_excel`.
- Checked and resolved non-absolute paths relative to the current workspace root path inside the command handler before spawning the blocking task to ensure proper lifetime handling of borrowed state objects.

### Database Migration and Workspace Initialization Robustness Fixes

Modified:
- [db.rs](../src-tauri/src/db.rs): Fixed a database migration bug on fresh databases. Previously, fresh databases were initialized with a schema version of `'2'`, which triggered the version 3 and 4 migrations (including `ALTER TABLE projects ADD COLUMN folder_name TEXT`) and failed with `duplicate column name: folder_name` since the table was already created using the latest schema. We added a column check on `projects` (`folder_name`) to detect fresh databases and set the initial schema version to `'4'` so migrations are skipped cleanly.
- [workspace.rs](../src-tauri/src/workspace.rs): Added robust error cleanups inside `initialize_workspace_from_existing_directory`. If any step fails during initialization (such as folder scanning or database transaction commits), the command deletes any partially created workspace manifest (`lamber.workspace.json`) and database (`lamber.sqlite`) to prevent leaving the folder in a corrupted, uninitializable state. It also automatically cleans up orphan manifests from failed runs on entry.

### Workspace Initialization & Subdirectories Bulk Import

Created/Modified:
- [workspace.rs](../src-tauri/src/workspace.rs): Implemented plain directories inspection logic, returning `importablePlainDirectory` when eligible subdirectories are found. Added options deserialization and `initialize_workspace_from_existing_directory` command which sets up directories, writes workspace manifest, establishes database transaction, deduplicates entries, writes/補足 `project.json` manifests and creates standard nested directories (`assets`, `documents`, `analyses`).
- [main.rs](../src-tauri/src/main.rs): Registered the `initialize_workspace_from_existing_directory` invoke handler.
- [workspaceService.ts](../src-ui/src/utils/workspaceService.ts): Added the status and new initialize command wrapper.
- [useWorkspaceStore.ts](../src-ui/src/store/useWorkspaceStore.ts): Exposed store method `initializeWorkspaceFromExisting`.
- [WorkspaceGate.tsx](../src-ui/src/components/workspace/WorkspaceGate.tsx): Unified workspace selection pickers under local inspection, rendering a checklist selection modal showing candidate project subdirectories and optional parameter options when opening plain directories.

Decisions:
- Standardized candidate directories to exclude backups, exports, standard projects, .git, and common target build directories.
- Stored project paths relative to workspaceRoot, keeping folders at the workspace root rather than moving them under `projects/`.
- Ensured transaction rollback and duplicates check to prevent workspace half-state or duplicate project database rows.

### Navigation Defaults and Back Button Adjustments

Modified:
- [useNavigationStore.ts](../src-ui/src/store/useNavigationStore.ts): Adjusted the initial view logic to always default `currentView` to `"hub"` upon application launch, avoiding restoring last active view directly (which bypasses the Hub and could immediately trigger a workspace gate or board view).
- [WorkspaceGate.tsx](../src-ui/src/components/workspace/WorkspaceGate.tsx): Allowed the back button to show even when there is no current workspace loaded (removed `currentWorkspace` check).
- [App.tsx](../src-ui/src/App.tsx): Passed the `onBack` handler and custom label to `WorkspaceGate` in the data management view, and wrapped the view in a header when `isWorkspaceReady` is false, ensuring the '返回集市' button is always consistently visible.
- [ProjectBoard.tsx](../src-ui/src/views/ProjectBoard.tsx): Passed the `onBack` handler and custom label to `WorkspaceGate` in the project board view when `isWorkspaceReady` is false. Standardized both header back buttons to use the text-button `← 返回集市` styling.
- [DataManagement.tsx](../src-ui/src/views/DataManagement.tsx): Standardized the header back button to the text-button `← 返回集市` style.

### Workspace Refactoring Phase 2: Project Board & Workspace Binding

Created:
- [useProjectStore.ts](../src-ui/src/store/useProjectStore.ts): Zustand store managing active currentProject context with local storage recovery and workspace linkage.

Modified:
- [workspace.rs](../src-tauri/src/workspace.rs): Added workspace paths normalization, verification functions, safe folder naming, standard subdirectories creation, and workspace root auto-healing logic on workspace relocation.
- [db.rs](../src-tauri/src/db.rs): Expanded `projects` table columns with migration check to version 4 (added `folder_name`, `relative_path`, `progress`, `deadline`, `linked_folder_type`, `linked_folder_relative_path`, `linked_folder_external_path`).
- [models.rs (benefit)](../src-tauri/src/benefit/models.rs): Extended Rust `Project` model with Phase 2 workspace mapping attributes.
- [repository.rs (benefit)](../src-tauri/src/benefit/repository.rs): Added DB mapping of extended project attributes for SQLite database repository.
- [commands.rs (benefit)](../src-tauri/src/benefit/commands.rs): Implemented workspace-scoped project operations: `create_project_in_workspace` (with directory structure creation and `project.json` manifest writing), `list_workspace_projects`, and `inspect_workspace_projects`.
- [service.rs (project_files)](../src-tauri/src/project_files/service.rs): Updated files scanning and sandboxing paths mapping to store documents relative to the workspace projects folder and to prevent delete-on-unbind behavior.
- [commands.rs (project_files)](../src-tauri/src/project_files/commands.rs) & [health.rs](../src-tauri/src/project_files/health.rs): Passed the active workspace root to file service calls.
- [ProjectBoard.tsx](../src-ui/src/views/ProjectBoard.tsx): Rebuilt creation flow using standard workspace paths, bound cards to missing directory warnings, removed old manual folder selection modals, and integrated project context selection.
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Synced global project state store context and added support for saving free calculations into any existing workspace project.

Decisions:
- Standardized project folders creation inside `workspaceRoot/{safeProjectName}/` with `assets/`, `documents/`, `analyses/` subfolders and a backup `project.json` manifest.
- SQLite remains the primary source of truth, with `project.json` serving as redundant metadata.
- Internal folders bound to projects are stored as relative workspace paths to guarantee portable workspaces, while external folders trigger visual alerts warning users of non-portable linkages.
- Unbinding local folders clears metadata and scanner references in the DB, but keeps actual disk files untouched.

### Workspace Runtime Foundation

Created:
- [workspace.rs](../src-tauri/src/workspace.rs): Workspace manifest model, UUID v4 workspace creation, current workspace runtime, database connection binding, recent workspace updates, last-opened restore, and directory-state inspection.
- [useWorkspaceStore.ts](../src-ui/src/store/useWorkspaceStore.ts): Frontend global workspace state.
- [WorkspaceGate.tsx](../src-ui/src/components/workspace/WorkspaceGate.tsx): Gate shown when database-backed modules are accessed without an open workspace.
- [workspaceService.ts](../src-ui/src/utils/workspaceService.ts): Frontend Workspace IPC wrapper and error parsing.

Modified:
- [main.rs](../src-tauri/src/main.rs): Removed startup initialization of AppData `projects_store.db`; registered Workspace commands; attempts to restore `lastOpenedWorkspacePath`.
- [config_manager.rs](../src-tauri/src/config_manager.rs): Added `recentWorkspaces` and `lastOpenedWorkspacePath` to local AppConfig.
- Project, file, root, health, relocation, import, template asset, and docfill commands now obtain the active SQLite connection from `WorkspaceRuntime`.
- [App.tsx](../src-ui/src/App.tsx): Keeps Hub focused on module selection and blocks database-backed non-board modules behind WorkspaceGate.
- [ProjectBoard.tsx](../src-ui/src/views/ProjectBoard.tsx): Moved workspace selection into the Project Board flow, making "Project Workspace" the first layer before project cards are loaded. Switching workspaces now opens a workspace overview instead of immediately opening a folder picker.
- [WorkspaceGate.tsx](../src-ui/src/components/workspace/WorkspaceGate.tsx): Refactored into a card-based workspace overview using locally recorded recent workspaces, with current-workspace marking and explicit open/create actions.
- [WorkspaceGate.tsx](../src-ui/src/components/workspace/WorkspaceGate.tsx) and [ProjectBoard.tsx](../src-ui/src/views/ProjectBoard.tsx): Clicking the current workspace card now directly closes the workspace overview and returns to the existing project board. Only selecting a different workspace performs a workspace switch.

Decisions:
- `lamber.workspace.json` stores only workspace identity metadata. Recent workspace history remains local to the machine.
- A directory with `lamber.sqlite` but no manifest is treated as `legacySuspected`; the app does not overwrite or migrate it in this phase.
- Opening or creating a workspace automatically registers the workspace root as a default project root when missing, so project folders under the workspace do not trigger the legacy "register as new root" prompt.
- Workspace selection belongs to the Project Board workflow. Hub should open the Project Board module, while ProjectBoard presents the Project Workspace layer before project list operations.
- Project workspace switching should show the recorded workspace overview first. Opening an arbitrary directory remains an explicit secondary action.
- Legacy JSON/AppData migration is deferred and no longer runs automatically on startup.

## 2026-05-27

### Project Background Persistence in Calculator Snapshots

Modified:
- [models.rs (benefit)](../src-tauri/src/benefit/models.rs): Added `project_background` field to `IctInput` struct.
- [useIctCalculations.ts](../src-ui/src/hooks/useIctCalculations.ts): Serialized `project_background` inside `buildInputDataPayload` to send it with the snapshot.
- [IctLifecycle.tsx](../src-ui/src/views/IctLifecycle.tsx): Restored the `projectBackground` state inside `fillCalculatorState` when loading snapshot properties.

### Project Nested Assets Folder Renaming to Project Name Suffix

Modified:
- [assets.rs (project_files)](../src-tauri/src/project_files/assets.rs): Renamed the bound project folder's template asset folder from `.lamber` to `{project_name}-图片`. Added `sanitize_folder_name` helper to clean folder names, updated `get_project_folder_info_from_db` to retrieve project names, and implemented self-adaptive fallback path checks to support legacy `.lamber/assets/` images and project renames.

### Project Template Data & Asset Separation Persistence

Created:
- [assets.rs (project_files)](../src-tauri/src/project_files/assets.rs): Core sandboxed image uploads, size (<= 20MB) and MIME type checks, soft-delete, and orphan asset garbage collector.

Modified:
- [db.rs](../src-tauri/src/db.rs): Added `project_template_assets` table structure and transaction-wrapped Version 3 schema upgrade. Fix connection mutable borrow in init.
- [repository.rs (benefit)](../src-tauri/src/benefit/repository.rs): Added `get_project_setting` and `save_project_setting` to `ProjectRepository` and its Sqlite, Json, and Dual implementations.
- [commands.rs (project_files)](../src-tauri/src/project_files/commands.rs): Registered six new Tauri commands for managing project settings and assets.
- [main.rs](../src-tauri/src/main.rs): Registered the six new Tauri command handlers in the application runtime.
- [docfill.rs](../src-tauri/src/docfill.rs): Refactored `internal_generate_docx` to load sandboxed images by resolving `assetId` values using database connection validation, bypassing frontend absolute path leakage.
- [projectService.ts](../src-ui/src/utils/projectService.ts): Mapped backend template setting and asset commands to frontend APIs.
- [TemplateForms.tsx](../src-ui/src/views/TemplateForms.tsx): Implemented loading, 1s auto-saving of forms, instant paste/drop upload, and legacy base64 image migrations. Fully refactored uncontrolled form fields (including demand checklist, customer confirm, env requirements, public URL, and security detail) using controlled state binding `getBind` to eliminate component-reload reset issues.
- [default.json (capabilities)](../src-tauri/capabilities/default.json): Removed the invalid `"core:protocol:allow-asset"` capability.
- [tauri.conf.json](../src-tauri/tauri.conf.json): Configured `assetProtocol`'s scope correctly for security sandbox file loading.

Decisions:
- Stripped base64 strings from JSON objects prior to writing to the SQLite database, storing only `assetId` references.
- Forced backend-only resolution of physical image paths using project database ownership validation.
- Allowed in-memory local base64 preview rendering in React frontend to ensure instant feedback while upload and saving happen in background.

## 2026-05-26


### Path Resilience & Database Cascade Deletion Fixes

Modified:
- [service.rs (project_files)](../src-tauri/src/project_files/service.rs):
  - Updated `bind_project_folder` to extract the parent directory of `folder_path` when creating a new project root, registering the parent directory as root, and mapping the subfolder name as the project's relative subpath.
  - Updated `scan_project_folder` and `add_project_file` to dynamically save the matched `project_directories` entry prior to saving files, and only set `directory_id` if a root is matched. If rootless (absolute-only mode), set `directory_id` to `None` to prevent SQLite `FOREIGN KEY constraint failed` errors.
- [import_scanner.rs](../src-tauri/src/project_files/import_scanner.rs):
  - Corrected `directory_id` assignment during project files insertion: if the imported project doesn't match any global root, set `directory_id` to `None` instead of `Some(dir_id)`. Removed unused `PathBuf` import.
- [benefit/repository.rs](../src-tauri/src/benefit/repository.rs):
  - Refactored `save_project` in `SqliteProjectRepository` to check row existence via `SELECT EXISTS` and run `UPDATE` or `INSERT` instead of `INSERT OR REPLACE`. This prevents SQLite's delete conflict resolution from triggering `ON DELETE CASCADE` deletions on child tables (`project_directories`, `project_files`, `benefit_schemes`, `benefit_snapshots`), which was previously wiping folder bindings and scheme calculations during Excel imports.

Decisions:
- Enforced strict nullable constraint safety for `directory_id` in `project_files` when directories are rootless, avoiding invalid foreign key references in SQLite.
- Adjusted root registration to logically target parent folders, enabling automatic grouping of adjacent project folders under the same root drive or parent directory.
- Replaced `INSERT OR REPLACE` with transactional `UPDATE`/`INSERT` checks on the main `projects` table to prevent unexpected cascade deletion cascades on child tables.

## 2026-05-25

### Project Files System and Path Resilience Upgrade (Phase 2)

Created:
- [roots.rs](../src-tauri/src/project_files/roots.rs) - Global Project Roots Configuration CRUD, defaults manager, and Tauri commands.
- [health.rs](../src-tauri/src/project_files/health.rs) - Health Check Service analyzing path linkages, exists states, mismatch counts, and auto-healing directories.
- [relocation.rs](../src-tauri/src/project_files/relocation.rs) - Transactional bulk relocation service to preview and swap directories paths across drives.
- [import_scanner.rs](../src-tauri/src/project_files/import_scanner.rs) - Recursive candidate folder scanner and importer utilizing SQL transactions to bulk import subfolders as projects (auto-determining file roles).
- [DataManagement.tsx](../src-ui/src/views/DataManagement.tsx) - Unified Data Management Dashboard supporting Roots, Health Checker, and Relocator.

Modified:
- [db.rs](../src-tauri/src/db.rs) - Updated schema versions and added new tables/migrations for project_roots and directories. Escaped `"exists"` column name to resolve SQLite syntax error.
- [repository.rs (project_files)](../src-tauri/src/project_files/repository.rs) - Extended repositories (Json, SQLite, Dual) with roots lookup, directories association, and files sync. Escaped `"exists"` column name in queries.
- [service.rs (project_files)](../src-tauri/src/project_files/service.rs) - Updated `bind_project_folder` to support `force_mode` parameter, added 5-level path resolving and metadata extraction, and integrated self-healing.
- [ProjectFilesTab.tsx](../src-ui/src/components/project/ProjectFilesTab.tsx) - Handled `NOT_IN_ROOT` error during folder binding by triggering a custom modal offering to create a root or bind as absolute-only.
- [ProjectBoard.tsx](../src-ui/src/views/ProjectBoard.tsx) - Added a "批量扫描导入" button in the board header, and implemented a custom modal showing candidate folders, file roles, and conflict resolution selector (`merge`/`new`/`skip`).
- [main.rs](../src-tauri/src/main.rs) - Registered new commands and services for roots, relocation, health, and scanner.
- [relocation.rs](../src-tauri/src/project_files/relocation.rs) & [import_scanner.rs](../src-tauri/src/project_files/import_scanner.rs) - Escaped `"exists"` column name in UPDATE and INSERT queries.

Decisions:
- Standardized relative path mapping using a combination of `project_roots` + `relative_path` + `file_fingerprint` to ensure path durability across different environments and drives.
- Enforced transactional integrity during batch operations (relocation, candidate imports) using SQLite database transactions to prevent corruption or locking.
- Escaped standard SQL reserved keyword `"exists"` within SQLite table structure and query strings to prevent runtime panics during boot.
- Retained "No-Line Rule" surface styling for all new components.

## 2026-05-23

### SQLite Database Integration and Dynamic Hot-Swapping (Phase 1)

Created:
- [db.rs](../src-tauri/src/db.rs) - SQLite connection management, foreign key configurations, and creation of 12 core tables.
- [migration.rs](../src-tauri/src/migration.rs) - JSON-to-SQLite transactional database migration service, auto-backup, report generator, and Tauri commands.

Modified:
- [Cargo.toml](../src-tauri/Cargo.toml) - Added `rusqlite` dependency with `bundled` feature.
- [repository.rs (benefit)](../src-tauri/src/benefit/repository.rs) - Implemented `SqliteProjectRepository` and Arc-wrapped `DualProjectRepository` wrapper for hot-swapping.
- [repository.rs (project_files)](../src-tauri/src/project_files/repository.rs) - Implemented `SqliteProjectFileRepository` and `DualProjectFileRepository` wrapper.
- [main.rs](../src-tauri/src/main.rs) - Registered new modules, initialized database connection, set up initial repository backends, and registered Tauri migration commands.
- [App.tsx](../src-ui/src/App.tsx) - Embedded startup check for SQLite database migration and created a premium ledger-style overlay popup modal to prompt user for migration and display statistics.

Decisions:
- Standardized SQLite database as primary storage, enabling transaction management instead of concurrency-sensitive JSON file writing.
- Kept `projects_store.json` backward compatibility and implemented automatic timestamped JSON backup before database inserts.
- Utilized dynamic dual repository swapping to support hot swapping repositories without app restarts after migration or skip actions.

### Initial project state mapping

Created:
- [AGENTS.md](../AGENTS.md)
- [AI_CONTEXT.md](./AI_CONTEXT.md)
- [PROJECT_STATUS.md](./PROJECT_STATUS.md)
- [ARCHITECTURE_MAP.md](./ARCHITECTURE_MAP.md)
- [CHANGELOG_AI.md](./CHANGELOG_AI.md)

Observed:
- The project is a desktop application combining a Rust/Tauri v2 backend and React 18 / Zustand frontend.
- Calculations (payback, NPV, selection fee, reverse-calculations) use Rust's `rust_decimal` to prevent float rounding errors.
- Real-time serialization of form states to HSL-formatted AI prompts allows context-aware chat capabilities.
- Local folder scanning maps folder structures into project file entities, with safety toggles between sandbox copies and raw links.
- 0-tolerance financial reconciliation blocks invalid user states during workflow transitions.

Decisions:
- Standardized UI borders are replaced by color surface adjustments to fulfill the design guidelines of "The Architectural Ledger".
- Navigational routes track entry sources (`entrySource`) to support bidirectional back-navigation without context degradation.
- AI direct write capabilities are locked behind user interaction boundaries.

Risks:
- Coordinates for Excel coordinate back-filling (`excel.rs` and `parse_benefit_excel`) are hardcoded. Edits to templates will break structural mappings.
- The 0-tolerance reconciliation check may experience rounding edge cases when converting user inputs.
- Concurrency during atomic JSON writes (`projects_store.json`) might result in race conditions.

Open questions:
- Should the local JSON-based storage layer be migrated to SQLite in the next phase to improve transactional consistency?
- Do we need to support dynamic annual discount rates rather than flat project-wide rates?
## 2026-06-11

### 智算报价业务项资金计划

Created:
- `src-ui/src/features/ai-compute-quote/fundingPlans.ts`：智算计划生成、模式切换、年度编辑、归一化与一致性校验纯函数。
- `src-ui/src/features/ai-compute-quote/AiComputeFundingPlanEditor.tsx`：参考 ICT 科目计划交互的独立折叠式业务项资金计划面板。

Modified:
- 智算行项目增加可持久化的 10 年资金计划，支持第一年一次性、按项目周期平均分年和自定义年度金额。
- 公式重算会同步自动模式计划，手工模式保持用户年度输入和可见差额。
- 智算输出包增加按 `side + ictSubjectCode` 合并的年度计划，并在预览弹窗展示第 1-10 年、合计和来源项。
- 专项测试增加计划生成、差额识别、同科目年度合并和禁用/未输出/未映射过滤覆盖。

Decisions:
- 智算计划只属于报价蓝图和输出预览；ICT 科目计划继续作为正式现金流与效益指标的唯一来源。
- 计划不一致仅提示，不阻止智算蓝图保存，也不会触发 ICT 正式状态写入。
## 2026-06-11

### 智算报价三 Tab 工作区与紧凑参数面板

Modified:
- 智算主编辑区拆分为参数、收入计算项、成本计算项三个独立 Tab，每个 Tab 使用独立滚动容器并保留自身滚动位置。
- 参数卡改为两列紧凑网格，常用参数优先，默认显示 8 项并支持展开全部和收起。
- 参数用途标签与公式引用统计直接从蓝图类别和公式 Token 推导，显示收入/成本引用数量并提供具体引用项提示。
- 单变量敏感性分析归入参数 Tab，收入与成本编辑页不再连续堆叠。
- 效益预览作为右侧独立区域持续可见，并增加输出包操作入口。

Decisions:
- Tab、展开和滚动位置只属于 UI 状态，不写入项目 setting，也不影响公式、资金计划或 ICT 输出。
- 保持现有计算项编辑器和资金计划组件不变，仅重组页面布局和参数信息密度。
## 2026-06-12

### 智算输出到 ICT 显式确认写入

Created:
- `src-ui/src/features/ai-compute-quote/ictExport.ts`：构建 ICT 差异预览、合并科目输出和正式 lifecycle/cashflow payload。

Modified:
- “输出到 ICT”改为加载正式快照、显示差异、确认后真实写入，不再只弹安全提示。
- 新增事务命令 `apply_ai_compute_quote_to_ict`，同时更新项目 lifecycle 和 cashflow 状态并使旧效益指标失效。
- 正式科目资金计划支持 `ai_compute_quote` 来源和导入追踪元数据。
- ICT 导航支持从智算进入后返回同一项目和场景。
- 专项测试覆盖同科目合并、原值差异、金额/年度计划 payload、来源追踪和未映射过滤。

Decisions:
- 安全边界从“禁止写入”调整为“禁止静默写入”；用户明确确认后允许正式提交。
- 差异预览与最终 payload 使用同一个纯函数模型，避免展示结果和实际写入不一致。
- lifecycle/cashflow 必须事务提交，禁止金额和资金计划出现半成功状态。

## 2026-06-15

### 智算金额来源与 ICT 边界重构

Created:
- `src-tauri/src/intelligent_compute.rs`：智算项目状态、金额来源 CRUD、归属校验和乐观版本控制。
- `src-ui/src/services/intelligentComputeService.ts`：智算状态与金额来源 IPC 服务。
- schema v7 表 `project_intelligent_compute_states` 与 `intelligent_compute_amount_sources`。

Modified:
- 项目增加 `project_type: ict | intelligent_compute`，新建、详情转换、项目看板标识和入口按类型控制，并同步写入 `project.json.projectType`。
- v6→v7 自动识别 `ai_compute_quote::active` 历史项目并迁移公式、行项目、年度计划、映射和同步快照；旧 setting 改为只读兼容。
- 智算页面改为项目级周期/折现率和多个独立金额来源，支持空白新增、复制默认停用、重命名、启停和最后来源保护。
- 编辑停止后只保存智算数据；同步改为先展示覆盖/释放/年度计划预览，再由用户确认。
- 多个启用来源按 `side + ICT subjectCode` 聚合，导入痕迹增加 `amountSourceIds` 和复合行引用。
- 新增 `sync_intelligent_compute_to_ict`，在单一事务中校验项目类型、同步 revision、完整来源集合和来源版本，运行 Rust 正式计算并更新 ICT 与智算同步快照。
- ICT 保存移除智算反向覆盖路径，历史 `ict_override / merge_conflict` 不再影响智算公式。
- 智算到 ICT 使用显式 `IctOrigin`，ICT 顶部显示来源并可返回智算；项目/Workspace 切换会清空来源和完整模块状态，异步加载使用请求代次防止串项目。
- 测试增加 schema v7、新库默认类型、历史迁移、来源 revision 回滚、多来源聚合、停用排除、旧科目释放和复合来源痕迹。

Decisions:
- 智算负责计算过程和金额来源；ICT 只承接确认后的聚合结果并维护正式资金计划与指标。
- 项目周期和折现率属于智算项目级状态，年度金额只存在于行项目资金计划。
- 保留 `ai_compute_quote` 路由和内部计算类型作为兼容层，但用户文案和新持久化结构统一采用“智算金额来源”。

### 智算测算计算项与效益结论信息结构优化

Created:
- `BenefitConclusionSidebar.tsx`：固定只读效益结论侧栏，按结论、核心金额、收益指标、关键构成和输出操作分组。
- `CalculationItemCard.tsx`：收入/成本项统一摘要卡，展示金额、占比、ICT 状态、资金计划状态和公式摘要。
- `CalculationItemDetailTabs.tsx`：计算项向下展开后的项目编辑、计算公式、资金计划三 Tab 容器。

Modified:
- ICT 科目下拉、参与测算和参与自动同步配置统一移入“项目编辑”，不增加 ICT 独立 Tab。
- 公式编辑器继续复用原 Token 编辑能力；资金计划继续复用原生成与校验函数，并移除第二层折叠交互。
- 右侧可收起结论抽屉改为固定只读侧栏；窄屏采用纵向布局以避免横向溢出。
- 删除操作移入项目编辑，默认卡片只保留启用、金额、状态和公式摘要。

Decisions:
- 本次只调整 UI 信息结构，不修改计算、保存、ICT 自动同步、人工覆盖或资金计划校验数据流。
- 计算项展开与详情 Tab 均为本地 UI 状态，不进入蓝图持久化。

### 智算统一控制 ICT 项目折现率

Modified:
- 智算新增稳定 `discount-rate / discount_rate` 参数和醒目折现率输入框，按百分数输入并转换为 ICT 小数口径。
- 蓝图持久化升级到 Version 4；Version 1-3 项目从项目当前正式折现率强制初始化一次，后续由智算值覆盖 ICT。
- 同步同时覆盖 lifecycle 输入、lifecycle parameters、cashflow assumptions、Rust 正式计算输入和项目表 `discount_rate`。
- 智算效益结论增加 ICT 项目净现值率，直接使用 Rust 返回的 `npv_rate`。
- 专项测试覆盖折现率归一化、旧蓝图迁移、同步覆盖和 SQLite 原子事务。

Decisions:
- 项目周期和折现率由智算优先控制；其他 ICT 财务状态继续沿用现有所有权。
- 不修改 NPV 等财务公式，仅消除同一参数在 lifecycle、cashflow 和项目汇总中的数据分叉。

### 智算到 ICT 科目释放与项目周期一致性

Modified:
- 智算自动同步新增旧映射释放：业务项取消映射、删除或改映射后，原 ICT 科目金额和 10 年计划在同一正式事务中清零。
- 同一 ICT 科目仍有其他智算来源时只写入剩余来源的金额和逐年计划，避免误清整个科目。
- 成功同步快照改为增量维护；释放科目删除，正常科目更新，人工覆盖和合并冲突科目保留原比较快照。
- 修改带 `ict_override` / `merge_conflict` 的业务项映射会恢复公式控制，避免旧人工覆盖阻断新映射。
- `years` 固定为 1-10 年项目周期唯一参数，并在参数区增加独立周期入口；同步继续覆盖 ICT lifecycle、cashflow、正式输入和项目汇总年限。
- 同步明细增加计划模式、释放旧映射和当前受控/释放科目统计；导入审计增加 `releasedSubjectCodes`。
- 专项测试覆盖甲供材料改映射到其他投入、删除映射、剩余来源合并、逐年计划 `[50,40] + [40,0] = [90,40]`、周期边界和暂停覆盖快照。
- 修复历史项目缺少旧科目同步快照时的释放遗漏：从同项目、同蓝图、同场景且未被人工修改的 ICT 资金计划导入痕迹恢复旧受控科目。
- 兜底恢复跳过已人工维护和已归零科目，避免误清 ICT 人工值或重复生成释放记录。
- 使用真实项目状态复核机器成本改映射：`cost_it_device` 与原十年付款计划归零，`cost_it_other` 保留 `211200000`。
- 增加完整 ICT 标准科目目录参数化测试，验证收入与成本任意旧科目改映射后均生成释放记录并清零，不依赖甲供材料特例。

Decisions:
- 项目周期变化不自动改变业务计划语义；仅平均分年模式重算，首年一次性和自定义计划保持原安排。
- 本阶段继续以智算到 ICT 为正式同步方向，保留现有 ICT 人工覆盖标记，但不增加完整反向同步。

## 2026-06-16

### 智算状态恢复与参数区展开控制

Modified:
- 智算页面打开时从 `projectState.lastResult` 恢复右侧正式 ICT 效益指标，避免必须点击“同步到 ICT”后才显示已保存结论。
- 恢复逻辑校验 `syncRevision > 0` 和完整 `IctResult` 字段；只有当前金额来源出现在 `controlledSubjects.sourceLineItemIds` 时才恢复“已同步”状态。
- 参数区移除关键/已修改/敏感筛选按钮，改为“展开全部 / 折叠全部”；搜索仍保留，搜索结果自动展开匹配类别并暂停拖拽排序。
- 参数卡名称独占标题行宽度，状态标记和说明/设置图标下移，分组网格最小列宽调整为 `240px`，减少长参数名称截断。
- 智算专项测试覆盖正式结果恢复、无效 `lastResult` 拒绝和当前金额来源同步痕迹判断。

Decisions:
- 打开页面只恢复已保存正式 ICT 结果，不生成同步预览，也不绕过用户确认写入 ICT。
- 参数区默认展示全部参数；“关键参数”“已修改”“参与敏感性分析”继续作为参数属性保留在设置菜单中，不再作为顶部隐藏筛选入口。

### 项目看板 project-detail 保存 handler 生命周期修复

Modified:
- ProjectBoard 的 `project-detail` 保存 handler 改为看板级常驻注册，不再随详情抽屉关闭而注销。
- handler 统一保存当前详情抽屉字段和所有发生变化的项目备注草稿，并同步刷新项目列表、当前项目、选中项目、备注草稿和阶段选项。
- 项目卡片备注编辑会先把对应项目设为当前项目，再标记 `project-detail` dirty，避免统一保存上下文指向空项目或旧项目。

Decisions:
- `project-detail` 是项目看板域 dirty scope，不等同于详情抽屉生命周期；卡片备注和详情字段共享同一个保存 handler。
- 本次只修复保存守卫和看板元数据持久化边界，不改变智算测算、ICT 同步或财务计算路径。

### 《会审纪要》模板页面布局组织优化

Modified:
- ICT 生命周期页工作区信息移入顶部 header 右侧按钮/下拉，保留打开目录和修改目录能力，取消该页单独工作区展示行。
- 《会审纪要》模板表单改为内容区顶部 Tab 分组：会审基础信息、项目内容、商务条款、风险与采购、生成确认。
- 技术方案可行性清单、中台能力调用、询价过程和 IT 资金来源相关内容改为低对比 surface 卡片展示。
- 会审模板增加前端完成度提示和固定底部效益评估栏，展示 NPV、净现值率、毛利率、动态回收期、IRR 与生成按钮。
- 会审模板页隐藏原独立底部指标面板；其他模板页面继续使用原模板表单布局和生成按钮。

Decisions:
- Tab 和完成度均为前端展示层，不写入模板状态，不改变字段 key、保存 payload、AI 上下文、询价生成或文档生成变量。
- 本次只在模板名包含“会审”时启用新组织方式，不修改后端接口、Rust 生成逻辑或财务计算公式。

### 一键生成文档模板共享布局扩展

Modified:
- 新增共享模板页面骨架组件，统一承载模板标题、横向 Tab、完成度提示、生成确认面板和固定底部效益评估栏。
- 《ICT项目立项签批表》改为项目基础信息、服务内容、商务条款、立项与采购、生成确认五个内容区 Tab。
- 《ICT项目需求导入表》改为需求基础信息、需求内容、商务与预算、实施与风险、生成确认五个内容区 Tab。
- 三类重点模板共享底部 NPV、净现值率、毛利率、动态回收期、IRR 与立即生成文件入口；ICT 生命周期页对这些模板隐藏旧外层卡片和重复指标面板。
- 需求导入表的技术方案清单、公示/招标材料、信息安全和截图附件保留原交互，改用低对比子模块卡片展示。

Decisions:
- 本次仅调整前端展示组织，不修改 Rust 后端、接口、字段 key、字段映射、模板变量、文件生成逻辑或财务测算逻辑。
- 《立项签批表》和《需求导入表》的完成度沿用前端现有字段值计算；Tab 不写入业务数据，也不改变字段原子编辑、常用内容、AI 分析、图片上传和生成行为。
