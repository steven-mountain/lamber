# agent-bridge — deepseek-harness (dsh) 接入

让 lamber 的 AI 顾问真正具备**工具执行**能力：由 [deepseek-harness](https://github.com/deepseek-ai/deepseek-harness)（`dsh`）承担 agent loop / 工具编排 / 审批，lamber 保留全部 Rust 业务逻辑（`calculator.rs`、`docfill.rs` 等）不变。

## 架构

```
React 前端 (AiChatPanel.tsx)
   │  Tauri invoke: ai_send_prompt / ai_agent_status / ai_agent_stop
   │  Tauri event:  ai://session-event
Rust 后端 (src-tauri/src/agent_bridge/)
   │  spawn 子进程 + newline-delimited JSON-RPC 2.0 over stdio
dsh 子进程 (--profile sdk --patch patch.yml)
   │  dsh-tool-lamber 插件的 execute() 内部
   │  HTTP POST 到 127.0.0.1:<随机端口>（Rust 侧 tiny_http 桥接服务）
   └─ 回调 benefit::calculator::calculate_ict_benefit
```

审批走同一条桥、方向相反：

```
dsh 工具调用前 tools/pre-execute 守卫 → {kind:'ask'}
   │
ctx.approval.request() → approval/request 瀑布 → lamber answerer 插件
   │  HTTP POST /lamber-bridge/approval（挂住不返回）
Rust ApprovalGate：登记槽位 → emit ai://approval-request → 条件变量挂起
   │
React AgentApprovalDialog 弹窗 → 用户点确认/拒绝
   │  invoke ai_resolve_approval(requestId, approved)
Rust 填充槽位 → 唤醒挂起线程 → HTTP 响应
   └─ answerer 返回 ApprovalOutcome → dsh 放行或拒绝该次调用
```

一条原则：**dsh 只做编排，一行业务数学都不在 TypeScript 里。** 插件是桥接服务的瘦客户端，桥接服务是既有 Rust 模块的瘦转发层。

## 目录结构

| 路径 | 作用 |
| --- | --- |
| `dsh-tool-lamber/` | 自定义工具插件（独立 npm 包，`defineTool` 注册 `run_benefit_calculation`） |
| `patch.yml` | dsh profile 补丁层，把插件挂进 `sdk` profile |
| `scripts/provision-profile.mjs` | 一次性准备：构建插件 + `dsh plugin add` 链接进 `$DSH_HOME/profiles/sdk/` |
| `scripts/check-bridge.mjs` | 只跑「插件工具体 → 桥接服务 → calculator」这一跳的排错脚本，不启动 dsh、不消耗 LLM |
| `scripts/check-approval.mjs` | 只跑「answerer → 桥接 → 审批网关 → 结果回传」这一跳，不启动 dsh、不消耗 LLM |
| `scripts/check-gating.mjs` | 打印每个工具是否需要审批，用于断言只读工具没被误拦 |
| `../src-tauri/src/agent_bridge/bridge_server.rs` | 仅监听 `127.0.0.1` 的 HTTP 桥接服务（tiny_http） |
| `../src-tauri/src/agent_bridge/calculation.rs` | `POST /lamber-bridge/calculate` 路由：项目 → 方案 → 最新快照 → 测算引擎 |
| `../src-tauri/src/agent_bridge/dsh_session.rs` | dsh 子进程管理 + JSON-RPC 2.0 客户端 |
| `../src-tauri/src/agent_bridge/approval.rs` | `POST /lamber-bridge/approval` 路由与 `ApprovalGate`（挂起/唤醒/超时/关闭排空） |
| `../src-tauri/src/agent_bridge/approval_log.rs` | 审批审计日志的 SQLite 落库与查询 |
| `../src-tauri/src/agent_bridge/mod.rs` | Tauri 命令、`AgentRuntime` 生命周期、事件转发 |
| `../src-ui/src/components/ai/AgentApprovalDialog.tsx` | 最简审批弹窗，挂在 App 根节点 |
| `../src-ui/src/components/ai/AgentLabView.tsx` | Agent 联调台（`#/agent-lab`），真实应用里驱动 Agent 与审批的唯一入口 |

## 本地准备（一次性）

```bash
cd agent-bridge && npm install && cd dsh-tool-lamber && npm install && cd ..
npm run provision
```

`provision` 做两件事：构建 `dsh-tool-lamber/lib`，然后 `dsh plugin --profile sdk add <插件绝对路径>`。

> **为什么必须 provision**：dsh 解析 `patch.yml` 里的插件包名是**相对 `$DSH_HOME/profiles/<profile>/`**，不是相对当前工作目录。只写 `--patch` 而不先把包链接进 profile 目录，插件加载不到。链接动作由 dsh 转发给 `pnpm` 执行——`pnpm` 已作为 `agent-bridge` 的 devDependency 装在本地，脚本会把 `node_modules/.bin` 放进 `PATH`，无需全局安装。

## 环境变量

| 变量 | 谁设置 | 说明 |
| --- | --- | --- |
| `DSH_HOME` | Rust / provision 脚本 | dsh 状态目录，默认 `agent-bridge/.dsh-home`（已 gitignore） |
| `DSH_TELEMETRY_MODE=DISABLED` | Rust（硬编码） | `sdk` profile 默认把遥测发往 `harness-telemetry.deepseeksvc.com`。lamber 处理客户财务数据，遥测一律关闭 |
| `DEEPSEEK_API_KEY` | 环境 / 后续接前端设置 | 缺省时 dsh 仍能启动并注册工具，只在真正调 LLM 那一步报错 |
| `LAMBER_BRIDGE_URL` | Rust（每次启动） | 桥接服务 origin，端口为绑定时分配的临时端口，**未硬编码** |
| `LAMBER_BRIDGE_TOKEN` | Rust（每次启动） | 每次启动随机生成的令牌，插件必须在请求头带上 |
| `LAMBER_BRIDGE_TOKEN_HEADER` | Rust（硬编码） | 令牌所在请求头名，默认 `x-lamber-bridge-token` |
| `LAMBER_REPO_ROOT` | 可选 | 显式指定仓库根；缺省时从可执行文件路径向上找含 `agent-bridge/patch.yml` 的目录 |
| `LAMBER_APPROVAL_TIMEOUT_SECS` | 可选 | 覆盖 lamber 等待用户确认的秒数，缺省 90。仅影响生产默认值；测试用 `ApprovalGate::new()` 显式传入，不改进程全局状态 |
| `LAMBER_AGENT_LAB` | 可选 | 设为 `1` 时启动后自动跳到 `#/agent-lab` 联调台。窗口没有地址栏，否则该路由无法到达；正常启动完全不受影响 |

**关于令牌**：只绑定回环地址不等于做了鉴权——同机任何进程都能访问该端口，进而读到客户项目的财务数据。因此桥接服务对每个请求校验一次性令牌（定长比较），令牌通过环境变量交给子进程，不落盘。


## 审批通道（闭环 B）

### 设计

**两半，缺一不可。**

* **守卫**（`tools/pre-execute`）决定*哪些*调用需要人。dsh 的 `defineTool` **没有**声明式的"风险等级"字段——被拦的工具就是某个 pre-execute 监听器对它返回 `{kind:'ask'}` 的工具。瀑布的终止默认值是 `{kind:'allow'}`（`dsh-tools/lib/index.js:3117`），所以守卫没点名的工具（包括只读的 `run_benefit_calculation`）原样放行。
* **答复器**（`approval/request`）决定*人说了什么*。它把问题经同一条带鉴权的桥转给 lamber，阻塞等待决定。

### 超时与失败语义

| 环节 | 时长 | 超时结果 |
| --- | --- | --- |
| Rust `ApprovalGate` | 90 秒（可用 `LAMBER_APPROVAL_TIMEOUT_SECS` 覆盖） | 返回 `{approved:false, reason:"等待用户确认超时…"}` |
| 插件 answerer | 180 秒 | `'rejected'` |
| 前端弹窗倒计时 | 跟随后端下发的 `timeoutSeconds` | 自行关闭（后端此时已判拒） |

答复器的 180 秒**故意长于**网关的 90 秒：正常路径应该是网关给出明确的 `rejected`，答复器那道界只兜"桥彻底不应答"。

**全程失败关闭。** 桥接报错、返回体格式不对、超时、锁中毒——一律 `'rejected'`，不挂起也不默认放行。dsh 侧把 `rejected` / `cancelled` / `unavailable` 分别映射成措辞不同的拒绝，模型能区分"用户说不"和"审批通道不可用"。

**没有 answerer 时的默认行为**（读源码确认，两条都是失败关闭）：
- 未挂载 `ApprovalService` → `ctx.get('approval')` 为 `undefined` → 直接拒绝，理由 `tool "X" requires approval (not yet supported)`。
- 挂载了但无人应答 → 瀑布落到 `'unavailable'` → 拒绝，理由 `no approval channel is available`。
- 会话策略 `approval/policy` 为 `'never'` 时，**根本不问人**，一律 `'rejected'`。`sdk` profile 默认是 `'ask'`（实测启动时会 emit `approval/policy {"policy":"ask"}`）。

### 并发

桥接服务改成**每请求一个线程**（上限 16）。审批会把请求挂住到用户应答为止，单线程 accept 循环会让一个弹窗卡死所有其它路由；`a_parked_approval_does_not_block_the_calculation_route` 用例专门守这条。

### 测试工具的边界

`write_test_marker` 只往 `os.tmpdir()` 下的一个新建临时目录写一个带时间戳的文本文件，**不碰 lamber 的工作区、数据库或任何项目文件**。在审批通道验证完成前，这个包里不允许出现任何真正修改业务数据的工具——`AGENTS.md` 明令禁止 AI 绕过用户确认写项目数据。

### 审批的持久化审计（`agent_approval_log`）

每个**已结算**的审批——批准、拒绝、超时、关闭时被中断——都写进工作区 SQLite 的 `agent_approval_log` 表（schema v9 → v10 新增，纯建表无数据迁移）。字段：`request_id` / `tool_name` / `call_id` / `reason` / `args_json` / `approved` / `decided_by` / `decision_reason` / `requested_at` / `decided_at`。

`decided_by` 区分四种来源：`user`（人工点击）、`timeout`（无人应答）、`shutdown`（运行时关闭时中断）、`internal`（锁中毒等异常）。查询用 Tauri 命令 `ai_list_approval_log(limit)`，联调台右栏直接展示。

设计要点：

- **网关本身不碰数据库。** `ApprovalGate` 只持有一个 `ApprovalRecorder` 回调，生产环境由 `approval_log::workspace_recorder` 注入。这样网关可独立测试，存储故障也不会卡住决定。
- **审计写失败绝不改变决定。** 工作区没打开或写库报错，只打一条 stderr 警告然后咽掉。让存储故障阻塞 Agent 是错的；让它**失败放行**更错。`an_audit_write_failure_does_not_alter_the_decision` 守这条。
- **连接是决定时才解析的**，不是启动时捕获的。所以 Agent 起来之后才打开的工作区，审批照样能记上。

**限制：** 审计表在**工作区**数据库里。如果审批发生时没有打开任何工作区，那条记录会丢（只留 stderr 警告）。这是有意的取舍——为审计日志单开一个全局数据库，会引入第二份存储位置和第二套迁移，代价大于收益。

### 挂起槽位的清理

| 场景 | 行为 | 谁兜底 |
| --- | --- | --- |
| 正常关闭（`ai_agent_stop`） | `gate.shutdown()` 立即拒绝所有挂起审批并唤醒线程，随后 `reopen()` 让下次启动可用 | Rust |
| 应用退出（`RunEvent::Exit`） | `main.rs` 的退出钩子调 `shutdown_approvals()`，挂起请求当场收到拒绝 | Rust |
| 关闭后新到的请求 | 网关处于 `closed` 状态，**不建槽位**，当场拒绝 | Rust |
| 进程崩溃 / 被 kill | 桥接套接字随进程消失，answerer 的 `fetch` 报错 → `'rejected'` | 插件侧 catch |
| 桥接活着但不应答 | 网关 90 秒超时 → 明确拒绝 | Rust |
| 桥接彻底失联 | answerer 自己的 180 秒上限 → `'rejected'` | 插件侧 |

**槽位不会永久占用**：每条路径要么填充并移除槽位，要么在超时分支里 `remove` 掉。`shutdown_denies_parked_approvals_immediately` 断言关闭后请求在 10 秒内释放（网关超时设成 300 秒，只有关闭能救它），`approvals_after_shutdown_are_denied_without_parking` 断言关闭后 `pending_count() == 0`。

**dsh 侧不会永久挂起**：崩溃时它的 HTTP 请求直接断开（`answerer_fails_closed_when_the_bridge_dies_mid_request` 用一个接受后立刻关闭套接字的假监听器验证了这一点），最坏情况也有 180 秒兜底。**没有做优雅通知**——lamber 崩溃时没有机会通知 dsh，靠的是连接断开与超时兜底，行为与既有的"失败关闭"原则一致。

### 前端弹窗：已接通，但真实点击未由我验证

**发现的真实缺口**：上一轮虽然写了 `AgentApprovalDialog.tsx` 并挂在 App 根节点，但**整个前端没有任何地方调用 `ai_send_prompt`**——`AiChatPanel` 仍走旧的 `AiRuntime` 直连 OpenAI 的路径。也就是说在真实应用里根本无法启动 Agent，审批弹窗永远不会出现。这一轮补了 `AgentLabView`（`#/agent-lab`）作为触发入口，弹窗在该路由上也一并挂载。

**已自动验证的部分：**

- 应用能真实启动并停在联调台路由，无报错（`LAMBER_AGENT_LAB=1` 启动，进程稳定运行）。
- `the_approval_prompt_contract_matches_the_frontend_dialog`：断言 `ApprovalPrompt` 序列化出的字段恰好是 `args` / `callId` / `reason` / `requestId` / `timeoutSeconds` / `toolName` 六个 camelCase 键，且弹窗源码订阅的事件名、读取的字段名、调用的命令名与传参形状全部对得上。
- `the_approval_dialog_is_mounted_on_every_agent_reachable_route`：断言弹窗同时挂在主界面和联调台两条渲染路径上，且联调台确实调用了 `ai_send_prompt` 与 `ai_list_approval_log`。

**未验证的部分（照实记录）：** 我**没有**用真实鼠标点击跑通确认/拒绝两条路径。本机对 `osascript` 未授予「辅助功能（Accessibility）」权限，`screencapture` 也未授予「屏幕录制」权限，两者都需要在系统设置里由用户本人授权。因此我既无法发出真实点击，也无法截图确认弹窗渲染结果。**这一项需要人工完成**，步骤见下。

上面那两个契约用例覆盖的正是"真实点击本来能发现的那类问题"——跨进程的名字漂移（事件改名、字段被 serde 改了拼写、命令参数对不上）。它们**不能**替代人工验证的是：弹窗的视觉呈现、按钮是否真的可点、遮罩层是否挡住交互。

#### 人工验证步骤

```bash
cd src-tauri && cargo build
```

```bash
DEEPSEEK_API_KEY=<你的key> LAMBER_REPO_ROOT=$PWD LAMBER_AGENT_LAB=1 ./src-tauri/target/debug/benefit-calculator
```

应用会直接停在「Agent 联调台」。然后：

1. 点「发送」（输入框默认就是调用 `write_test_marker` 的指令）。
2. 左栏应依次出现 `turn/start`、`request/header`、`tool/call` 等事件。
3. 弹出「AI 请求执行操作」对话框，显示工具名 `write_test_marker` 和参数 JSON，带倒计时。
4. **点「确认执行」** → 左栏出现 `tool/result`，内含临时目录里的标记文件路径；点「刷新审批日志」，右栏出现一条 `已批准 · user`。
5. 再发一次，这次**点「拒绝」** → `tool/result` 显示被拒，右栏新增一条 `已拒绝 · user`。
6. 什么都不点，等 90 秒 → 右栏出现 `已拒绝 · timeout`。

如果第 3 步弹窗没出现，先跑 `cargo test agent_bridge` 看契约用例是否还过，再看应用 stderr。

## 怎么跑测试

分层设计，失败时能直接指认是哪一半坏了。

```bash
cd src-tauri
cargo test agent_bridge
```
默认跑的 19 个用例**不需要 Node、不需要网络**：

- 计算路由（6）：默认方案 / 按阶段 / 按方案名 / 按方案 id 选择、未知 scenario 拒绝、令牌鉴权、桥接端到端返回引擎数字。
- 审批网关（6）：超时判拒且不挂死、确认唤醒挂起请求、拒绝如实回传、响应不存在的请求报错、审批路由仍需令牌、挂起的审批不阻塞计算路由。
- 审计与关闭（5）：决定跨进程重启仍可查、拒绝与超时可区分、关闭立即释放挂起审批、关闭后请求不建槽位、审计写失败不改变决定。
- 前后端契约（2）：审批事件字段与弹窗读取的名字一致、弹窗挂在所有能触发 Agent 的路由上。

另有 2 个数据库迁移用例在 `cargo test db::`：`v9_database_gains_agent_approval_log_and_preserves_rows`、`fresh_database_uses_schema_v10_*`。

```bash
cargo test agent_bridge -- --ignored --nocapture
```
需要先完成上面的 provision。9 个用例，按"从底层往上缩范围"排列：

| 用例 | 验证什么 | 需要 API key |
| --- | --- | --- |
| `only_the_write_tool_is_gated_behind_approval` | 只读工具没被审批误伤（守卫策略本身） | 否 |
| `plugin_tool_body_reaches_the_calculator_over_the_bridge` | 插件 `execute()` → 桥接 → calculator，数字与引擎直算一致 | 否 |
| `answerer_receives_the_users_decision_through_the_gate` | answerer → 桥接 → 网关 → 模拟前端确认/拒绝 → 结果回到插件（两个方向都测） | 否 |
| `answerer_fails_closed_when_nobody_answers` | 无人应答时 answerer 明确返回 `rejected`，不挂起 | 否 |
| `dsh_advertises_the_lamber_tool_in_its_request_header` | dsh 子进程启动、`initialize` 握手、工具 schema 注入 | 否 |
| `dsh_advertises_the_gated_tool_alongside_the_readonly_one` | 两个工具都在目录里 | 否 |
| `dsh_tool_call_reaches_the_calculator_and_returns_real_numbers` | 闭环 A 完整链路 | **是** |
| `answerer_fails_closed_when_the_bridge_dies_mid_request` | lamber 崩溃（套接字中途断开）时 answerer 失败关闭而非挂起 | 否 |
| `dsh_gated_tool_runs_only_after_the_user_confirms` | 闭环 B 完整链路：确认→真的写出标记文件；拒绝→不执行；`approval/decided` 审计事件与模拟操作一致 | **是** |

排错顺序建议：`check-gating.mjs`（策略对不对）→ `check-bridge.mjs`（桥接坏没坏）→ `check-approval.mjs`（审批通道坏没坏）→ `dsh_advertises...`（插件/协议坏没坏）→ 最后带 key 的两个完整闭环（模型/提示词的问题）。

反复跑时注意：dsh 会把会话持久化到 `$DSH_HOME/sessions/`，**同一个 `sessionId` 复用会报 `id collision`**。测试里每次用新 UUID；手工调试时可以直接删掉该目录。

## 与任务书中「已验证事实」的差异

实际接入时发现以下几点与预设不同，以此处为准：

1. **包版本**：`@deepseek-ai/dsh` 与 `@deepseek-ai/dsh-tools` 最新为 `0.1.2-alpha.5`（不是 `alpha.4`），本项目已固定到 `alpha.5`。`@deepseek-ai/cordis` 仍为 `^4.0.2`。
2. **provider / model 名**：`initialize` 的 `provider` 不是 `deepseek`。`sdk` profile 注册的路由是 **`deepseek-official`**，默认模型 **`deepseek-v4-flash`**；传 `deepseek` 会得到 `no adapter registered for provider "deepseek"`。
3. **`output.schema` 的属性默认是可选的**。除了 `type: 'object'` 必须显式写 `additionalProperties` 之外，对象的每个属性还要加 `required: true`，否则推断出的类型全是 `T | undefined`，`execute()` 的返回值过不了类型检查。
4. **API key 可以走环境变量**。dsh 的报错文案建议「通过 credentials service 存储」，但实测 `DEEPSEEK_API_KEY` 环境变量会被直接采用（用假 key 能拿到上游 401，而不是 `MISSING_CREDENTIAL`），所以 Rust 侧用 env 注入即可。
5. **`dsh plugin add` 会有一条无害告警**：`dsh-tool-lamber declares no dsh.bundle — installed as a plain dependency, not a profile layer`。插件由 `patch.yml` 显式挂载，不依赖 bundle 自动激活，功能正常。
6. **`pnpm` 无需全局安装**，作为 `agent-bridge` 的 devDependency 即可，provision 脚本负责把它放进 `PATH`。

协议本身（`initialize` / `session/prompt` / `shutdown`，以及 `session.event` / `session.status` / `subagent.*` 四种通知，一行一个 JSON 对象）与任务书描述完全一致，已对照 `node_modules/@deepseek-ai/dsh-sdk-protocol/lib/types/types.d.ts` 核对。

### 闭环 B（审批通道）实现时发现的差异

以读到的源码为准（`@deepseek-ai/dsh-user-approval@0.1.2-alpha.5`、`@deepseek-ai/dsh-tools@0.1.2-alpha.5`）：

1. **没有"审批等级"这种声明式字段。** `defineTool` 的选项只有 `name` / `description` / `parameters` / `output` / `timeoutMs` / `isConcurrencySafe` / `execute` / `finalizeContent` / `presentCall` / `presentResult`。要让一个工具需要审批，只能注册 `tools/pre-execute` 监听器对它返回 `{kind:'ask', reason}`。任务书里"标记为需要审批的等级"这个说法在当前版本没有对应实现。
2. **审批请求里没有工具参数。** `ApprovalRequestEvent` 只有 `{agent, toolName, callId?, reason?, signal?}`；其文档明确写着 `callId` "links to an already presented tool call, so arguments are not duplicated here"。要在弹窗里显示参数，必须自己做关联——本实现由守卫按 `callId` 把参数记进进程内的 `pendingCalls` 表，答复器再取回。**这也是 answerer 没有拆成独立 npm 包的原因**：两个包在 pnpm link 下有可能拿到两份模块实例、两张表，守卫写进去的参数答复器读不到。守卫与答复器因此和工具同包，源码上仍分文件（`approval.ts` / `pendingCalls.ts`）。
3. **`ApprovalOutcome` 只有四个值**：`'allowed-once'` / `'rejected'` / `'cancelled'` / `'unavailable'`。只有 `allowed-once` 是放行，且只对这一次调用有效——没有"永久允许"。
4. **`approval/asked` 和 `approval/decided` 是 session 事件**，会经 `session.event` 通知流到 Rust 侧。这比解析 `tool/result` 更适合做审批结果的断言，闭环 B 的完整用例就是这么验的。
5. **`approval.request()` 要求有开启的 turn**，空闲时调用会直接抛错（"an idle ask rejects before appending anything"），因为审计事件对必须落在日志的提交边界内。
6. **桥接服务原来的单线程 accept 循环撑不住审批**：一个挂起的审批会卡死其它所有路由。已改成每请求一线程（上限 16）。这是闭环 A 遗留的设计缺陷，被闭环 B 暴露出来。

## 目前还没做的部分（留给下一轮）

* **真实写操作工具。** 审批通道已验证可用，但本轮**只**放了无害的 `write_test_marker`。要接真正会改 lamber 数据的工具（改方案、生成文档等），下一轮把它们加进 `approval.ts` 的 `GATED_TOOLS`，并逐个补桥接路由与端到端用例。
* **对话流展示。** `AiChatPanel.tsx` / `AiRuntime.ts` 仍未改动，工具调用过程没有在对话里可视化。Rust 侧已把所有通知按 `{method, params}` 打包 emit 到 `ai://session-event`，前端订阅这一个事件、按 `method` 分派即可，新增通知类型不需要再加 Tauri 事件名。审批弹窗是独立事件 `ai://approval-request`，已接好。
* **审批弹窗的产品化。** 现在是能演示流程的最简实现：单一模态、JSON 原样展示参数、倒计时。没有做"记住选择""按工具批量授权""历史审批记录"，也没有按 DESIGN.md 精修视觉。
* **API key 来源。** 目前从环境变量读。前端的 key 存在 `localStorage.lamber_ai_api_key`，接入时应改为由前端传入或落到 lamber 自己的凭证存储，而不是依赖环境变量。
* **单文件 SEA 打包 / 瘦身。** 现在依赖开发机上的 `agent-bridge/node_modules`（约 520 个包），发布形态未处理。

## 依赖说明

Rust 侧新增一个直接依赖：`tiny_http 0.12`（`default-features = false`，仅 `ascii`/`chunked_transfer`/`httpdate`/`log` 四个小传递依赖，不含 TLS）。

选它而不是 hyper/axum 的原因：lamber 后端整体是同步的（`rusqlite` + `Arc<Mutex<Connection>>`，`calculate_ict_benefit` 是同步函数），在异步 handler 里做同步加锁是反模式；桥接服务只有一个回环路由、并发量极低，用阻塞式服务器跑在独立线程上与既有代码风格一致。也没有手写 HTTP 解析——那是比引入一个小依赖更大的技术债。
