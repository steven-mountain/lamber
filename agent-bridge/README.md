# agent-bridge — deepseek-harness (dsh) 接入

让 lamber 的 AI 顾问真正具备**工具执行**能力：由 [deepseek-harness](https://github.com/deepseek-ai/deepseek-harness)（`dsh`）承担 agent loop / 工具编排 / 审批，lamber 保留全部 Rust 业务逻辑（`calculator.rs`、`docfill.rs` 等）不变。

## 架构

```
React 前端（AiChatPanel 唯一使用 dsh；AgentLabView 保留排障）
   │  Tauri invoke: ai_send_prompt / ai_agent_status / ai_agent_stop
   │  Tauri event:  ai://session-event
Rust 后端 (src-tauri/src/agent_bridge/)
   │  spawn 子进程 + Agent Client Protocol (ACP) over stdio
dsh 子进程 (--profile acp --patch patch.yml)
   │  dsh-tool-lamber 插件的 execute() 内部
   │  HTTP POST 到 127.0.0.1:<随机端口>（Rust 侧 tiny_http 桥接服务）
   └─ 回调 benefit::calculator::calculate_ict_benefit
```

审批**不走这条桥**，走 ACP 那条连接、方向相反——这是 ACP 化之后最容易记错的一点：

```
dsh 工具调用前 tools/pre-execute 守卫 → {kind:'ask'}
   │
ctx.approval.request() → approval/request 瀑布
   │  ← dsh-acp 自己的处理器抢先接住（排在 dsh-tool-lamber 之前）
dsh-acp 转成 ACP 请求 session/requestPermission 发给客户端
   │  params 里只有 sessionId / toolCall.toolCallId / options 两项
Rust dsh_session.rs 的 on_receive_request 处理器
   │  按 toolCallId 去 ToolCallIndex 取回工具名与参数
   │  （这两样来自更早那条 session/update 的 tool_call 通知）
Rust ApprovalGate：登记槽位 → emit ai://approval-request → 条件变量挂起
   │
React AgentApprovalDialog 弹窗 → 用户点确认/拒绝
   │  invoke ai_resolve_approval(requestId, approved)
Rust 填充槽位 → 唤醒挂起的 blocking 任务
   └─ 回 session/requestPermission：Selected(allow-once) 或 Selected(reject-once)
```

> **`/lamber-bridge/approval` 这条 HTTP 路由已经删除**，插件里的 answerer 与
> `pendingCalls.ts` 也一并删除。ACP 下 `dsh-acp` 会先接住 `approval/request`
> 事件，插件自己的答复器永远排不上，留着就是死代码加第二条通往同一个「批准」的路。

一条原则：**dsh 只做编排，一行业务数学都不在 TypeScript 里。** 插件是桥接服务的瘦客户端，桥接服务是既有 Rust 模块的瘦转发层。

## 目录结构

| 路径 | 作用 |
| --- | --- |
| `dsh-tool-lamber/` | 自定义工具插件（独立 npm 包，`defineTool` 注册 `run_benefit_calculation`） |
| `patch.yml` | dsh profile 补丁层，把插件挂进 `acp` profile |
| `dsh-home-template/` | 不含会话、凭据、瞬态 lock 或绝对链接的清洁用户运行目录模板 |
| `scripts/provision-profile.mjs` | 一次性准备：构建插件 + `dsh plugin add` 链接进 `$DSH_HOME/profiles/acp/` |
| `scripts/check-bridge.mjs` | 只跑「插件工具体 → 桥接服务 → calculator」这一跳的排错脚本，不启动 dsh、不消耗 LLM |
| `scripts/check-gating.mjs` | 打印每个工具是否需要审批，用于断言只读工具没被误拦 |
| `../src-tauri/src/agent_bridge/bridge_server.rs` | 仅监听 `127.0.0.1` 的 HTTP 桥接服务（tiny_http） |
| `../src-tauri/src/agent_bridge/calculation.rs` | `POST /lamber-bridge/calculate` 路由：项目 → 方案 → 最新快照 → 测算引擎 |
| `../src-tauri/src/agent_bridge/dsh_session.rs` | dsh 子进程管理 + ACP 客户端（握手与版本断言、`session/new`、`session/prompt`、`session/requestPermission` 应答） |
| `../src-tauri/src/agent_bridge/distribution.rs` | 安装资源 / 开发资源定位、用户 DSH_HOME 初始化与模型/baseURL patch 生成 |
| `../src-tauri/src/agent_bridge/tool_calls.rs` | 按 `toolCallId` 关联 `tool_call` 通知与随后的权限请求（原 `pendingCalls.ts` 的继任者） |
| `../src-tauri/src/agent_bridge/approval.rs` | `ApprovalGate`（挂起/唤醒/超时/关闭排空）与展示文案镜像表 |
| `../src-tauri/src/agent_bridge/approval_log.rs` | 审批审计日志的 SQLite 落库与查询 |
| `../src-tauri/src/agent_bridge/mod.rs` | Tauri 命令、`AgentRuntime` 生命周期、事件转发 |
| `../src-ui/src/components/ai/AgentApprovalDialog.tsx` | 最简审批弹窗，挂在 App 根节点 |
| `../src-ui/src/components/ai/AgentLabView.tsx` | Agent 联调台（`#/agent-lab`）；阶段 1 可从设置卡进入以执行装机 gate，不是最终产品 Chat 路径 |
| `../src-ui/src/components/settings/AiAgentSettingsCard.tsx` | 产品内 key、dsh 模型和 baseURL 设置；密钥不回显，保存后停止旧 dsh |
| `../docs/verification/acp-approval-manual-check.md` | ACP 下四条审批路径的真人点击验证记录 |

## 开发环境准备（一次性）

```bash
cd agent-bridge && npm install && cd dsh-tool-lamber && npm install && cd ..
npm run provision
```

`provision` 做两件事：构建 `dsh-tool-lamber/lib`，然后 `dsh plugin --profile acp add <插件绝对路径>`。

这只适用于开发与旧的专项测试。装机版不运行 provision，也不读取仓库 `.dsh-home`：Windows
打包脚本用锁文件准备 production runtime，应用首次启动时把清洁模板复制到用户数据目录，
再将已构建插件同步进去。

> **为什么必须 provision**：dsh 解析 `patch.yml` 里的插件包名是**相对 `$DSH_HOME/profiles/<profile>/`**，不是相对当前工作目录。只写 `--patch` 而不先把包链接进 profile 目录，插件加载不到。链接动作由 dsh 转发给 `pnpm` 执行——`pnpm` 已作为 `agent-bridge` 的 devDependency 装在本地，脚本会把 `node_modules/.bin` 放进 `PATH`，无需全局安装。

## 环境变量

| 变量 | 谁设置 | 说明 |
| --- | --- | --- |
| `DSH_HOME` | Rust / provision 脚本 | 装机版使用 `<app_data_dir>/dsh-runtime`；开发 provision 默认 `agent-bridge/.dsh-home` |
| `DSH_TELEMETRY_MODE=DISABLED` | Rust（硬编码） | `sdk` profile 默认把遥测发往 `harness-telemetry.deepseeksvc.com`。lamber 处理客户财务数据，遥测一律关闭 |
| `DEEPSEEK_API_KEY` | Rust | 装机版从 Lamber 应用配置注入；仅 debug 构建在配置为空时允许环境变量兜底 |
| `LAMBER_BRIDGE_URL` | Rust（每次启动） | 桥接服务 origin，端口为绑定时分配的临时端口，**未硬编码** |
| `LAMBER_BRIDGE_TOKEN` | Rust（每次启动） | 每次启动随机生成的令牌，插件必须在请求头带上 |
| `LAMBER_BRIDGE_TOKEN_HEADER` | Rust（硬编码） | 令牌所在请求头名，默认 `x-lamber-bridge-token` |
| `LAMBER_REPO_ROOT` | 可选、仅开发 | 安装资源定位失败后显式指定仓库根；发布运行不依赖它 |
| `LAMBER_APPROVAL_TIMEOUT_SECS` | 可选 | 覆盖 lamber 等待用户确认的秒数，缺省 90。仅影响生产默认值；测试用 `ApprovalGate::new()` 显式传入，不改进程全局状态 |
| `LAMBER_AGENT_LAB` | 可选 | 设为 `1` 时启动后自动跳到 `#/agent-lab` 联调台。窗口没有地址栏，否则该路由无法到达；正常启动完全不受影响 |

**关于令牌**：只绑定回环地址不等于做了鉴权——同机任何进程都能访问该端口，进而读到客户项目的财务数据。因此桥接服务对每个请求校验一次性令牌（定长比较），令牌通过环境变量交给子进程，不落盘。


## 审批通道（闭环 B，ACP 化后）

### 设计

**仍是两半，但两半现在分居协议两侧。**

* **守卫**（`tools/pre-execute`，在插件里）决定*哪些*调用需要人。dsh 的 `defineTool` **没有**声明式的"风险等级"字段——被拦的工具就是某个 pre-execute 监听器对它返回 `{kind:'ask'}` 的工具。瀑布的终止默认值是 `{kind:'allow'}`（`dsh-tools/lib/index.js:3117`），所以守卫没点名的工具（包括只读的 `run_benefit_calculation`）原样放行。这一半 ACP 化后**完全没动**。
* **应答方**（`session/requestPermission`，在 lamber 里）决定*人说了什么*。ACP 化前它是插件里的 `approval/request` 监听器 + 一条 HTTP 路由；现在 `dsh-acp` 抢先接住那个 Cordis 事件并转成 ACP 请求发给客户端，所以答复器移进了 Rust。

> 顺带一提：ACP 下 lamber 是这条连接上**唯一**的权限应答方。`acp` profile 里 dsh 自带的
> 工具若触发权限请求，也会走到同一个弹窗。策略是"照问不误、从不默许"——索引里查不到对应
> 的 `tool_call` 公告时，弹窗会以调用 id 命名并给通用文案，但仍然会问。

### 超时与失败语义

| 环节 | 时长 | 超时结果 |
| --- | --- | --- |
| Rust `ApprovalGate` | 600 秒（10分钟）（可用 `LAMBER_APPROVAL_TIMEOUT_SECS` 覆盖） | 应答 `Selected(reject-once)`，审计理由为"等待用户确认超时…" |
| 前端弹窗倒计时 | 跟随后端下发的 `timeoutSeconds` | 自行关闭（后端此时已判拒） |

**ACP 下 dsh 不给权限请求设上限**，它就等客户端的应答。SDK 时期插件答复器那道 180 秒的
兜底因此没有了对应物：网关的 600 秒（10分钟）是**唯一**能结束一次无人应答的机制，必须永远会触发。

**全程失败关闭。** 桥接报错、返回体格式不对、超时、锁中毒——一律 `'rejected'`，不挂起也不默认放行。dsh 侧把 `rejected` / `cancelled` / `unavailable` 分别映射成措辞不同的拒绝，模型能区分"用户说不"和"审批通道不可用"。

**没有应答方时的默认行为**（读源码确认，都是失败关闭）：
- 未挂载 `ApprovalService` → `ctx.get('approval')` 为 `undefined` → 直接拒绝，理由 `tool "X" requires approval (not yet supported)`。
- 挂载了但瀑布无人接 → 落到 `'unavailable'` → 拒绝，理由 `no approval channel is available`。
- 客户端答 `Cancelled`（本实现只在 agent 没给出所需 kind 的选项时用它）→ dsh 记为 `'cancelled'`，同样不放行。

### 并发

**两处并发，别混为一谈。**

* **ACP 连接**：权限处理器跑在连接的派发循环上，直接在里面挂起会卡住整条连接——包括
  弹窗打开期间本该继续流回来的 `session/update`。所以审批是 `cx.spawn` +
  `spawn_blocking` 出去等的，连接线程用 2 worker 的多线程 tokio runtime。
* **桥接服务**：仍是每请求一个线程（上限 16）。审批已经不走这条路了，但工具调用可能并发，
  这个结构保留不变。

### 测试工具的边界

`write_test_marker` 只往 `os.tmpdir()` 下的一个新建临时目录写一个带时间戳的文本文件，**不碰 lamber 的工作区、数据库或任何项目文件**。在审批通道验证完成前，这个包里不允许出现任何真正修改业务数据的工具——`AGENTS.md` 明令禁止 AI 绕过用户确认写项目数据。

### 审批的持久化审计（`agent_approval_log`）

每个**已结算**的审批——批准、拒绝、超时、关闭时被中断——都写进工作区 SQLite 的 `agent_approval_log` 表（schema v9 → v10 新增，纯建表无数据迁移）。字段：`request_id` / `tool_name` / `call_id` / `reason` / `args_json` / `approved` / `decided_by` / `decision_reason` / `requested_at` / `decided_at`。

`decided_by` 区分四种来源：`user`（人工点击）、`timeout`（无人应答）、`shutdown`（运行时关闭时中断）、`internal`（锁中毒等异常）。查询用 Tauri 命令 `ai_list_approval_log(limit)`，联调台右栏直接展示。

设计要点：

- **网关本身不碰数据库。** `ApprovalGate` 只持有一个 `ApprovalRecorder` 回调，生产环境由 `approval_log::workspace_recorder` 注入。这样网关可独立测试，存储故障也不会卡住决定。
- **审计写失败绝不改变决定。** 工作区没打开或写库报错，只打一条 stderr 警告然后咽掉。让存储故障阻塞 Agent 是错的；让它**失败放行**更错。`an_audit_write_failure_does_not_alter_the_decision` 守这条。
- **连接是决定时才解析的**，不是启动时捕获的。所以 Agent 起来之后才打开的工作区，审批照样能记上。

#### 没打开工作区时怎么办：缓冲 + 回填（已拍板，不再是开放取舍）

审计表在**工作区**数据库里，而审批可能发生在没有打开任何工作区的时候。**不允许静默丢失。**

考虑过两个方向：

1. **阻塞审批直到有工作区可写** —— 否决。弹窗里没有"打开工作区"这个动作，用户站在弹窗前无法解除这个条件，Agent 的这一轮就会挂在一个他解不开的前提上，直接违反超时路径已经确立的"绝不挂起"原则。而且审批本来就不一定和项目有关（`write_test_marker` 就完全无关），拿工作区去卡一个与它无关的决定是错的耦合。
2. **缓冲到本地文件，工作区打开后回填** —— 采用。

实现：写不进工作区库的决定，追加到应用数据目录下的 `agent-approval-spool.jsonl`（一行一条 JSON）。`workspace::open_workspace_internal`（**唯一**一处数据库变为可用的地方，启动时恢复上次工作区也走它）在 `switch_workspace` 之后调用 `drain_spool_on_workspace_open`，把缓冲整体搬进 `agent_approval_log`。

几个刻意的性质：

- **回填在一个事务里完成，成功之后才删缓冲文件。** 中途失败就保留缓冲，下次工作区打开时重试——宁可重放，不可丢失。
- **重放是幂等的。** 行以 `request_id` 为主键、用 `INSERT OR REPLACE` 写入，同一条决定回填多次也只有一行。`backfilling_the_same_decisions_twice_does_not_duplicate_rows` 守这条。
- **崩溃写坏的半行不会挡住其它记录。** 追加是 `O_APPEND` 逐行写，最多损坏尾行；回填时跳过解析失败的行并告警，其余照常入库（`a_corrupt_spool_line_does_not_block_backfilling_the_others`）。
- **缓冲写失败仍不改变决定。** 磁盘满或只读时退回到 stderr 告警。这是**唯一剩下的丢失路径**，而且它是响的。

限制：缓冲文件按应用数据目录存放，是**跨工作区**的。若在无工作区时产生审批、随后打开的是另一个工作区，记录会回填进那个工作区。对单人桌面工具这是合理的（决定是这台机器上这个人做的），但它确实意味着审计行的归属是"回填时所在的工作区"，不是"决定发生时的工作区"——决定发生时本来就没有工作区。

### 挂起槽位的清理

| 场景 | 行为 | 谁兜底 |
| --- | --- | --- |
| 正常关闭（`ai_agent_stop`） | `gate.shutdown()` 立即拒绝所有挂起审批并唤醒线程，随后 `reopen()` 让下次启动可用 | Rust |
| 应用退出（`RunEvent::Exit`） | `main.rs` 的退出钩子调 `shutdown_approvals()`，挂起请求当场收到拒绝 | Rust |
| 关闭后新到的请求 | 网关处于 `closed` 状态，**不建槽位**，当场拒绝 | Rust |
| 进程崩溃 / 被 kill | ACP stdio 连接随进程关闭，dsh 的权限请求以传输关闭失败 | 传输层 |
| 无人点弹窗 | 网关 600 秒（10分钟）超时 → 明确拒绝并应答 `reject-once` | Rust（**唯一**兜底） |

**槽位不会永久占用**：每条路径要么填充并移除槽位，要么在超时分支里 `remove` 掉。`shutdown_denies_parked_approvals_immediately` 断言关闭后请求在 10 秒内释放（网关超时设成 300 秒，只有关闭能救它），`approvals_after_shutdown_are_denied_without_parking` 断言关闭后 `pending_count() == 0`。

**dsh 侧不会永久挂起**：lamber 崩溃时 ACP 的 stdio 连接随进程关闭，dsh 的
`session/requestPermission` 以传输关闭失败；子进程本身也会随进程组一起被收掉
（crate 对 spawn 出来的进程组装了 `ChildGuard`）。**没有做优雅通知**——lamber 崩溃时没有
机会通知 dsh，靠的是连接断开，行为与既有的"失败关闭"原则一致。

### 前端弹窗：已接通，但真实点击未由我验证

**发现的真实缺口**：上一轮虽然写了 `AgentApprovalDialog.tsx` 并挂在 App 根节点，但**整个前端没有任何地方调用 `ai_send_prompt`**——`AiChatPanel` 仍走旧的 `AiRuntime` 直连 OpenAI 的路径。也就是说在真实应用里根本无法启动 Agent，审批弹窗永远不会出现。这一轮补了 `AgentLabView`（`#/agent-lab`）作为触发入口，弹窗在该路由上也一并挂载。

**已自动验证的部分：**

- 应用能真实启动并停在联调台路由，无报错（`LAMBER_AGENT_LAB=1` 启动，进程稳定运行）。
- `the_approval_prompt_contract_matches_the_frontend_dialog`：断言 `ApprovalPrompt` 序列化出的字段恰好是 `args` / `callId` / `reason` / `requestId` / `timeoutSeconds` / `toolName` 六个 camelCase 键，且弹窗源码订阅的事件名、读取的字段名、调用的命令名与传参形状全部对得上。
- `the_approval_dialog_is_mounted_on_every_agent_reachable_route`：断言弹窗同时挂在主界面和联调台两条渲染路径上，且联调台确实调用了 `ai_send_prompt` 与 `ai_list_approval_log`。

**已完成真实点击验证**：见 [`docs/verification/approval-channel-manual-check.md`](../docs/verification/approval-channel-manual-check.md)。四条路径（弹窗渲染 / 点确认 / 点拒绝 / 不操作等超时）全部通过，未发现弹窗未渲染、按钮点不动、被遮罩挡住或状态不同步的问题。

取证方式说明：只有"弹窗渲染"用截图（启动瞬间应用必然在最前，裁剪到窗口区域），其余三项一律以**持久化审计表 `agent_approval_log` 与文件系统产物**为证据。原因有二——全屏截图会连带拍到桌面上其它窗口的私有内容，不适合入库；而审计行 + 标记文件的时间戳比截图更难伪造，也更贴近"事后可追溯"这个验收目标本身。

**已知限制：确认/拒绝两次点击是人工完成的，不是自动化验证。** 本机对 `osascript` 的辅助功能（Accessibility）授权始终未生效——即使在系统设置里授权后，`osascript` 仍持续报 `-1719` / `-25211`，推测是 TCC 决定在授权前就已被宿主进程缓存、需重启宿主应用才会生效。因此这两条路径由仓库所有者本人点击，AI 侧只做截图与审计数据核对。

这意味着：

- **超时路径是完全自动的**（无需点击），可以在 CI 或无人值守环境里重复验证。
- **确认/拒绝两条路径目前无法自动化重跑。** 若要把它们纳入自动化，需要另找不依赖 Accessibility 的驱动方式（例如给前端加一个仅测试构建可用的注入入口），但那样验证的就不再是"真实点击"了——这个取舍留给后续判断。
- 与之相对，`ai_resolve_approval` 的**命令契约**（事件字段名、命令参数名）是有自动化用例守着的，跨进程改名会被测试抓住；测不到的只有视觉呈现与鼠标可达性。

#### 人工复现步骤

```bash
cd src-tauri && cargo build
```

```bash
DEEPSEEK_API_KEY=<你的key> LAMBER_REPO_ROOT=$PWD LAMBER_AGENT_LAB=1 ./src-tauri/target/debug/benefit-calculator
```

应用会直接停在「Agent 联调台」。`LAMBER_AGENT_LAB=autorun` 则会额外自动发一次指令，弹窗无需点「发送」即可出现。然后：

1. 点「发送」（输入框默认就是调用 `write_test_marker` 的指令；`autorun` 模式可跳过）。
2. 左栏应依次出现 `turn/start`、`request/header`、`tool/call` 等事件。
3. 弹出「AI 请求执行操作」对话框，显示工具名 `write_test_marker` 和参数 JSON，带倒计时。
4. **点「确认执行」** → 左栏出现 `tool/result`，内含临时目录里的标记文件路径；点「刷新审批日志」，右栏出现一条 `已批准 · user`。
5. 再发一次，这次**点「拒绝」** → `tool/result` 显示被拒，右栏新增一条 `已拒绝 · user`。
6. 什么都不点，等 600 秒（10分钟） → 右栏出现 `已拒绝 · timeout`。

如果第 3 步弹窗没出现，先跑 `cargo test agent_bridge` 看契约用例是否还过，再看应用 stderr。

## 怎么跑测试

分层设计，失败时能直接指认是哪一半坏了。

```bash
cd src-tauri
cargo test agent_bridge
```
默认套件中的 Agent Bridge 用例**不需要 Node、不需要网络**；此外还包含运行资源定位、
清洁模板复制和旧配置兼容测试：

- 计算路由（6）：默认方案 / 按阶段 / 按方案名 / 按方案 id 选择、未知 scenario 拒绝、令牌鉴权、桥接端到端返回引擎数字。
- 审批网关（6）：超时判拒且不挂死、确认唤醒挂起请求、拒绝如实回传、响应不存在的请求报错、抢在公告返回前到达的答复不丢失、旧的 `/lamber-bridge/approval` 路由确已下线（404）。
- ACP 关联与策略（4）：工具调用索引取回公告过的名字与参数、重复公告不留两行、索引有上限会淘汰、Rust 展示文案镜像表与插件 `GATED_TOOLS` 名单一致。
- 审计与关闭（5）：决定跨进程重启仍可查、拒绝与超时可区分、关闭立即释放挂起审批、关闭后请求不建槽位、审计写失败不改变决定。
- 前后端契约（2）：审批事件字段与弹窗读取的名字一致、弹窗挂在所有能触发 Agent 的路由上。

另有 2 个数据库迁移用例在 `cargo test db::`：`v9_database_gains_agent_approval_log_and_preserves_rows`、`fresh_database_uses_schema_v10_*`。

```bash
cargo test agent_bridge -- --ignored --nocapture
```
开发树用例需要先完成上面的 provision；打包资源用例还需要先执行
`prepareAgentRuntime()`。当前共有 10 个 ignored 用例，除下表原有闭环外还包括阶段 0/1
的模型图片能力、严格 OpenAI 端点拒绝、清洁 home 和打包 staging 握手验证：

| 用例 | 验证什么 | 需要 API key |
| --- | --- | --- |
| `only_the_write_tool_is_gated_behind_approval` | 只读工具没被审批误伤（守卫策略本身） | 否 |
| `plugin_tool_body_reaches_the_calculator_over_the_bridge` | 插件 `execute()` → 桥接 → calculator，数字与引擎直算一致 | 否 |
| `acp_handshake_negotiates_the_expected_protocol_version` | `acp` profile 装配、`--patch` 加载、ACP 握手、**协议版本断言**、`session/new` | 否 |
| `a_finished_turn_reports_its_stop_reason` | 投递即返回的一轮会发出 `session/turn-ended` 终结事件 | **是** |
| `dsh_tool_call_reaches_the_calculator_and_returns_real_numbers` | 闭环 A 完整链路，且只读工具全程不弹审批 | **是** |
| `dsh_gated_tool_runs_only_after_the_user_confirms` | 闭环 B 完整链路：`tool_call` 先于权限请求到达、弹窗认出工具名与参数、确认→真的写出标记文件、拒绝→不写 | **是** |

没有 key 时，带 key 的三个用例会打印「跳过：未设置 DEEPSEEK_API_KEY」并通过——**这是跳过，不是验证过**。
带真实 key 的完整运行结果记在
[`docs/CURRENT_TASK.md`](../docs/CURRENT_TASK.md) 的 Validation 一节。

排错顺序建议：`check-gating.mjs`（拦截策略对不对）→ `check-bridge.mjs`（桥接坏没坏）→ `acp_handshake_negotiates_the_expected_protocol_version`（profile / 协议 / 握手坏没坏，不需要 key）→ 最后带 key 的几个完整闭环（模型/提示词的问题）。

反复跑时注意：dsh 仍会把会话持久化到 `$DSH_HOME/sessions/`。ACP 下会话 id 由
**dsh 自己生成**（`session/new` 的返回值），不再由 lamber 指定，所以 SDK 时代那个
"复用同一个 sessionId 会报 id collision" 的坑不复存在。手工调试时该目录仍可直接删掉。

## 与任务书中「已验证事实」的差异

实际接入时发现以下几点与预设不同，以此处为准：

1. **包版本**：`@deepseek-ai/dsh` 与 `@deepseek-ai/dsh-tools` 最新为 `0.1.2-alpha.5`（不是 `alpha.4`），本项目已固定到 `alpha.5`。`@deepseek-ai/cordis` 仍为 `^4.0.2`。
2. **provider / model 名**：路由不是 `deepseek`，而是 **`deepseek-official`**，默认模型 **`deepseek-v4-flash`**；传 `deepseek` 会得到 `no adapter registered for provider "deepseek"`。
3. **`output.schema` 的属性默认是可选的**。除了 `type: 'object'` 必须显式写 `additionalProperties` 之外，对象的每个属性还要加 `required: true`，否则推断出的类型全是 `T | undefined`，`execute()` 的返回值过不了类型检查。
4. **API key 可以由环境变量注入 dsh**。装机版由 Lamber 应用配置读取密钥并通过子进程环境注入；只有 debug 构建在配置为空时保留 `DEEPSEEK_API_KEY` 开发兜底。用假 key 实测会拿到上游 401，而不是 `MISSING_CREDENTIAL`。
5. **`dsh plugin add` 会有一条无害告警**：`dsh-tool-lamber declares no dsh.bundle — installed as a plain dependency, not a profile layer`。插件由 `patch.yml` 显式挂载，不依赖 bundle 自动激活，功能正常。
6. **`pnpm` 无需全局安装**，作为 `agent-bridge` 的 devDependency 即可，provision 脚本负责把它放进 `PATH`。

### ACP 协议层重写时发现的差异

以读到的源码为准（`agent-client-protocol@2.0.0`、`@deepseek-ai/dsh-acp@0.1.2-alpha.5`）。
握手兼容性的一次性验证记录见
[`docs/verification/acp-rust-crate-handshake.md`](../docs/verification/acp-rust-crate-handshake.md)。

1. **依赖版本以 `cargo tree` 实测为准。** `agent-client-protocol 2.0.0` 把
   `agent-client-protocol-schema` **精确锁死在 `=1.5.0`**（不是早期调研以为的
   1.7.0）。两者都讲 ACP 线上协议 **v1**——线上版本是个整数，跟包版本号不是一套体系。
2. **dsh 不做版本协商。** `initialize(_params)` 忽略客户端请求的版本，无条件回自己的
   `PROTOCOL_VERSION`（`dsh-acp/lib/index.js:1143-1146`）。也就是说 dsh 升到 v2 那天，
   **握手期不会有任何报错来提示我们**，不匹配只会在后面某个字段上炸开。因此客户端侧在
   `dsh_session.rs` 里加了显式断言：协商结果不等于 `EXPECTED_PROTOCOL_VERSION` 就直接
   启动失败。`acp_handshake_negotiates_the_expected_protocol_version` 守着这条。
3. **`session/requestPermission` 里没有工具名，也没有参数。** dsh 构造的 params 只有
   `{sessionId, toolCall: {toolCallId}, options: [allow-once, reject-once]}`
   （`dsh-acp/lib/index.js:1118-1140`），连守卫写的 `reason` 都被丢掉。工具名与参数来自
   更早那条 `session/update` 的 `tool_call` 通知（`title` 与 `rawInput`）——dsh 在发问前
   会先 `drainUpdates()`，所以那条通知一定已经到了。Rust 侧 `tool_calls.rs` 按
   `toolCallId` 做这个关联，它是插件里 `pendingCalls.ts` 的继任者。
   弹窗要显示的说明文案则由 `approval.rs` 的镜像表提供，`gated_tool_names_match_the_plugin`
   保证它和插件的 `GATED_TOOLS` 不会各走各的。
4. **`AcpAgentConfig` 没有 `env_remove`，也没有 `current_dir`。** 子进程继承 lamber 的环境
   变量，所以"配置里没有 key"必须显式传空串来表达——dsh 的凭证层把空值等同于未设置
   （`dsh-credentials-local/lib/index.js:427`），这与旧代码的 `env_remove` 语义一致，不是
   将就。工作目录则由 `session/new` 的 `cwd` 参数显式给出，不再依赖子进程的 cwd。
5. **ACP 会话 id 由 agent 决定。** `session/new` 返回 dsh 生成的 id，与 SDK 协议下由
   lamber 指定相反。前端仍然给自己的会话起名，`AgentRuntime` 负责把两套 id 映射起来。
6. **`session/prompt` 是长请求**，要整轮结束才回 `stopReason`。lamber 侧做成投递即返回，
   轮次结束另发一条 `session/turn-ended` 事件，否则一次 `ai_send_prompt` 会挂住整轮。
7. **处理器跑在连接的派发循环上。** 权限处理器里直接挂起会卡住整条连接的消息收发，所以
   审批是 `cx.spawn` + `spawn_blocking` 出去等的；连接线程用的是 2 worker 的多线程
   tokio runtime。tokio 只出现在这一层，后端其余部分仍是同步的。
8. **`session/resume` 不等于 `session/load`。** dsh 只实现前者，且不重放历史消息；跨重启
   要显示的历史对话仍得靠 lamber 自己的会话记录。
9. **ACP 支持图片内容块**（`ImageContent`，base64 + mimeType）。协议重写时未启用；
   阶段 2 已在视觉模型下实际接通，普通模型仍明确拒绝。

### 闭环 B（审批通道）实现时发现的差异（SDK 协议时期，历史记录）

以下是 `--profile sdk` 时期的记录，保留作为背景。其中第 2、4 条描述的 HTTP 审批通道已在
ACP 重写中删除，**不要照着它开工**；第 1、3、5、6 条仍然成立。
以读到的源码为准（`@deepseek-ai/dsh-user-approval@0.1.2-alpha.5`、`@deepseek-ai/dsh-tools@0.1.2-alpha.5`）：

1. **没有"审批等级"这种声明式字段。** `defineTool` 的选项只有 `name` / `description` / `parameters` / `output` / `timeoutMs` / `isConcurrencySafe` / `execute` / `finalizeContent` / `presentCall` / `presentResult`。要让一个工具需要审批，只能注册 `tools/pre-execute` 监听器对它返回 `{kind:'ask', reason}`。任务书里"标记为需要审批的等级"这个说法在当前版本没有对应实现。
2. **审批请求里没有工具参数。** `ApprovalRequestEvent` 只有 `{agent, toolName, callId?, reason?, signal?}`；其文档明确写着 `callId` "links to an already presented tool call, so arguments are not duplicated here"。当时由守卫把参数记进进程内的 `pendingCalls` 表、答复器再取回。**ACP 下这个约束依然在，但关联点搬到了 Rust 侧**（`tool_calls.rs`），插件里的答复器与 `pendingCalls.ts` 已删除。
3. **`ApprovalOutcome` 只有四个值**：`'allowed-once'` / `'rejected'` / `'cancelled'` / `'unavailable'`。只有 `allowed-once` 是放行，且只对这一次调用有效——没有"永久允许"。
4. ~~**`approval/asked` 和 `approval/decided` 是 session 事件**，会经 `session.event` 通知流到 Rust 侧。~~ ACP 下没有这两个事件：审批结果只体现在客户端自己给出的应答，以及随后的 `tool_call_update` 上。
5. **`approval.request()` 要求有开启的 turn**，空闲时调用会直接抛错（"an idle ask rejects before appending anything"），因为审计事件对必须落在日志的提交边界内。
6. **桥接服务原来的单线程 accept 循环撑不住审批**：一个挂起的审批会卡死其它所有路由。已改成每请求一线程（上限 16）。这是闭环 A 遗留的设计缺陷，被闭环 B 暴露出来。

## 目前还没做的部分（留给下一轮）

* **真实写操作工具。** 审批通道已验证可用，但本轮**只**放了无害的 `write_test_marker`。要接真正会改 lamber 数据的工具（改方案、生成文档等），下一轮把它们加进 `approval.ts` 的 `GATED_TOOLS`**以及 `approval.rs` 的镜像表**（两处，有测试卡着），并逐个补桥接路由与端到端用例。
* **dsh 自带工具的审批。** ACP 下 lamber 现在是**唯一**的权限应答方，`acp` profile 里 dsh 自带的工具（bash、文件编辑等）若触发权限请求，也会弹到同一个弹窗；索引里没有对应公告时，弹窗会以调用 id 命名并给通用文案。目前是"照问不误、从不默许"，还没有按工具分类的策略。
* **阶段 3A 已补齐逐 token 展示。** 插件 firehose → 鉴权显示桥接 → 按步骤预览，最终以 ACP 提交替换。2026-09-07 已删除 `AiRuntime.ts` 和临时回退入口；3A 记录保留历史验证背景。
* **审批弹窗的产品化。** 现在是能演示流程的最简实现：单一模态、JSON 原样展示参数、倒计时。没有做"记住选择""按工具批量授权""历史审批记录"，也没有按 DESIGN.md 精修视觉。
* **干净 Windows 装机 gate。** production runtime、内置 Node、可写用户 home 和设置页已在本地资源树验证握手，但尚未在“无仓库、无系统 Node、无环境变量”的 Windows 机器安装 NSIS 并用界面密钥取得真实回复。按 2026-09-05 拆分决定，它作为 Gate 1B 必须在阶段 3 前补齐。
* **单文件 SEA。** 阶段 0 产物约 113.1 MiB，但启动即崩溃，且 dsh 的动态包/插件加载仍需外部资源。在不改 dsh 源码的边界下不可用；发布形态已选择锁定的 npm production tree + 内置 Node。

## 依赖说明

ACP 重写新增两个直接依赖：`agent-client-protocol 2.0.0`（ACP 客户端，传递依赖
`agent-client-protocol-schema 1.5.0`）与 `tokio 1`（仅
`rt-multi-thread`/`sync`/`time`/`io-std`/`io-util` 五个 feature）。tokio 只服务于
`dsh_session.rs` 这一层，后端其余部分仍是同步的。

闭环 A 时期已有的直接依赖：`tiny_http 0.12`（`default-features = false`，仅 `ascii`/`chunked_transfer`/`httpdate`/`log` 四个小传递依赖，不含 TLS）。

选它而不是 hyper/axum 的原因：lamber 后端整体是同步的（`rusqlite` + `Arc<Mutex<Connection>>`，`calculate_ict_benefit` 是同步函数），在异步 handler 里做同步加锁是反模式；桥接服务只有一个回环路由、并发量极低，用阻塞式服务器跑在独立线程上与既有代码风格一致。也没有手写 HTTP 解析——那是比引入一个小依赖更大的技术债。


## 阶段 2 产品链路（2026-09-05）

- 产品入口：`AiChatPanel` 的模型设置中启用 dsh，默认关闭。联调台继续用于排障，非产品聊天入口。
- AI 浮窗挂载原审批对话框，事件只向一个承载窗口投递，避免漏弹或重复；审批机制与工具集合不变。
- `DshRuntime.ts` 先监听再 invoke；Rust `turns.rs` 在排队前绑定前端 sessionId/requestId，避免早到/迟到事件串话。
- `ai_cancel_prompt` 发送标准 `session/cancel`，匹配请求 id 防止旧停止操作取消新轮次；关闭子进程会终结未完成轮次。
- `session_store.rs` 使用独立 app-data SQLite 保存前端 id、ACP id 和 cwd；`session/resume` 复用原 id。
  错误 cwd、缺失会话或映射冲突明确报错，不新建兜底。恢复后通过标准 config option 同步模型。
- 清空/删除由用户点击触发，停止活动轮次后关闭 ACP 会话、清除映射。dsh 磁盘历史不自动删除。
- `prompt.rs` 构造 text/image blocks，检查数量、mime、base64 和大小；握手能力不支持图片时明确拒绝。

### 为什么上下文通过 text 注入

继续使用 `buildAiChatContext` 和 `PromptRenderer`，在每一轮发送最新已保存状态与未保存草稿。
已检查 alpha.5 的 `agentOptions(config)`、`newSession`、`resumeSession`、`prompt`：
标准 ACP 参数没有暴露每会话 system/instructions 更新入口；进程级补丁属于所有会话的共享配置，
不适合不同项目/页面的动态数据。因此选择任务书允许的 text block 注入，保持现有来源边界和元单位规则。
不写全局 system 文件，不修改 dsh 源码或业务数据。

### 实测限制

`assistantUpdates` 由已提交的 `assistant/message` 驱动，名称虽然是 chunk，当前不逐 token 推送。
阶段 3A 已另行通过插件火线事件补齐增量，见下节。真实取消验证用透明本地代理观察官方 API 的 delta 后取消，11–18ms 得到 `Cancelled`，下一轮正常。
重启恢复记忆与视觉模型识图均通过；完整能力、已知缺失、人工验收范围见
[阶段 2 对照表](../docs/verification/dsh-stage2-product-integration.md)。

## 阶段 3A 实时显示（2026-09-05）

`dsh-tool-lamber/src/stream.ts` 在 `LAMBER_STREAM_DISPLAY=1` 时启用。
公开 `agent/pre-step` hook 等待 `/lamber-bridge/stream` 绑定前端 session/request，
随后订阅 `session/event` 的 assistant/chunk，顺序发送 text/reasoning/tool-call 字节增量，
以及 assistant/message 的 messageId、turn、step 关联。只处理显示，不执行业务写入。
Rust 校验绑定不变后发 `session/stream`，前端按步骤关联并以 ACP 为最终权威。

不要给独立旧桥接用例开启此环境变量，除非它提供 bind 和显示路由；模型前绑定失败会明确报错。
产品运行时显式开启，仍需用户手动打开默认关闭的 dsh 聊天开关。

升级契约测试：仓库根目录 `npm run test:dsh-stream`。这会运行实际 dsh/profile/plugin，
确认多条增量在提交前显示并由生产前端适配器验证最终文本一致；订阅失效会失败，不能跳过。
真实模型与产品流程、旧运行时接替表、模板确认页已知问题见
[阶段 3A 验证](../docs/verification/dsh-stage3a-streaming-and-handover.md)。


### 会话项目权限（2026-09-06）

`projectScope.ts` 使用运行时注入的 `exec.agent.session.id`，不接受模型提供会话身份。
Rust `ProjectBindings` 在 prompt 前恢复 `ai-sessions.sqlite` 中的绑定，独立于每轮 `Turns`。
`/lamber-bridge/authorize` 先核验权限，`/calculate` 在真实业务读取前再次校验；未知工具及通用会话全部拒绝。
新建/历史会话策略、工作区隔离与限制见 [模块设计](../docs/modules/ai-session-workspace.md#项目权限路线图②)，
真实调用、重启恢复、并发及缺失身份的证据见 [验收记录](../docs/verification/session-project-binding.md)。
`check-bridge.mjs` 的第三个参数是由测试夹具登记的 ACP 身份，只用于模拟工具执行上下文；生产工具不允许从参数取得身份。
新增插件回归命令：`npm run test:scope --prefix agent-bridge/dsh-tool-lamber`。


### 跨项目查询（路线图③）

新增 `query_projects` / `/lamber-bridge/query-projects`，是唯一聚合只读白名单工具。通用聊天与项目会话均可调用；其余工具权限不变。
复用 `ProjectService.get_projects`，仅投影项目行与 summary_metrics，不 JOIN 明细，不输出本机路径/note/logs。
支持参数化筛选、排序，默认20/最多50条；结果明示总命中数、截断、全体金额合计与 isBoundProject，防止跨项目数字混用。
此处修订②原“通用会话全部拒绝”的描述，仅新增该工具类别。[设计](../docs/modules/ai-project-query.md)与[验收](../docs/verification/cross-project-query.md)。

## 桥接契约与启动就绪（2026-09-06）

`bridge-contract.json` 是契约版本和能力路由清单的唯一权威来源。插件 build/typecheck
先生成 `src/contract.generated.ts`，再编入 `lib/contract.generated.js`；Rust 使用
`include_str!` 把同一文件编入二进制。选择构建快照而不是让插件读取 Rust 环境版本，
是为了让“旧程序 + 新插件”真正暴露差异，避免两端重复报告同一个新值而假通过。
改变不兼容的路由语义时必须提升 version；新增必需路由时更新清单。

插件 `apply` 先向鉴权后的 `/lamber-bridge/handshake` POST 自己的版本和必需路由。
Rust 不匹配返回 409；只有握手成功才注册工具与原有权限、审批、流式 hook。
插件注册完成写一条 `LAMBER_BRIDGE_STARTUP:ready` 到已有 stderr 诊断通道，
失败写 mismatch/unreachable 后抛出启动错误。Rust 消费明确的启动回执，**先完成
插件就绪、后发送 ACP initialize**，再沿用 `AcpRuntime::start -> Result` 失败通道。
这是必要的生命周期同步：上游 ACP 会在其他 Cordis 插件仍加载时接受请求，单靠
`async apply` 抛错并不能保证调用方未创建会话。15 秒是缺少回执时的启动截止时间，
不是重试或等待后放行；HTTP 请求最多 5 秒，一次失败即拒绝。

仅握手请求把旧服务的 404 识别成版本错配，所有既有路由路径、404 行为和业务语义
不变。业务 HTTP 错误保留后端 error 正文，去掉模型可见的 URL/状态码包装。
审批策略、会话限权、聚合白名单和财务引擎均不参与握手。

构建前先 `npm run build --prefix agent-bridge/dsh-tool-lamber`。Cargo build script
同时检查开发插件与存在的 staging 插件的**编译输出**是否匹配源契约，避免仅刷新
清单但忘记编译。Windows 打包改为 build --no-bundle → 从实际目标二进制读取编译
契约并与 staging lib 比对 → bundle；不运行目标程序，因此不依赖 GUI stdout 或目标
平台可执行性。构建与插件校验无需新增依赖。

验证：`npm run test:contract --prefix agent-bridge/dsh-tool-lamber`、
`npm run test:packaging`、`cargo test` 和真实 Key ignored 套件。
旧二进制的显示端没有新错误识别能力，实测会在启动中断时显示 ACP disposed；
完整升级后才有人话文案，不能把只替换插件当成完整升级。
[完整证据与旧版限制](../docs/verification/bridge-contract-handshake.md)。

## 2026-09-07：长文本审批与契约 v2

审批 UI、可选修改参数、双版本审计扩展已实现；当时 `GATED_TOOLS` 仅有 `write_test_marker`；现已接入阶段 B，见下节。
ACP permission 仍是 grant/reject；Rust 在批准时按可信 sessionId/callId 保存一次性参数交接，插件执行前通过 `/lamber-bridge/reviewed-arguments` 取实际批准值。
交接复核绑定/工作区/工具名/原始参数，只允许消费一次，缺失或不匹配拒绝；因此不能只把前端编辑显示成成功而继续执行原参数。
新增路由纳入唯一契约源，版本升至2；旧二进制/插件组合在启动时拒绝，必须完整重建应用资源。
审计复用现有 args_json 的版本化对象，无 DB 迁移。设计见 [ai-approval-review.md](../docs/modules/ai-approval-review.md)，阶段 A 真人 gate 见 [验收记录](../docs/verification/ai-runtime-retirement-stage-query-approval-a.md)。

## 2026-09-07：模板文本写入与契约 v3

新增 fill_template_fields 与 /lamber-bridge/fill-template-fields，GATED_TOOLS 两端同步登记。需求表共享 JSON 目录派生插件参数和 Rust 白名单，仅支持 8 项文本；项目绑定/通用拒绝/审批前置不变。

工具执行只提交可信 session/call/原参数，由 Rust 消费已批准值并复核旧值与版本，再调用原模板保存事务。reviewed-arguments 路由只供 marker 使用。契约升至3；必须同步构建开发插件、staging 与 Rust 应用。

[模块设计](../docs/modules/ai-template-write.md) / [验证及 Gate B 真人步骤](../docs/verification/template-write-b.md)。前面“仅测试写工具”的叙述是历史阶段状态，以本节为准。

## 2026-09-07 补充：主动模板读取与契约 v4

新增 read_template_fields(templateId) 和 /lamber-bridge/read-template-fields，工具总数5。只读无需审批，项目取自可信会话绑定，通用及额外projectId拒绝。原GATED_TOOLS不变，bash/glob不放行。

插件构建直接编译UI的demand.ts及共享JSON；Rust返回白名单保存值和完整缺项依据，插件执行同一纯函数输出缺项。总文本24000字/清单100行上限与截断信息明确返回，不暴露路径，不改被动页面注入或模板保存服务。

契约v4，开发/staging插件及Rust应用须同步构建。新增 `npm run test:template-read --prefix agent-bridge/dsh-tool-lamber` 使用常规Rust测试产出的投影证据验证跨语言链路；先运行cargo test。[验证记录](../docs/verification/template-read.md)。
