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

一条原则：**dsh 只做编排，一行业务数学都不在 TypeScript 里。** 插件是桥接服务的瘦客户端，桥接服务是既有 Rust 模块的瘦转发层。

## 目录结构

| 路径 | 作用 |
| --- | --- |
| `dsh-tool-lamber/` | 自定义工具插件（独立 npm 包，`defineTool` 注册 `run_benefit_calculation`） |
| `patch.yml` | dsh profile 补丁层，把插件挂进 `sdk` profile |
| `scripts/provision-profile.mjs` | 一次性准备：构建插件 + `dsh plugin add` 链接进 `$DSH_HOME/profiles/sdk/` |
| `scripts/check-bridge.mjs` | 只跑「插件工具体 → 桥接服务 → calculator」这一跳的排错脚本，不启动 dsh、不消耗 LLM |
| `../src-tauri/src/agent_bridge/bridge_server.rs` | 仅监听 `127.0.0.1` 的 HTTP 桥接服务（tiny_http） |
| `../src-tauri/src/agent_bridge/calculation.rs` | `POST /lamber-bridge/calculate` 路由：项目 → 方案 → 最新快照 → 测算引擎 |
| `../src-tauri/src/agent_bridge/dsh_session.rs` | dsh 子进程管理 + JSON-RPC 2.0 客户端 |
| `../src-tauri/src/agent_bridge/mod.rs` | Tauri 命令、`AgentRuntime` 生命周期、事件转发 |

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

**关于令牌**：只绑定回环地址不等于做了鉴权——同机任何进程都能访问该端口，进而读到客户项目的财务数据。因此桥接服务对每个请求校验一次性令牌（定长比较），令牌通过环境变量交给子进程，不落盘。

## 怎么跑测试

分层设计，失败时能直接指认是哪一半坏了。

```bash
cd src-tauri
cargo test agent_bridge
```
默认跑的 6 个用例**不需要 Node、不需要网络**：临时 SQLite 工作区 + 桥接服务 + 计算路由，覆盖默认方案 / 按阶段 / 按方案名 / 按方案 id 选择、未知 scenario 拒绝、以及令牌鉴权。

```bash
cargo test agent_bridge -- --ignored --nocapture
```
需要先完成上面的 provision。三个用例：

1. `plugin_tool_body_reaches_the_calculator_over_the_bridge` —— 起真实桥接服务，用 `node scripts/check-bridge.mjs` 直接调用插件的 `execute()`，断言回来的 NPV／利润率与引擎直算结果一致。**不需要 API key**，验证的是插件↔Rust 传输契约。
2. `dsh_advertises_the_lamber_tool_in_its_request_header` —— 真实拉起 dsh 子进程，`initialize` → `session/prompt`，等 `request/header` 通知，断言 `run_benefit_calculation` 出现在工具目录里。**不需要 API key**，这是判断「插件加载 / 协议握手 / schema 注入」是否正常的中间检查点。
3. `dsh_tool_call_reaches_the_calculator_and_returns_real_numbers` —— 完整闭环：模型发出 `tool/call` → 桥接服务收到 HTTP 请求 → `tool/result` 里带回 calculator.rs 算出的真实数字。**需要 `DEEPSEEK_API_KEY`**；未设置时用例自行跳过并打印提示。

排错顺序建议：先 `check-bridge.mjs`（是不是桥接坏了）→ 再 `dsh_advertises...`（是不是插件/协议坏了）→ 最后带 key 的完整闭环（是不是模型/提示词的问题）。

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

## 目前还没做的部分（留给下一轮）

* **闭环 B：审批通道。** dsh 的审批是**进程内 Cordis 事件**，`approval/request` **不会**通过 SDK 的 JSON-RPC 协议自动转发给外部客户端。要打通 `dsh → answerer 插件 → Rust → React 弹窗 → 用户确认 → dsh`，需要在 dsh 侧另写一个 answerer 插件把请求转出来。目录位置预留在 `dsh-tool-lamber/` 同级（建议新建 `dsh-answerer-lamber/`），`patch.yml` 里再加一条 `insert`。在此之前**不要**往插件里加任何会写数据的工具——AGENTS.md 明确禁止 AI 绕过用户确认写 `projects_store.json`。
* **前端展示。** `AiChatPanel.tsx` / `AiRuntime.ts` 本轮未改动。Rust 侧已经把所有通知按 `{method, params}` 打包 emit 到 `ai://session-event`，前端订阅这一个事件、按 `method` 分派即可，新增通知类型不需要再加 Tauri 事件名。
* **API key 来源。** 目前从环境变量读。前端的 key 存在 `localStorage.lamber_ai_api_key`，接入时应改为由前端传入或落到 lamber 自己的凭证存储，而不是依赖环境变量。
* **单文件 SEA 打包 / 瘦身。** 现在依赖开发机上的 `agent-bridge/node_modules`（约 520 个包），发布形态未处理。

## 依赖说明

Rust 侧新增一个直接依赖：`tiny_http 0.12`（`default-features = false`，仅 `ascii`/`chunked_transfer`/`httpdate`/`log` 四个小传递依赖，不含 TLS）。

选它而不是 hyper/axum 的原因：lamber 后端整体是同步的（`rusqlite` + `Arc<Mutex<Connection>>`，`calculate_ict_benefit` 是同步函数），在异步 handler 里做同步加锁是反模式；桥接服务只有一个回环路由、并发量极低，用阻塞式服务器跑在独立线程上与既有代码风格一致。也没有手写 HTTP 解析——那是比引入一个小依赖更大的技术债。
