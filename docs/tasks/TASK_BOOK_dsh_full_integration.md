# 任务书：dsh 完全融合（全面替换 Chat 链路 + 打包内可用）

> **2026-09-10 接续说明：** 当前主任务已切换为 [AI UI 整体升级改造：全面采用 dsh 官方 WebUI](./TASK_BOOK_ai_webui_upgrade.md)。本书保留此前引擎接入的施工与验证历史；产品目前已唯一使用 DshRuntime，AiRuntime 已删除，工具和业务卡片也已接入。下文关于“仅实验台可用”“两套密钥”“无持久化”等开工现状及排序均属于原轮次，不能当作当前事实或重复执行依据。原有尚未完成的配置、Windows 安装和用户真人 gate 仍须逐项核对，不因接续任务自动通过。

> **排序**：本文排在 `docs/tasks/TASK_BOOK_cowork_session_project_binding.md`、
> `docs/tasks/TASK_BOOK_create_intelligent_compute_project.md`、
> `docs/tasks/TASK_BOOK_demand_analysis_image_completion.md` **全部之前**。
> 那三份都假设"dsh 已经是产品里能用的东西"，而现在它不是——它只在开发机的仓库里、
> 通过一个隐藏路由 `#/agent-lab` 能跑。前提没落地之前不要开那三份。

## 本次已经拍板的两个决定

1. **dsh 全面替换现有 Chat 链路**（`src-ui/src/ai/AiRuntime.ts` 那条 OpenAI 兼容 SSE），不做长期并存。
2. **打包安装版里 dsh 必须能跑**，不是只在开发机跑通就算完。

---

## 开工前要认下来的两条代价

> **更正说明**：本文初稿曾把"换服务商"和"图片输入"列为替换的两条净损失。**这两条都是错的**，
> 已按 dsh 源码更正，见"关键事实 5 / 6"。前者是 dsh 明确设计的扩展点，后者是模型选型问题，
> 都不构成放弃替换方案的理由。下面只保留真实存在的两条代价。

1. **上下文注入方式要重做。** 现在是前端 `buildAiChatContext` + `PromptRenderer` 把项目上下文拼进 messages 数组；
   ACP 的 `session/prompt` 在 Rust 侧只收一个 `text`（`dsh_session.rs` 的 `Command::Prompt`）。
2. **安装包体积会明显变大**（详见"关键事实 1"，`agent-bridge/node_modules` 现在 309M）。

---

## 已确认的关键事实（写代码前不用重新调研）

### 1. 打包安装版里根本没有 dsh —— 这是最致命的一条

- `repo_root()`（`src-tauri/src/agent_bridge/mod.rs:270-284`）靠从可执行文件向上找
  `agent-bridge/patch.yml` 来定位；找不到就报"未找到 agent-bridge 目录"。
- `DshLaunchConfig::from_repo_root()`（`dsh_session.rs:121-135`）把 `dsh_bin` 指向
  `agent-bridge/node_modules/.bin/dsh`，`dsh_home` 指向 `agent-bridge/.dsh-home`。
- `src-tauri/tauri.conf.json` 的 `bundle` 段**只有 `active` / `targets` / `icon`，没有 `resources`**；
  `scripts/package-windows.mjs` 里也没有任何 agent-bridge 相关处理。
- `agent-bridge/package.json` 里 `@deepseek-ai/dsh` 是 **`devDependencies`**，字面意义上的开发期依赖。
- `dsh` 是 Node CLI，跑起来还需要 Node 运行时（本机 v22）。
- 体积基数：`agent-bridge/node_modules` = **309M**，`.dsh-home` = 580K。

**结论：NSIS 装出来的应用，第一次发消息就是"未找到 dsh 可执行文件"。目前所谓的"接入"只在开发机的仓库目录里成立。**

### 2. API key 有两套，dsh 那套用户碰不到

- Chat 模式：`endpoint` / `model` / `apiKey` 存 localStorage，面板里可改（`AiChatPanel.tsx:84-86, 160-169`）。
- dsh 模式：key 只在**应用进程启动时**从 `DEEPSEEK_API_KEY` 环境变量读一次
  （`dsh_session.rs:131`），provider / model 硬编码为 `deepseek-official` / `deepseek-v4-flash`（`:129-130`）。
- 已知可用事实（README 已验证）：dsh 接受 `DEEPSEEK_API_KEY` 环境变量注入，不必走它自带的 credentials service；
  空值等同未设置（`dsh_session.rs:153-160` 的注释解释了为什么用空串而不是 `env_remove`）。

**结论：装机用户没有任何界面能配 dsh 的 key。替换之后，"配 key"必须变成产品功能，不能再是开发者的环境变量。**

### 3. dsh 在产品界面上不存在

- `ai_send_prompt` 的**唯一**调用方是 `src-ui/src/components/ai/AgentLabView.tsx:113`。
  该文件自己的文档注释写着："Deliberately plain — this is a lab bench, not the product surface."
- 入口是隐藏路由 `#/agent-lab`（`App.tsx:31`），正常用户点不到。
- 真正的 `AiChatPanel` 走 `AiRuntime.ts` 的 `fetch` + SSE（`AiRuntime.ts:81-89`），和 dsh 毫无关系。
- `harnessSessionId` 至今是纯元数据，`sessionTypes.ts:12-19` 的注释直说了这一点。

### 4. 流式输出没有适配器

- `dsh-acp` 会发这些 `session/update`：`agent_message_chunk`、`agent_thought_chunk`、`tool_call`、
  `tool_call_update`、`usage_update`、`config_option_update`。
- Rust 侧把它们**原样**转成一条 `ai://session-event`（`mod.rs:216-226`），不做任何解释。
- 前端只有实验台把它 `JSON.stringify` 打进日志（`AgentLabView.tsx:73-74`）。
- **没有任何代码把 `agent_message_chunk` 拼成 `AiChatMessage`。** 这一层是替换必须新写的。

### 5. 图片能力是模型选型问题，不是能力缺失

- `dsh-acp` 的 `initialize` 返回 `agentCapabilities.promptCapabilities.image`，由 `supportsAcpImagePrompts()` 算出
  （`agent-bridge/node_modules/@deepseek-ai/dsh-acp/lib/index.js:78-88`），要求**当前配置模型**的
  `inputModalities` 含 `"image"`；不满足时带 image block 的 prompt 会被 `admitAcpPrompt` 拒掉（同文件 `110-112`）。
- **dsh 自带的默认模型目录里就有视觉模型**：`deepseek-v4-flash-vision-exp`，
  `inputModalities: ["text", "image"]`，并带 `imagePixelBudget` / `imageMaxBytes`
  （`@deepseek-ai/dsh-llm-deepseek/lib/index.js:1839-1846`）。
- lamber 现在硬编码的是 `deepseek-v4-flash`（`dsh_session.rs:130`），该条目没声明 image——
  **`promptCapabilities.image` 为 false 的话，原因在这里，不在 dsh。**
- 模型目录本身也是配置项：`Config.models`（`z.array(catalogModel).default(DEFAULT_MODELS)`，
  `dsh-llm-deepseek/lib/index.js:1873`），可以自行声明模型 id 与其 `inputModalities`。
- Rust 侧现在把握手答案丢掉了：`AgentHandshake::from_response`（`dsh_session.rs:182-190`）只留了
  `protocol_version` 和 `agentInfo`，没读 `agent_capabilities`。**读出来是一行的事。**

**结论：图片输入不是替换的阻塞项，是"配哪个模型"的问题。** 但模型必须做成可配（见阶段 1 第 5 条），
不能继续硬编码——否则用户永远只能用到硬编码那一个模型的能力。

### 6. 换服务商是 dsh 明确支持的，有三条路

`@deepseek-ai/dsh-llm` 的包描述原话是 "Provider-neutral LLM service interface for the DeepSeek Harness"。
从轻到重三条路，**优先验证第 1 条**：

1. **只改 baseURL**（最轻，可能零代码）：`dsh-llm-deepseek` 的 `Config.baseURL` 是一等公民字段
   （`lib/index.js:1861`），解析顺序为 `config.baseURL ?? $DEEPSEEK_BASE_URL ?? "https://api.deepseek.com"`
   （`:1965`），其 HTTP 客户端自述是 OpenAI-compatible（`:604`）。
   **注意**：`BASE_URL_ENV` 的源码注释写着 "honored only from trusted layers"，
   走配置层比走环境变量更稳妥。
2. **自己写 provider adapter**（中等）：`LlmAdapter` 基类由 `@deepseek-ai/dsh-llm` 导出，
   注册接口 `ctx.llm.registerAdapter(providers, adapter)` 是公开且带文档的扩展点
   （`dsh-llm/lib/index.js:1242`，基类文档在 `:1075-1079`）。挂载方式与现有 `dsh-tool-lamber`
   插件一样走 `patch.yml`，不需要改 dsh 本体。
3. 改 dsh 源码 —— **不要走这条**，前两条已经够用，改源码会让后续升级无法跟随上游。

**~~待验证的风险（第 1 条路）~~ → 阶段 0 已实测，第 1 条路被证伪：**
严格 OpenAI 端点用 HTTP 400 拒绝了 `thinking`、`reasoning_effort`、`dsh_plugin_packages`
三个 dsh 自有字段（记录见 `docs/verification/dsh-stage0-distribution-validation.md`）。
**所以"只改 baseURL 就能换服务商"不成立。** baseURL 配置项已在阶段 1 做出来，
但仅验证过 DeepSeek 官方端点，UI 不得宣称支持 Ollama 或任意 OpenAI-compatible 服务。
真要通用服务商，只剩第 2 条路（自写 adapter），需另开任务书解决这三个字段的投影问题。

### 7. 不能中断

`AcpRuntime` 的 `Command` 枚举只有 `NewSession` 和 `Prompt`（`dsh_session.rs:194-206`），没有 cancel；
而 `dsh-acp` 侧是有 `cancelPrompt` 的（`lib/index.js:821` 附近）。
Chat 模式现在有 Abort、能停止生成（多会话工作区那一轮已验证过"停止后流立即中断"）。
**替换后如果没有 cancel，就是从"能停"退化成"发出去只能等"。**

### 8. 工作目录是仓库根，不是用户工作区

每个 ACP 会话都用 `agent.config.cwd` 开（`mod.rs:151`），而 `cwd` = `repo_root`（`dsh_session.rs:128`）。
dsh 自带的文件/bash 工具会在 **lamber 源码目录**里干活，而不是用户的项目工作区。
打包之后这个路径根本不存在，`session/new` 会拿到一个无效 cwd。

### 9. 会话映射只在内存里

`RunningAgent.sessions: HashMap<String, String>`（`mod.rs:148-157`）不持久化。
应用一重启，所有前端会话都会重新 `session/new`，`$DSH_HOME/sessions/` 里的历史用不上。
现在实验台每轮都用 `lab-${Date.now()}`，掩盖了这个问题。

> 附带澄清一条**过时信息**：README 里"复用同一个 sessionId 会报 id collision"是 SDK 协议时期的坑，
> ACP 下会话 id 由 dsh 自己生成，该坑已不存在（`agent-bridge/README.md:267-269` 已改正）。
> `send_prompt` 现在确实会复用映射，不用再为此设计规避方案。

### 10. 工具面只有两个，且都不写业务数据

只有 `run_benefit_calculation`（只读）和 `write_test_marker`（无害测试）。
即使界面接通，用户问出来的东西也很有限——但**扩工具不在本次范围**（见"不要做的事"）。

### 11. 逐 token 流：ACP 层丢弃了它，但插件能自己订阅（2026-09-05 二次排查更正）

> **更正**：本条初稿的结论是"只能等上游或接受降级"。**那是不完整的**——
> 当时只查了 `dsh-acp` 及其 config，没查会话事件核心。二次排查发现第三条路，见下。

**准确表述：token 级增量不但存在，而且已经进了会话事件流；是 ACP 层收到后丢弃了它。**

- **每个 chunk 实时进事件流**：`@deepseek-ai/dsh-agent-loop/lib/index.js:626-633`——
  ```js
  for await (const chunk of stream) {
      chunkSeqs.push(this.session.append("assistant/chunk", { turn, step, chunk }).seq);
      assembler.push(chunk);
  }
  ```
  即 LLM 流的每个 delta 都在**消息提交之前**被 append 成 `assistant/chunk` 事件。
- **ACP 层收到了但不处理**：`dsh-acp/lib/index.js:1105` 订阅的是
  `ctx.on("session/event", ...)`，**全量事件**；`onSessionEvent`（`:880`）的分支只处理
  `assistant/message` 和 `tool/call`，`assistant/chunk` 被静默丢弃。
- **`session/event` 是公开事件**：`dsh-session/lib/types/index.d.ts:64` 有类型声明，
  `:127` 称其为 "firehose"，`:4` 写明用法 "subscribe to `session/event`, drain on `session/flush`"。
- **lamber 的插件已经有 ctx**：`dsh-tool-lamber` 的 `apply(ctx: Context)` 拿到的正是
  Cordis `Context`，且已用它挂了审批守卫。**订阅这条 firehose 不需要改 dsh 源码。**
- `dsh-acp` 的 config 仍无流式开关（只有 `provider` / `model` / `sessionListPageSize`）——
  但这已经不重要了，因为不必走 ACP 那条路。

#### 可行的架构：增量走桥接，提交走 ACP

插件订阅 `session/event` → 过滤 `assistant/chunk` → 经现成的 loopback bridge
（`bridge_server.rs`，已有令牌鉴权）推回 lamber，**只用于显示**；
已提交的消息仍由 ACP 的 `agent_message_chunk` 送达，**它是权威版本**。
这跟现有 `dsh-tool-lamber → postBridge → bridge_server` 是同一个 seam，不是新架构。

#### 代价与风险（要认下来再做）

1. **依赖 dsh 内部事件名**：`assistant/chunk` 不是 ACP 契约的一部分，
   dsh 升级可能改名或改结构。必须写一个**订阅不到就大声失败**的测试，
   不能让它悄悄退化成"又不流式了但没人发现"。
2. **两条传输、一条消息**：增量走 HTTP 桥接，提交走 stdio ACP，**两者之间没有顺序保证**。
   适配器必须以提交版本为准**替换**累积的增量文本，而不是追加——否则会出现重复正文。
3. **UTF-8 拆字问题回来了**：关键事实 11 初稿说"结构上不可能发生"，那个结论**绑定 commit 边界**；
   一旦改吃真实 delta，字符被切在两个 chunk 中间就重新可能，适配器必须按字节缓冲处理。
4. **`chunk` 是 LLM 层结构**（`text-delta` / `reasoning-delta` / `tool-call-delta`），
   不是 ACP 的 ContentBlock，需要自己做映射。

#### 上游 issue 仍然要提

草稿见 `docs/tasks/upstream-issue-acp-incremental-output.md`。
理由变了：不再是"我们被卡住了"，而是**"你们已经发了 `assistant/chunk`，ACP 层只是没投影"**——
这个 ask 具体得多，也更容易被接受。上游一旦支持，上面这个 workaround 连同它的三条风险就能删掉。

---

## 分阶段要做的事

每个阶段有硬性 gate。**上一阶段的 gate 没过，不要开下一阶段。**

### 阶段 0 · 四个前提验证（不写产品代码，只查事实）

1. 把 `agent_capabilities.prompt_capabilities.image` 在 `AgentHandshake` 里保留下来并记录真实值，
   **分别在配置 `deepseek-v4-flash` 和 `deepseek-v4-flash-vision-exp` 两种情况下各测一次**。
   预期是前者 `false`、后者 `true`；若与预期不符，说明对关键事实 5 的理解有偏差，停下来查清再往下走。
2. **验证换服务商的第 1 条路**（关键事实 6）：把 `dsh-llm-deepseek` 的 `baseURL` 指向一个
   OpenAI 兼容服务（本地 ollama 即可），确认能否完成一次正常问答。
   重点看它自有的 `thinking` / `reasoningEffort` 参数和 files API 会不会让对方报错。
   跑通 → 后续"自选服务商"是纯配置；跑不通 → 记录**具体是哪个字段/哪个调用**被拒，
   作为将来写自有 adapter 的输入。**本阶段不写 adapter。**
3. 确认 dsh 在**没有仓库目录**的环境下能不能起来：至少手工把 `node_modules/.bin/dsh` + `.dsh-home` +
   Node 运行时复制到一个临时目录，用 `LAMBER_REPO_ROOT` 指过去跑通一次握手。
4. 量出分发形态的真实体积：完整 `node_modules`（309M 基数）、`pnpm prune --prod` 之后、
   以及 SEA 单文件（CURRENT_TASK 记的"未做 SEA 单文件打包 / 瘦身"就是这一项）三者各是多少。

**Gate：**
- 第 1 条两次结果都要记录。**`deepseek-v4-flash` 为 `false` 是预期内的正常结果，不是阻塞**，
  处理办法是把模型做成可配（阶段 1 第 5 条），不是停工。
- 第 2 条**跑不通也不阻塞**——它只决定"自选服务商"是配置项还是要写 adapter，
  两条路 dsh 都支持。把结论如实记下来即可。
- 第 3 条若跑不通 → **停下来汇报**，说明 dsh 存在只能靠仓库运行的硬依赖，分发方案要重想。
  这是本阶段唯一真正的阻塞点。

### 阶段 1 · 让 dsh 在装机版里能起来

1. **定位方式改造**：`repo_root()` 从"向上找 `agent-bridge/patch.yml`"改成分层查找：
   打包资源目录 → `LAMBER_REPO_ROOT` 环境变量（保留，开发用）→ 开发期仓库布局。
   三条路径都要有明确的失败文案，**不要让用户看到"请先在 agent-bridge/ 目录运行 npm install"这种开发者话术**。
2. **打包**：`tauri.conf.json` 加 `bundle.resources`，`scripts/package-windows.mjs` 里加对应的准备步骤
   （prune / 复制 / 校验）。按阶段 0 第 4 条的实测结果选形态，并把选择理由写进
   `docs/modules/release-packaging.md`。
3. **首次运行初始化**：`.dsh-home` 是可写目录，不能直接用只读的安装目录。
   改成安装时把模板 `.dsh-home` 复制到用户数据目录（`app_data_dir`），首次运行时初始化。
4. **cwd 改成用户工作区**：`session/new` 传当前打开的 workspace 路径，不再传仓库根（对应关键事实 8）。
   工作区未打开时不要退回仓库根——那个目录在装机版里不存在；这种情况应当明确拒绝并提示用户先打开工作区。
5. **模型与 key 配置搬进产品**：新增设置项（沿用现有设置面板体系），至少覆盖 **key、模型、baseURL** 三项，
   把 key 存进现有配置/凭据存储；`DshLaunchConfig` 从配置读，环境变量降级为开发期兜底。
   - **模型必须可配**，不能继续硬编码 `deepseek-v4-flash`（关键事实 5）——否则图片能力被一个常量锁死。
     模型下拉的候选来自 dsh 的模型目录，其中视觉模型要能被选到。
   - **baseURL 做成可配**（关键事实 6 第 1 条路）。阶段 0 第 2 条若验证通过，这就是"自选服务商"的完整实现；
     若验证不通过，本阶段仍然把字段做出来，但在 UI 上标明当前仅验证过 DeepSeek 官方端点，
     **不要**假装支持一个没测过的东西。
   - key 改动后需要重启 dsh 子进程才生效——这条要在 UI 上说清楚，或者直接在保存时重启
     （`ai_agent_stop` 已经能停，下一条 prompt 会重新拉起）。模型 / baseURL 同理。
   - **`Debug` 脱敏（`redacted()`）现有实现不要动。**

**Gate（2026-09-05 拆分）：** 原 gate 是"干净 Windows 机器装 NSIS 包跑通"。
因当前开发机是 macOS，按"是否真的与 Windows 有关"拆成 1A / 1B 两半：
**1A 必须过才能进阶段 2；1B 可挂起，但必须在阶段 3 之前补上。**

**Gate 1A · 与平台无关，mac 上即可验证（阻塞阶段 2）**

1. 未打开 workspace 时的拒绝路径走一次，文案明确、不退回仓库根。
2. 设置页保存 key / 模型 / baseURL 后，dsh 子进程重启并生效。
3. 安装资源缺件时的用户文案（把 staging 树里的文件删一个来触发），不出现开发者话术。
4. 确认运行的是分发内 Node 而非系统 Node（把系统 node 从 PATH 摘掉后仍能起来）。
5. **补跑阶段 1 中因缺 `DEEPSEEK_API_KEY` 被跳过的 3 个真实 LLM 用例。**
   这条与 Windows 无关，且实为**阶段 2 的前置**——见下方"阶段 2 的开工前提"。

**Gate 1B · 真·Windows-only，可挂起**

- NSIS 文件名 / SHA-256 / 安装包体积；能否正常安装
- Windows 路径形态：带空格路径、中文用户名路径、反斜杠
- 杀毒软件对随包 `node.exe` 的拦截情况
- Windows 下子进程 spawn 行为
- 首次启动至设置页可操作耗时；配 key 后发消息取得首条回复

> **不要用 mac 的 .app 构建冒充 1B。** Tauri 的 `resource_dir()` 在 mac 上是
> `.app/Contents/Resources`，在 Windows NSIS 上是安装目录，路径形态不同。
> mac 上验过只能证明 `distribution.rs` 三级查找的**逻辑**成立，
> 证明不了 Windows 的资源解析成立。`docs/verification/dsh-stage1-local-packaging-validation.md`
> 里那张"待记录"表继续保留，逐条填，不要因为进了阶段 2 就当它不存在。

### 阶段 2 · 接到产品界面（此时仍与 Chat 并存，默认关闭）

> **开工前提（硬性）**：`DEEPSEEK_API_KEY` 必须就位，阶段 1 那 3 个真实 LLM 用例必须先跑通。
> 原因不是流程洁癖：本阶段第 1、2 条的正确性只能靠**真实模型的 chunk 流**验证。
> mock 出来的 chunk 不会有真实边界情况——UTF-8 字符被切在两个 chunk 中间、
> `agent_thought_chunk` 与 `agent_message_chunk` 交错、最后一个 chunk 与
> `session/turn-ended` 的顺序竞争。拿 mock 写出来的适配器接上真实流很可能当场就错，
> 而且错在"看起来能跑"的地方。

> **1A-2b 未偿项**：本阶段是在 Gate 1A 第 2b 项（保存配置后旧 dsh 进程终止、新进程读到新值）
> 未完成的情况下开工的（2026-09-05 项目决定）。它与本阶段工作正交，但**必须在阶段 3 之前补完**，
> 见 `docs/verification/dsh-stage1a-gate-validation.md`。

#### 施工顺序（六条之间有隐含依赖，不要随意打乱）

**`1 → 2 → 4 → 3 → 5 → 6`**，理由如下：

- **1 必须最先。** 在流式适配器做出来之前，本阶段任何改动在界面上都**看不见**，
  没有肉眼可验证的反馈面。它一落地，聊天面板第一次真正显示 dsh 的输出，
  后面每一条都能当场确认。
- **2 紧跟 1。** cancel 需要一轮真实生成跑到一半才能验，正好复用刚做好的流式链路。
  而且 1、2 是那 3 个真实 key 用例唯一能覆盖到的两条——**趁 key 还在手上一起验掉**，
  不要拖到后面重新找一次凭据。
- **4 排在 3 之前。** 上下文注入方式（走 dsh 的 instructions / system 层，还是拼进 `text`）
  会决定会话怎么建、建在哪一层；反过来会话映射不会影响注入方式。
  顺序颠倒的话，3 做完可能因为 4 的选型而返工。
- **5 排在 4 之后。** 图片是 prompt 内容的一部分，注入方式定了它才有确定的落点。
- **6 是收尾的对照表**，不是一条独立工作。

#### 一条硬性约束：别动 `AiRuntime.ts`

适配器建成 `AiChatPanel` 里挂开关的形式，**不要去改 `AiRuntime.ts` 本体**。
它虽然阶段 3 要删，但在整个阶段 2 期间必须保持可用——它是唯一的回退路径，
也是判断"新链路是不是变差了"的对照组。阶段 2 期间把它改坏，就没有对照组了。

1. **流式适配器**：把 `ai://session-event` 的 `agent_message_chunk` / `agent_thought_chunk` /
   `tool_call` / `tool_call_update` 映射成 `AiChatMessage` 的增量更新，
   按发送时固定的 `sessionId` 写回（多会话工作区那一轮已经确立了这条规则，照它做，不要重新发明）。
   `usage_update` / `config_option_update` 先只记录，不上界面。
2. **cancel**：给 `AcpRuntime` 的 `Command` 加一条 cancel，接 `dsh-acp` 的 `cancelPrompt`，
   前端"停止生成"按钮接过去。这是替换的必要条件，不是可选项（关键事实 7）。
3. **`harnessSessionId` 真正接上**：前端会话 ↔ ACP 会话的映射写进会话元数据，
   并在 Rust 侧持久化（关键事实 9），使应用重启后仍能落到同一个 dsh 会话。
   **若发现 dsh 侧无法按 id 恢复会话，停下来汇报**，不要悄悄退化成"每次重开新会话"。
4. **上下文注入**：现有 `buildAiChatContext` + `PromptRenderer` 产出的项目上下文，
   要在 dsh 路径上以确定的方式进入 prompt。优先考虑 dsh 的 instructions / system 层，
   其次才是拼进 `text`。选哪种要在 `agent-bridge/README.md` 里写明理由。
5. **图片输入**：按 ACP 的 image content block 传。前提是阶段 1 第 5 条的模型配置里选了视觉模型
   （关键事实 5）——这是配置问题，不是"看运气"。
   若在非视觉模型下运行，界面要明确告知"当前模型不支持图片"，并引导去改模型配置；
   **不要**用"把图转成文字描述"之类的替代品糊过去。
6. 此阶段结束时，Cowork 路径应当在功能上覆盖 Chat 路径的：流式、停止、多会话、上下文注入、
   （条件性的）图片输入。**逐项列表对照，缺一项就不算覆盖。**

**Gate：** 一张逐项对照表，Chat 有的能力在 dsh 路径上都有对应实现或明确的"已知缺失"结论。
对照表要能落到 `docs/verification/`，不是只在对话里说过。

### 阶段 3 · 切换与下线

> **拆分（2026-09-05）**：阶段 2 结束时发现**逐 token 流缺失**（关键事实 11），
> 而阶段 3 第 1 步"默认切到 dsh 路径"恰好依赖它。故拆成 3A / 3B：
> **3A 现在就能做，3B 等流式结论。**

#### 3A-0 · 补做逐 token 流（2026-09-05 二次排查后新增，优先做）

> **执行状态：已完成（2026-09-05）**。插件订阅、请求绑定、字节预览、ACP 提交替换、
> 失效硬失败测试和真实模型验收已完成。详见 [3A 验证](../verification/dsh-stage3a-streaming-and-handover.md)。

关键事实 11 已更正：**不必等上游**。插件订阅 `session/event` 的 firehose，
过滤 `assistant/chunk`，经现成 loopback bridge 推回 lamber 做显示。

1. **先验证订阅拿得到**：在 `dsh-tool-lamber` 的 `apply(ctx)` 里挂
   `ctx.on('session/event', ...)`，确认真的收得到 `assistant/chunk` 且带得到
   `turn` / `step` / `chunk`。**拿不到就停下来汇报**，不要改用轮询之类的替代品绕过去。
2. 新增一条 bridge 路由承载增量（参照现有 `CALCULATE_ROUTE` 的写法），
   把 chunk 推回 Rust，再经既有 `ai://session-event` 通道送到前端。
3. 适配器按关键事实 11 的"代价与风险"处理三件事：
   **以 ACP 提交版本替换（不是追加）累积文本**、**按字节缓冲避免 UTF-8 拆字**、
   **把 LLM 层的 `text-delta` / `reasoning-delta` / `tool-call-delta` 映射到现有消息模型**。
4. **写一个订阅失效就大声失败的测试。** 这条是硬要求：
   依赖的是 dsh 内部事件名，升级可能改名，不能让它悄悄退化成"又不流式了但没人发现"。
5. 做完后 3B 的开工条件即满足（流式已补齐），不必再等上游或等试用结论。

**3A-0 的 Gate：** 真实模型下，长回答在生成过程中逐步显示，且最终文本与 ACP 提交版本一致
（不重复、不截断、无乱码）。结论落 `docs/verification/`。

#### 3A · 不依赖流式，现在就做

> **执行状态：已完成（2026-09-05），试用有独立缺陷记录**。接替表和 Computer Use 业务流程
> 见 3A 验证；需求信息页生成正确，生成确认页部分字段回退待修。本轮未进入 3B。

1. **摸清 `AiRuntime` 的全部调用方**，特别是模板页图片分析那条间接依赖
   （`handleSendImageToAi` → `templateAssetSelection` → `AiChatPanel`）。
   产出一张"每条链路由谁接替"的表——这是 3B 删除动作的前置，与流式无关。
2. **带着开关做一轮真实业务流程试用**（测算 → 模板 → 生成文档），开关**手动打开**，
   不动默认值。
   这一步顺带回答一个关键问题：**非流式在真实使用里到底能不能忍。**
   - 能忍 → 不必等上游，直接走 3B 的"接受降级"路径。
   - 不能忍 → 拿到了具体的体感证据，正好补进上游 issue，比空口描述有力。
   - 无论哪种，结论都要落到 `docs/verification/`，写清是谁在什么场景下试的。
3. `AgentLabView` 保留并在文档里说清它是排障入口、不是产品路径。

#### 3B · 默认切换与稳定后下线

> **执行状态（2026-09-05）**：第 4 步默认切换与临时回退代码已完成；第 5 步待真实使用稳定。
> 回退仅作用于新建的独立会话，不持久化全局偏好；旧历史不会冒充 ACP 续聊。
> 自动化、核心聊天 Computer Use 及用户指定模板目录页面核对通过。Gate 1A 空项没有因此关闭。
> 详见 [3B 验证](../verification/dsh-stage3b-default-transition.md)。

4. 把开关默认切到 dsh 路径，保留一次性回退能力。**✅ 2026-09-05 已完成。**
5. 真实使用稳定后，删 `AiRuntime.ts` 及其 endpoint/model/apiKey 相关 UI。
   **▶ 2026-09-06 用户拍板：可以删了。执行清单见下。**

##### 3B-5 删除清单（用户已授权）

前置已满足：3A 的**接替表**在
`docs/verification/dsh-stage3a-streaming-and-handover.md`；默认链路已经过真实业务使用。

1. **按接替表逐条核验**，不是搜一遍 `AiRuntime` 的 import 就完事。
   表里已点名最容易漏的一条：**模板图片间接链路**
   （`TemplateForms.handleSendImageToAi` → `templateAssetSelection` 的 DOM/storage/Tauri 事件
   → `AiChatPanel` 附件 → `loadAiTemplateAsset`）。它不 import `AiRuntime`，却走同一个面板。
2. **同时移除临时回退入口本身。** 这是 3B 记录里写死的一条：
   *"不能留下可点击但无实现的回退选项。"* 含模型设置里"临时回退并新建会话"
   与"返回新版并新建会话"两个按钮，以及回退会话的窗口内存状态。
3. **一并删除只属于旧链路的东西**：`useStreamingParser`（接替表已注明它只属于旧链路、
   不能拿它解析 dsh）、旧 endpoint / model / apiKey 的 UI 与 localStorage 读写
   （`lamber_ai_endpoint` / `lamber_ai_model` / `lamber_ai_api_key`）。
4. **`invokeToolIsolated` 占位一并删**——接替表已确认它没有调用方。
5. **保留 `AgentLabView`**，并在文档里说清它是排障入口、不是产品路径。
6. **删完扫一遍残留引用**，确认应用里不存在指向已删实现的按钮、菜单项或设置项。

**Gate：** 删除的那次提交必须能逐条指出接替表里每条链路由谁接替；
且应用内不存在点了没反应的回退入口。

> ⚠ **删除后就没有回退路径了。** 2026-09-06 刚发生过桥接契约错配导致工具全线 404
> （见 `TASK_BOOK_bridge_contract_handshake.md`）。那次故障与 `AiRuntime` 无关
> ——它挡不住桥接问题——所以不构成推迟删除的理由；但删除后再遇 dsh 侧故障，
> 只能靠回滚版本恢复。发版节奏上要有这个准备。

**3B 的开工条件（满足其一即可）：**

- **3A-0 完成**（自行订阅 `assistant/chunk` 补齐流式）——**这是现在的首选路径**；**或**
- 上游接受并发布了 in-flight delta 投影（issue 草稿见
  `docs/tasks/upstream-issue-acp-incremental-output.md`）；**或**
- 3A 第 2 步的真实试用结论是"非流式可接受"，且该结论已落盘。

**在此之前不要动默认值。** 开关默认关闭、两条链路并存，是当前唯一正确的状态。

**Gate：** 删除 `AiRuntime.ts` 的那次提交，必须能指出替换它的每一条链路各在哪里
（即 3A 第 1 步那张表）。

> **发版阻塞项**（与开发进度分开记，别混在一起）：
> Gate 1B（Windows NSIS 安装验收）与用户真人验收**均未完成**，
> 2026-09-05 决定推迟到有 Windows 机器时再做。
> 它们不阻塞 3A、也不阻塞 3B 的开发，但**阻塞发版**——
> 任何面向真实用户的版本发出去之前必须补完。

---

## 不要做的事

- **不在本阶段扩工具面。** `create_intelligent_compute_project`、`request_project_image`、
  文档生成工具等等，全都等 dsh 融合完成、会话绑定做完之后再说。本次工具集保持
  `run_benefit_calculation` + `write_test_marker` 两个不变。
- **不做会话绑定项目 + 硬性限权**（那是 `TASK_BOOK_cowork_session_project_binding.md`）。
  本次只让 dsh 成为可用的产品链路，不改权限模型。
- **不改审批机制**：`ApprovalGate`、`agent_approval_log`、`AgentApprovalDialog`、`GATED_TOOLS`
  一律不动。它们已通过真人点击四条路径验证，是本次唯一不该碰的成品。
- 不改 `calculator.rs` / `docfill.rs` / 测算引擎 / NPV / 现金流 / 税额 / 甄选费 / 0 容差校验。
- **不改 ACP 协议层的握手与版本断言逻辑**（`EXPECTED_PROTOCOL_VERSION` 那条显式断言是有意为之，
  见 `TASK_BOOK_acp_protocol_rewrite.md`），阶段 0 只是**读出**能力字段，不是去改协商。
- 不做多 dsh 子进程 / 进程池。保持"整个 App 一个常驻子进程、靠 sessionId 区分会话"的架构。
- 不打开 dsh 的遥测（`DSH_TELEMETRY_MODE=DISABLED` 是有意设的，涉及客户财务数据，不要在打包时漏掉这个环境变量）。
- 不为了缩体积去删 dsh 的运行期依赖树里"看起来没用"的包——要瘦身就走 prune / SEA 这类有据可依的方式，
  并且必须在阶段 0 量过、阶段 1 gate 上验证过。
- **不改 dsh 的源码**。换服务商有 baseURL 配置和自写 adapter 两条官方路（关键事实 6），
  两条都够用；改源码会让后续升级跟不上上游。
- **本次不写自有 LLM adapter**。阶段 0 第 2 条只做验证、只记录结论；
  真要写 adapter 是单独一份任务书，不要在融合任务里顺手做。
- 不在阶段 3 之前删除 `AiRuntime.ts`。

---

## 验证要求

1. **阶段 0 的四条结论**全部记录进 `docs/verification/`（新建文件）：
   两个模型下 `promptCapabilities.image` 的实测值、指向 OpenAI 兼容服务的实测结果
   （通过 / 具体哪个字段被拒）、无仓库环境的握手结果、三种分发形态的体积。
   结论不管好看不好看都照实写。
2. **干净机器安装验证**（阶段 1 gate）：无仓库、无 Node、无环境变量的 Windows 机器，
   记录安装包体积、首次启动耗时、配 key 到拿到第一条回复的完整过程。
3. **能力对照表**（阶段 2 gate）：流式增量、停止生成、多会话隔离、会话重启恢复、上下文注入、
   图片输入、模型可配、baseURL 可配，逐项标"已实现 / 已知缺失"，缺失项写明原因。
   其中图片输入要在**视觉模型**配置下实测一次，证明它是配得出来的，而不是只在文档里成立。
4. **审批链路回归**：替换之后再跑一次真人点击四条路径（确认 / 拒绝 / 超时 / 无工作区），
   证明换了触发界面之后审批仍然工作。沿用 `docs/verification/acp-approval-manual-check.md` 的记录格式。
5. **工作区 cwd 验证**：确认 dsh 会话的 cwd 是用户工作区而不是仓库根；工作区未打开时的拒绝路径也要走一次。
6. `cargo test`、`cargo test agent_bridge -- --ignored`（带真实 key）、
   `npm run lint --prefix src-ui`、`npm run build --prefix src-ui`、
   `npm run typecheck --prefix agent-bridge/dsh-tool-lamber` 全过。
7. 完工后更新 `docs/CURRENT_TASK.md`、`docs/CHANGELOG_AI.md`、`agent-bridge/README.md`、
   `docs/modules/release-packaging.md`；照实区分"已验证"与"未验证 / 已知限制"，
   不要把没在干净机器上跑过的路径记成通过。
