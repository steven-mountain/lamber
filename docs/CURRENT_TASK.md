# 采购甄选费 · 可选投入科目写入

- **Status:** Done
- **Objective:** 采购甄选测算不再固定写入集成服务，允许用户从投入科目中选择目标，并按供应商承担甄选服务费的业务口径写入。

## Progress

1. [x] 新增目标投入科目选择器，覆盖 IT/移动云、CT、非IT/CT和综合类投入；中标服务费作为系统自动写入科目不参与选择。
2. [x] 写入口径统一为“目标科目 = 最高限价 - 甄选服务费（供应商报价 + 上浮）”“中标服务费 = 甄选服务费”，两项合计等于最高限价。
3. [x] 两个科目改用统一批量更新入口，避免同一投入分组连续更新互相覆盖，并同步科目付款计划和既有税额校验。
4. [x] 目标科目代码随方案工作副本和快照保存；旧方案、Excel 导入和无效代码默认回退集成服务。
5. [x] 智能反算仅在反算当前甄选目标科目时刷新报价、服务费和最高限价；签批表 A 表使用保存的目标科目名称与税率。
6. [x] 清理前端既有 lint 债务：修复 AI 上下文防抖实现，拆分组件文件中的非组件导出，并稳定项目加载、模板自动保存及 ICT 计算相关 Hook 生命周期。

## Validation

- `npm run test:selection-batch --prefix src-ui`：通过。
- `npm run build --prefix src-ui`：通过。
- `cargo test`：36 项全部通过（仅仓库既有 warnings）。
- `npm run lint --prefix src-ui`：通过，0 错误、0 警告。
- `npm run test:ai-compute-quote --prefix src-ui`、`npm run test:tax-split --prefix src-ui`：通过。

## Scope Boundary

- 未修改 Rust 甄选费阶梯公式、NPV、现金流公式、税额计算或 0 容差门槛。
- Excel 模板仍只在 T26/T29 保存报价与上浮，因无目标科目坐标，导入时按兼容规则默认集成服务。


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
# AI 多 Session 会话工作区（前端阶段已完成）

- **Status:** 已完成前端多 Session UI、消息隔离、项目归属元数据、localStorage 持久化和窄窗抽屉。
- **模块文档:** [docs/modules/ai-session-workspace.md](./modules/ai-session-workspace.md)

## 已完成

- `AiChatPanel` 的单一本地 `messages` 状态迁移到 `useAiSessionStore`，每个 Session 独立保存 `AiChatMessage[]`。
- 新增 Session Sidebar、当前项目 / 其他会话分组、最近更新时间排序和当前会话恢复。
- Session 操作菜单已支持列表内重命名和确认后删除；删除当前会话自动选择最近会话，删除最后一个会话自动补建空白会话，删除生成中的会话会先停止流式请求。
- 流式输出按发送时固定的 `sessionId` 写回，生成期间切换会话不会串消息；现有 AiRuntime、PromptRenderer、SSE、Abort、图片输入和项目上下文链路保持不变。
- 默认 AI 窗口宽度调整为 `780px`；小于 `680px` 时侧栏变为覆盖式抽屉，输入区不被压缩。
- localStorage 保存版本化 Session 快照；图片附件持久化时去除大体积 base64、保留元数据。

## Validation

- `npm run build --prefix src-ui`：通过。
- 应用内浏览器实测：创建 3 个会话、消息隔离、切换恢复、刷新恢复均通过。
- 应用内浏览器实测：会话重命名即时更新并持久化，确认删除后会话数量正确减少。
- 本地 mock SSE 实测：生成期间切换会话只更新发起会话；停止生成后流立即中断且输入框恢复可用。
- `900px` 双栏与 `420px` 抽屉两种布局均完成截图检查。
- `npm run lint --prefix src-ui`：本次新增/修改文件无新错误；全仓仍被既有 `useAiContextStore.ts` 的 `@typescript-eslint/no-this-alias` 阻断。

## Scope Boundary

- 未修改 `AiRuntime.ts`、PromptRenderer、Rust、数据库 schema、效益测算、文档生成或项目管理逻辑。
- 未接入 deepseek-harness，未新增 dsh、JSON-RPC、Agent Tool、Approval、Sub-Agent 或服务端 Session。
- 按当前范围未新增项目创建入口、会话项目/通用归属移动，也未调整 AI 窗口背景配色。

---

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
# ACP 协议层重写（`--profile sdk` → `--profile acp`）

- **Status:** ✅ 已完成。代码、带真实 key 的集成测试、四条审批路径的真人点击验证全部通过。
- **真人点击验证记录:** [docs/verification/acp-approval-manual-check.md](./verification/acp-approval-manual-check.md)
- **任务书:** [docs/TASK_BOOK_acp_protocol_rewrite.md](./TASK_BOOK_acp_protocol_rewrite.md)
- **前置验证:** [docs/verification/acp-rust-crate-handshake.md](./verification/acp-rust-crate-handshake.md)
- **模块文档:** [agent-bridge/README.md](../agent-bridge/README.md)

## 为什么要重写

ACP 是**双向**协议：agent 也会向客户端发请求。原来那个手写的 JSON-RPC 客户端是纯
「发请求等回应」模型（`next_id` + `pending` 表），收到服务端主动发起的请求只会当日志丢掉。
`session/requestPermission` 恰恰就是这样一条请求，所以传输层是整体换掉，不是打补丁。

## 改了什么

### 传输层

- `agent-client-protocol 2.0.0` + `tokio` 从 `[dev-dependencies]` 转正为 `[dependencies]`。
  实测解析到的 schema 是 **1.5.0**（crate 把它精确锁死在 `=1.5.0`），不是早期调研说的 1.7.0。
- `dsh_session.rs` 重写：`DshSession`（手写 JSON-RPC）→ `AcpRuntime`（crate 的 client builder）。
  握手、`session/new`、`session/prompt`，外加 `on_receive_request` 处理
  `session/requestPermission`、`on_receive_notification` 强类型解析 `session/update`。
- **协议版本显式断言**：dsh 的 `initialize` 不校验入参、无条件回自己的版本
  （`dsh-acp/lib/index.js:1143-1146`），所以它升到 v2 的那天握手期不会报错。
  客户端侧现在协商结果不等于 `EXPECTED_PROTOCOL_VERSION` 就直接启动失败。
- tokio **只**出现在这一层：连接跑在自己的线程和 runtime 上，同步侧通过命令通道 +
  `std::sync::mpsc` 回执与它交互，后端其余部分未引入异步。
- `session/prompt` 做成**投递即返回**（ACP 那条请求要整轮结束才回），轮次结束另发
  `session/turn-ended` 事件。ACP 会话 id 由 dsh 生成，`AgentRuntime` 负责把前端自己的
  会话名映射到它。
- 探针 `src-tauri/examples/acp_handshake_probe.rs` 已删除——验证代码不原地转正。

### 审批链路

- 触发入口从 `POST /lamber-bridge/approval` 换成 ACP 的 `session/requestPermission`。
  `ApprovalGate`、`agent_approval_log` 落库、`ai://approval-request` 事件、
  `ai_resolve_approval` 命令这些核心机制**原样复用**。
- 新增 `tool_calls.rs`：`session/requestPermission` 只带 `toolCallId`，工具名与参数来自
  更早那条 `tool_call` 通知，这个索引负责按 id 关联。它是插件里 `pendingCalls.ts` 的继任者。
- 新增 `approval.rs` 里的展示文案镜像表：`dsh-acp` 把守卫写的 `reason` 丢掉了，弹窗文案
  只能存在 Rust 侧。`gated_tool_names_match_the_plugin` 读插件源码守住两边不漂移。
- **顺带修掉一个真实的时序缺口**：`handle_request` 原先先公告、后登记槽位，中间有个窗口，
  在此期间到达的答复会被判成「请求不存在」，用户的点击会被丢掉、问题继续挂到超时。改成
  登记后在锁内公告，`an_answer_racing_the_announcement_is_not_lost` 守这条。

### 删除的东西（不留双通道）

- `dsh-tool-lamber/src/approval.ts` 里的 `approval/request` 答复器与 `askLamber()`
- `dsh-tool-lamber/src/pendingCalls.ts`
- Rust `mod.rs` 的 `APPROVAL_ROUTE` 分支（桥接现在只剩一条只读路由）
- `agent-bridge/scripts/check-approval.mjs`
- `examples/acp_handshake_probe.rs`
- 依赖 `@deepseek-ai/dsh-user-approval`（插件已不再引用其类型）

`tools/pre-execute` 守卫（`GATED_TOOLS` / `isGatedTool`）按任务书要求**保留未动**。

### profile

- `dsh-tool-lamber` 已有意识地链进 `.dsh-home/profiles/acp/`，`patch.yml` 与
  `provision-profile.mjs` 的默认 profile 都改成 `acp`。
- `.dsh-home/profiles/sdk/` 目录保留未删，但生产代码已不再引用它。

## Validation

- `cargo test`：**57 passed, 0 failed**。
- `cargo test agent_bridge -- --ignored`（**带真实 `DEEPSEEK_API_KEY`**）：**6 passed, 0 failed**。
  覆盖 `initialize`（含版本断言）→ `session/new` → `session/prompt` → 真实模型响应 →
  真实 gated 工具触发 `session/requestPermission` 全链路。
- `npm run build --prefix src-ui`：通过。
- `npm run typecheck --prefix agent-bridge/dsh-tool-lamber`：通过。

### 集成测试的旁证（不只看「测试变绿」）

真实模型跑完后，系统临时目录下每次运行**恰好**多出一个标记文件，内容是
`备注: ACP 联调`——正是测试提示词里要求模型填的 `note` 参数。这同时证明了三件事：

1. 模型真的解析了指令并发出 `write_test_marker` 的工具调用（不是桩）；
2. 审批真的经由 ACP 的 `session/requestPermission` 走通，确认后工具才执行；
3. 用例内部循环了「确认」与「拒绝」两种立场，而每次运行只产出**一个**文件——
   拒绝那一轮确实没有执行，不是"执行了但断言没看见"。

`dsh_tool_call_reaches_the_calculator_and_returns_real_numbers` 另外断言了桥接**恰好**
被打一次、`projectId` 正确、工具结果里出现 `calculator.rs` 算出的真实 NPV，且只读工具
全程**没有**触发任何审批弹窗。

### 真人点击验证（四条路径全通过）

2026-09-04，经应用内 `#/agent-lab` 联调台操作，触发入口为 ACP 的
`session/requestPermission`。完整记录与原始报文见
[acp-approval-manual-check.md](./verification/acp-approval-manual-check.md)。

| 路径 | 结果 | 证据 |
| --- | --- | --- |
| 点「确认执行」 | ✅ `decided_by=user`、已批准 | 审计表 + 标记文件（决定后 10ms 写出，顺序正确） |
| 点「拒绝」 | ✅ `decided_by=user`、已拒绝 | 审计表；无新标记文件 |
| 不操作等 90 秒 | ✅ `decided_by=timeout` | 审计表；无新标记文件；随后可继续正常发指令 |
| 无工作区时审批 | ✅ 决定照常生效，缓冲后回填 | `agent-approval-spool.jsonl` → 打开工作区后入库、缓冲文件被删 |

事件流同时证实了本次改动最关键的结构性判断：弹窗里的工具名与参数确实取自更早那条
`tool_call` 通知的 `title` / `rawInput`——权限请求本身只带 `toolCallId`。

旧的 [approval-channel-manual-check.md](./verification/approval-channel-manual-check.md)
是 SDK 协议时期、完全不同触发机制下测的，**未被援引为本次的结论**；该文件按要求保留未动。

## 本次未覆盖的部分（照实记录）

1. 审批只覆盖了 `write_test_marker` 一个工具——`GATED_TOOLS` 目前也只有它。
2. dsh 自带工具（bash、文件编辑等）触发权限请求时的表现未验证。ACP 下 lamber 是唯一的
   权限应答方，那条路径存在但没走过。
3. 并发审批（同时挂起两个问题）未验证。

---

# agent-bridge：deepseek-harness 接入（闭环 A + 闭环 B 已完成）

- **Status:** 闭环 A、闭环 B 均已完成，并通过真实点击验证。
- **分支:** `feat/agent-bridge-approval-loop-b`（基于 `feat/agent-bridge-dsh-loop-a`）
- **模块文档:** [agent-bridge/README.md](../agent-bridge/README.md)
- **验证记录:** [docs/verification/approval-channel-manual-check.md](./verification/approval-channel-manual-check.md)

## 背景与架构原则

lamber 原有的 AI copilot 只是一个裸的 OpenAI 兼容流式客户端，`invokeToolIsolated()` 是空壳，没有真正的工具执行能力。本次用 [deepseek-harness](https://github.com/deepseek-ai/deepseek-harness)（`dsh`）补上，而不是自己再造一遍 agent loop。

**原则：lamber 保留全部 Rust 业务逻辑（`calculator.rs`、`docfill.rs` 等）不变；dsh 只负责 agent loop / 工具编排 / 审批，通过本地回环 HTTP 桥接回调 Rust。**

```
React 前端  →  Tauri 命令  →  Rust 后端
                                 │ spawn 子进程 + newline-delimited JSON-RPC 2.0 over stdio
                              dsh 子进程（--profile sdk；已由 ACP 重写取代，见本文首节）
                                 │ 自定义插件工具 execute() 内 fetch
                              127.0.0.1 桥接服务（tiny_http，令牌鉴权）
                                 └─ benefit::calculator / 审批网关
```

## 闭环 A：Agent 工具执行 ✅

`React → Rust → dsh → run_benefit_calculation → Rust calculator → dsh → UI`。

- `agent-bridge/dsh-tool-lamber/`：独立 npm 包形态的 dsh 工具插件。工具体只把参数 POST 给桥接服务，不含任何业务数学。
- `src-tauri/src/agent_bridge/`：桥接服务（`tiny_http`，仅监听 127.0.0.1，一次性随机令牌鉴权）、dsh 子进程管理、手写 JSON-RPC 2.0 客户端。
- 验证：模型真实发出 `tool/call`，桥接被打一次且 `projectId` 正确，`tool/result` 带回 `calculator.rs` 算出的真实 NPV。

## 闭环 B：人工审批通道 ✅

`dsh → answerer 插件 → Rust 桥接 → Tauri emit → React 弹窗 → 用户确认 → 原路返回`。

dsh 的审批是进程内 Cordis 事件（`approval/request`，waterfall），**不会**经 SDK 的 JSON-RPC 自动转发给外部客户端，因此在 dsh 侧写了 answerer 插件做转发。

- 无害测试工具 `write_test_marker`（只写系统临时目录的标记文件），挂审批钩子。
- 超时分层、失败关闭（`unavailable` 一律判拒）、令牌鉴权、`Debug` 脱敏、挂起槽位清理均已实现并验证。
- 审批审计落 `agent_approval_log` 表（schema v9 → v10）；无工作区时先进 `agent-approval-spool.jsonl` 缓冲，工作区打开时单事务回填。

### 真实点击验证（四条路径全通过）

| 路径 | 结果 | 证据 |
| --- | --- | --- |
| 弹窗渲染 | ✅ 工具名 / 参数 JSON / 倒计时 / 双按钮 / 未被遮罩挡住 | 窗口截图 |
| 点「确认执行」 | ✅ `approved=1, decided_by=user` | 审计表 + 标记文件（决定后 14ms 写出，顺序正确） |
| 点「拒绝」 | ✅ `approved=0, decided_by=user` | 审计表；无新标记文件 |
| 不操作等 90 秒 | ✅ `approved=0, decided_by=timeout` | 审计表；无新标记文件 |

**限制（照实记录）：** 确认/拒绝两次点击由**人工完成**，不是自动化验证。本机对 `osascript` 的辅助功能授权始终未生效（-1719/-25211）。超时一路无需点击，是完全自动的。

## Validation

- `cargo test`：53 passed，无回归。
- `cargo test agent_bridge -- --ignored`（带真实 `DEEPSEEK_API_KEY`）：9 passed。
- `npm run build --prefix src-ui`：通过。

## 尚未开始的部分

1. **新建项目工具 `create_project`** —— 第一个真正写 lamber 业务数据的工具。审批通道已验证可用，是它的前置条件；现在可以做了。需要考虑：参数校验、与现有 `create_project_in_workspace` 命令的关系、审批文案要能让用户看清将写入什么。
2. **多会话对应 Harness Session** —— `AiChatPanel` 的前端多 Session 工作区已经完成，但 `harnessSessionId` 仍只是预留字段；尚未把终端用户会话接入 dsh / Harness 执行链。
3. **参数抽取二次确认** —— 模型从自然语言里抽出的参数（金额、项目名、年限等）在执行前让用户核对修正。当前审批弹窗只做"批准 / 拒绝"，不能改参数。

## Scope Boundary

- 未改动 `calculator.rs` / `docfill.rs` / 测算引擎 / NPV / 现金流 / 税额 / 甄选费 / 反算 / 0 容差校验。
- 未改动 `AiRuntime.ts`；`AiChatPanel.tsx` 仅增加前端 Session 容器与响应式布局，仍使用既有 OpenAI-compatible SSE 调用链。
- 插件包内除 `run_benefit_calculation` 与无害的 `write_test_marker` 外无其它工具；**没有**任何会写 lamber 业务数据的工具。
- 未做 SEA 单文件打包 / 瘦身。
