# 任务书：dsh 协议层从 `--profile sdk` 整体重写为 `--profile acp`

前置验证已通过（`docs/TASK_BOOK_acp_protocol_spike.md`，见 `docs/verification/` 下的探针报告）：Rust 官方 crate `agent-client-protocol` 能和 `dsh --profile acp` 真实握手、走完 `initialize → session/new → session/prompt → 真实模型响应` 全流程，`session/requestPermission` 的 handler 注册 API 也确认存在。这次是把探针验证过的东西，按生产代码标准正式接入，不是重新验证一遍。

这份任务书排在 `docs/TASK_BOOK_cowork_session_project_binding.md`、`docs/TASK_BOOK_create_intelligent_compute_project.md` **之前**——那两份都建立在"审批通道长什么样、会话怎么绑"这些假设上，这次协议重写会改变这些假设的具体实现方式，必须先落地。

## 已确认的关键事实（写代码前不用重新调研）

- **启动方式几乎不变**：同一个 `dsh` 二进制，`--profile acp` 换掉 `--profile sdk`。`dsh_session.rs::spawn()` 的进程启动、环境变量注入这套逻辑基本保留，只是要发的请求/能接的请求变了。
- **ACP 是双向协议，现有 Rust 客户端模型撑不住**：`dsh_session.rs` 现在是纯"发请求等回应"模型（`next_id: AtomicI64` + `pending: HashMap`），收到服务端主动发起的请求（`session/requestPermission`）现在直接当日志丢掉、不回应。这次要整个换成 `agent-client-protocol` crate 的 client builder API（`on_receive_request`/`on_receive_notification` 那套），不是在现有结构上打补丁。
- **`dsh-acp` 自带一套 `approval/request` 处理器，且在 Cordis 事件链里排在 `dsh-tool-lamber` 自己的 `approval.ts` 前面，会抢答**——这是 ACP 下最大的结构性变化。已验证/已上线的那套 HTTP 审批通道（`askLamber()` → `POST /lamber-bridge/approval` → Rust `mod.rs` 里的 `APPROVAL_ROUTE` 分支）在 ACP 下会变成**永远走不到的死代码**，因为 `dsh-acp` 的 handler 会先接住这个 Cordis 事件、直接转发给 ACP 的 `session/requestPermission`。这不是要不要保留的选择，是结构上被短路了。
  - 好消息：`approval.rs` 里 `ApprovalGate`（park-on-Condvar、`handle_request()`/`resolve()`）、`approval_log.rs`（`agent_approval_log` 落库）、Tauri 事件（`ai://approval-request`）、`ai_resolve_approval` 命令，这些核心机制跟触发方式无关，可以原样复用，只换触发入口。
  - `tools/pre-execute` 那半部分（`GATED_TOOLS`/`isGatedTool` 判断哪个工具需要审批）不受影响，保留不动。
- **协议版本探针已验证兼容，但发现一个必须补的硬约束**：`dsh-acp` 的 `initialize()`（`dsh-acp/lib/index.js:1143`）**不校验**客户端传来的协议版本，无条件回 `1`。这次探针刚好双方都是 v1 所以侥幸吻合，但以后任何一边升级到 v2，握手阶段不会报错，问题会推迟到某个字段解析时才炸。**正式实现必须在客户端侧对 `initialize` 返回的 `protocolVersion` 加显式断言（不等于预期值就直接失败退出），不能假设它以后还会一直是 1。**
- **依赖版本以探针实测为准，不要照抄早期调研的假设**：`agent-client-protocol` crate 2.0.0 实际锁定的是 `agent-client-protocol-schema = "=1.5.0"`，不是最初调研时看到的 1.7.0。接入时用 `cargo tree` 现查一遍实际解析到的版本，写代码前确认一次。
- **`provision-profile.mjs` 会自动把 `dsh-tool-lamber` 链进任意 profile**，包括新建的 `acp` profile——探针阶段因为 `acp` profile 的 `bundles` 没包含它、且没传 `--patch`，侥幸没被挂载，结果不受影响。**这次是正式接入，`dsh-tool-lamber` 必须被有意识地链进 `acp` profile**（不然工具都调不了），链上之后要确认 `run_benefit_calculation`（不需要审批）和会话绑定/限权（如果那时候已经做完）这些逻辑在新协议下依然按预期工作。
- **ACP 支持图片内容块，但这次不启用**：`ImageContent` 类型确认存在（base64+mimeType，PNG/JPEG/WebP/GIF），但要用还需要挂 `dsh-attachment`/`dsh-attachment-local` 插件、且模型路由要声明支持图片输入——这些留到以后单独做，这次协议重写顺带具备了这个能力，但不主动接。
- **`session/resume` 不等于 `session/load`**：`dsh-acp` 只实现了前者，且明确不重放历史消息。跨 App 重启后 UI 要显示的历史对话，还是要靠 lamber 自己本地的会话记录，不能指望 ACP 帮忙补。这个不用额外做什么，只是别假设 resume 会自动把消息推回来。

## 要做的事

1. **Rust 传输层重写**：`agent-client-protocol` 从探针阶段的 `[dev-dependencies]` 转为正式的 `[dependencies]`（连同 `tokio`）。重写 `dsh_session.rs` 的核心模型：用 crate 的 client builder 完成握手（带上面提到的协议版本显式断言）、`session/new`、`session/prompt`，并且注册 `on_receive_request` 处理 `session/requestPermission`、注册通知处理器解析 `session/update`（强类型的 `SessionUpdate` 枚举：`agent_message_chunk`/`tool_call`/`tool_call_update`/`usage_update` 等）。探针脚本 `acp_handshake_probe.rs` 达到目的后可以删除或归档，不要把验证代码直接原地转正当产品代码——正式实现要有自己的模块划分和错误处理。
2. **`acp` profile 正式接入**：把 `dsh-tool-lamber` 有意识地链进 `.dsh-home/profiles/acp/`，跑一遍确认工具能被模型正常调用。
3. **新写权限应答 handler**：Rust 端实现 `session/requestPermission` 的处理逻辑，复用 `approval.rs` 现有的 `ApprovalGate`/审计落库/Tauri 事件机制，只是触发入口从 HTTP route 换成这个新 handler。
4. **退休旧的 HTTP 审批通道**：删除 `dsh-tool-lamber/src/approval.ts` 里 `askLamber()` 的应答半部分、`pendingCalls.ts`、Rust 侧 `/lamber-bridge/approval` 路由（`mod.rs` 里的 `APPROVAL_ROUTE` 分支）、`agent-bridge/scripts/check-approval.mjs`。不要留双通道当备用方案。`tools/pre-execute` 那半部分保留。
5. **`AgentRuntime`/`mod.rs` 收口**：`send_prompt`、`describe()`、session 生命周期，从 sdk 协议模型迁移到 acp 协议模型；处理好 tokio 异步和现有同步后端之间的桥接，不要把整个 Rust 后端顺手改成到处 tokio——只在跟 dsh 通信这一层引入异步，别的地方保持原样。
6. **审批链路重新验证**：四条路径（确认/拒绝/超时/无工作区）在新的 `session/requestPermission` 触发入口下**重新走一遍真人点击验证**，不能援引旧的 SDK-profile 时代的验证记录当结论——那是在完全不同的触发机制下测的。结果记录进 `docs/verification/` 下的新文件，旧文件（`approval-channel-manual-check.md`）保留不动，作为历史记录。
7. **验证要求**：`cargo test` 全过；带真实 `DEEPSEEK_API_KEY` 的集成测试覆盖 `initialize`（含版本断言）→ `session/new` → `session/prompt` → 真实模型响应 → 至少一次真实的 gated 工具触发 `session/requestPermission` 全链路；四条审批路径的真人点击记录写进新的验证文档。
8. 完工后更新 `docs/CURRENT_TASK.md`：记录本次改动、废弃了什么（旧 HTTP 审批通道、`sdk` profile 的生产代码引用）、新增了什么、还剩什么未做。

## 不要做的事

- 不要保留 `sdk` profile 相关代码路径作为"双通道"或回退方案——决定是彻底切换到 ACP。`.dsh-home/profiles/sdk/` 目录本身可以先不删，但生产代码不应该再引用它。
- 不要在这次顺带做 Cowork 会话绑定项目 + 硬性限权（`TASK_BOOK_cowork_session_project_binding.md`），排在这次之后，且那份文档里"怎么接审批"的具体描述要在这次落地后重新核对。
- 不要在这次顺带做智算项目创建工具（`TASK_BOOK_create_intelligent_compute_project.md`），排得更后。
- 不要现在就启用图片输入（挂 attachment 插件、接 `fs/*` 相关能力）——协议层面具备了这个能力就够了，真正启用留到以后单独做。
- 不要实现 `terminal/*`、`elicitation/*`、`fs/readTextFile`/`fs/writeTextFile` 的 client 端处理——`dsh-acp` 从不调用这些，属于死代码，不要防御性地写。
- 不要吞掉/淡化探针阶段发现的两个真实风险（协议版本不校验、schema 版本跟早期调研预期不符）——必须在代码里落实成显式的版本断言和现查依赖版本，不能只是在任务书里提一句就算完成。
- 不要在没有真人点击验证四条审批路径之前，就宣称审批机制在 ACP 下"迁移完成"——参照探针阶段"跑通才算数"的标准，编译通过、类型对得上不算验证。

## 需要提前提醒的事

用于验证的 `DEEPSEEK_API_KEY` 在本次对话记录里已经多次以明文形式出现，建议尽快在 DeepSeek 后台吊销并重新生成一个新的，正式实现阶段的密钥管理要避免再次让密钥出现在任何会被记录/提交的地方（探针阶段的处理方式——只作为子进程环境变量传递、不落任何文件、报告里用占位符——是正确的做法，继续照这个标准执行）。
