# 任务书：Cowork 会话 = 绑定已有项目 + 工具调用硬性限权

> **排序更新**：项目决定把 dsh 协议层从窄协议（`--profile sdk`）整体重写成 ACP（`--profile acp`），任务书见 `docs/TASK_BOOK_acp_protocol_rewrite.md`，**排在这份之前**。ACP 重写会废弃本文写作时假设的 HTTP 审批通道（`/lamber-bridge/approval`），改由 ACP 原生的 `session/requestPermission` 触发；本文"绑定会话 + 硬性限权"具体怎么接入审批链路，要在协议重写落地后重新核对，不要照抄本文现在的描述直接开工。

这是"多会话接 Harness Session"（`CURRENT_TASK.md` 尚未开始的部分第 2 项）的落地第一步，**排在**"新建智算项目"工具（`docs/TASK_BOOK_create_intelligent_compute_project.md`）**之前**。那份任务书里的 create/update/calculate 工具，都要在本次的会话绑定机制做完之后，才能作为"绑定会话里可用的工具"接上去——先后顺序不要颠倒。

产品语义：会话分两种。Chat 模式＝现在的 `AiRuntime.ts` 路径，只读上下文注入，不能操作，本次不碰。Cowork 模式＝dsh 路径，新建时从**已有 lamber 项目列表**里选一个绑定（不是任意文件系统文件夹选择器），绑定之后这个会话里所有工具调用都被硬性限制只能操作这一个项目，不允许模型跑去操作别的 projectId。

## 已确认的关键事实（写代码前不用重新调研）

- dsh 子进程是**整个 App 生命周期内唯一一个、懒加载、常驻**的进程（`AgentRuntime.inner: Mutex<Option<RunningAgent>>`，`agent_bridge/mod.rs`），不是每个会话一个进程。所有 `session_id` 共用同一条 stdin/stdout JSON-RPC 管道并发跑。**这意味着"锁定项目"这件事不能挂在进程启动时的全局配置上**（那样会把整个 App 所有会话都锁死在一个项目上），必须按 `session_id` 精确区分。
- `sessionId` 已经在协议层被使用（`session/prompt` 请求里带 `sessionId`），但目前**同一个 sessionId 复用会报 "id collision"**（`agent-bridge/README.md:243`），现在唯一的调用方 `AgentLabView.tsx` 靠每次生成 `lab-${Date.now()}` 绕开这个问题——等于现在每一轮对话事实上都是新会话，没有连续性。这个问题不解决，真正的多轮 Cowork 会话做不起来，必须在本次一并查清楚：collision 到底是"同一 sessionId 并发调用两次"触发的，还是"任何形式的复用"都不行。如果是后者，现有的会话/多轮对话设计假设是错的，要先停下来汇报，不要自己拍板改。
- 桥接服务（`bridge_server.rs` + `workspace_handler`）目前收到的 HTTP 请求里**没有 sessionId**，只有工具自己往 `postBridge` 里塞的业务参数（比如 `projectId`）。`run_benefit_calculation` 现在完全信任模型传来的 `projectId`，没有任何校验（`calculation.rs:82-84` 直接拿去查库）。
- 桥接的 handler 闭包（`workspace_handler`）是在 `launch()` 时构建一次、之后一直复用的（因为进程本身常驻），`runtime`/`gate`/`announce` 都是在这个闭包里捕获的现成模式——技术上很适合在旁边加一个"session → 锁定的 projectId"的登记表，在校验时查表比对，不用改动这套闭包捕获的基本结构。

## 要做的事

1. **先验证一个前提，不确认清楚不要往下搭**：dsh 的工具 `execute(args, exec)` 里的 `exec` 上下文，是否暴露一个**不受模型控制、由 dsh 运行时自己提供**的当前 sessionId（类似现有工具已经在用的 `exec.signal`）。这决定了硬性限权能不能做成"运行时可信来源"而不是"模型自己说它在哪个会话，你就信了"。如果 `exec` 里没有这个东西，说明现在的工具执行上下文设计压根不知道自己属于哪个会话，这个硬性限权方案要推翻重新想——查不到就停下来汇报，不要绕过去假装能做。
2. **Rust 侧新增 session → project 绑定登记表**：一个 `Mutex<HashMap<String, String>>`（sessionId → 锁定的 projectId），挂在 `AgentRuntime` 或同级的全局状态上；新增 Tauri 命令（例如 `ai_bind_session_to_project(session_id, project_id)`），在用户为一个 Cowork 会话选定项目之后、发第一条消息之前调用，写入这张表。
3. **改造现有的 postBridge 调用链路**：工具的 `execute()` 要把第 1 条里确认到的可信 sessionId 一并发给桥接服务（不能只发模型自己填的 `projectId`）。桥接路由处理业务逻辑之前，先用 sessionId 查第 2 条的登记表拿到锁定的 projectId，跟请求里实际要操作的 projectId 做比对，不一致直接拒绝（fail closed，参照现有令牌鉴权、超时自动拒绝的一贯风格，不要设计成"默认放行、异常才拒绝"）。
4. **先在 `run_benefit_calculation` 这条现成链路上把校验跑通**：给它加上第 3 条的校验逻辑，虽然它现在没有任何写操作，但可以拿它先验证"会话绑定项目 A、请求项目 A"（放行）和"会话绑定项目 A、请求项目 B"（拒绝）这两条路径，作为整套机制的第一个真实用例。
5. **前端会话创建流程（Cowork 模式）**：新建 Cowork 会话时，从**已有 lamber 项目列表**（`Project` 表，已经有 `folder_path`）里选一个，不做任意文件系统文件夹浏览。选完调用第 2 条的绑定命令，再允许开始对话。Chat 模式的现有创建流程原样不动。
6. **验证要求**：带真实 `DEEPSEEK_API_KEY`，覆盖"合法项目放行"和"跨项目请求被拒绝"两条真实链路（不能只测校验函数本身）；结果记录进 `docs/verification/`。

## 不要做的事

- 不做"新建智算项目"这个写工具（`create_project_in_workspace` 包装）——那是绑定机制做完之后，绑定会话里能用的下一个工具，这次不碰，见 `TASK_BOOK_create_intelligent_compute_project.md`。
- 不碰 Chat 模式 / `AiRuntime.ts` 路径，维持现状。
- 不做任意文件系统文件夹浏览选择器（`select_local_folder`）——本次绑定范围只在"已有 lamber 项目列表"里选，不是文件系统层面的任意文件夹，这跟之前讨论过的 Claude Projects 截图那种通用文件夹选择器是两回事。
- 不改审批弹窗、不给这次的校验逻辑加人工审批——"跨项目请求被拒绝"是硬性权限校验，不是需要用户点确认/拒绝的审批场景，两者不要混着做。
- 不做多 dsh 子进程/进程池改造——保持现有"整个 App 一个常驻 dsh 子进程，靠 sessionId 区分会话"的架构，除非第2条那个 collision 前提查出来必须改，那种情况先停下来汇报，不要自己拍板做这么大的架构变更。
- 不做参数二次确认 UI、不做多个 Cowork 会话之间的数据隔离展示优化——这次只做"绑定 + 硬性限权"这一件事，别的顺手的事不要做。
