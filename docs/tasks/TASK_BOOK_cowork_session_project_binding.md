# 任务书：会话绑定项目 + 工具调用硬性限权（2026-09-06 重写版）

> **本文是重写版。** 初版（2026-09-04 前后）的产品语义前提是"Chat 模式与 Cowork 模式并存、
> Chat 不碰"，该前提**已彻底作废**——dsh 现在是默认且唯一的 AI 链路，旧 `AiRuntime` 只剩
> 一个临时回退入口，且计划在 dsh 融合任务书阶段 3B 第 5 步删除。
>
> 初版里三条"待验证的前提"现在**已经有答案了**（见关键事实 2、3），不要再按初版重新调研一遍。
> 初版内容可从 git 历史取回，本文不保留。
>
> **位置**：路线图 [ROADMAP_ai_capabilities.md](./ROADMAP_ai_capabilities.md) 第 ② 项。
> 前置 ① 已完成。本文是 ③④⑥ 所有工具的地基。

## 目标

给每个 AI 会话绑定一个 lamber 项目，绑定后该会话里的所有工具调用**被硬性限制只能操作这一个
项目**，模型不能跑去操作别的 `projectId`。

## 已确认的关键事实（写代码前不用重新调研）

### 1. 现在没有"两种会话"了

dsh 是默认链路（阶段 3B 已切换）。所以本文不再区分 Chat / Cowork：
**绑定机制适用于所有会话**。旧 `AiRuntime` 回退入口不在本文范围，也不要给它加绑定——
它即将被删除。

### 2. 可信 sessionId 已确认可得（初版的阻塞问题，现已解决）

初版第 1 条要求先验证"工具执行上下文能否拿到不受模型控制的 sessionId，查不到就停下来汇报"。
**答案是能，有两条独立途径，都已在代码或类型声明中确认：**

- **工具侧**：`ToolExecutionInput.agent?: Agent`
  （`@deepseek-ai/dsh-tools/lib/types/index.d.ts:207-208`），注释原文
  *"The agent on whose behalf the call runs (**set by the agent loop**)"*——
  由运行时设置，不是模型参数。会话 id 取 `exec.agent?.session.id`。
- **轮次侧**：`ctx.on('agent/pre-step', ({ agent, turn, signal }, next) => ...)`，
  取 `agent.session.id`。**这条已经在生产代码里跑着**——
  `dsh-tool-lamber/src/stream.ts:26-30` 用它在模型执行前完成绑定。

⚠ `agent` 是**可选字段**。取不到时必须 **fail closed**（拒绝调用），
不得退化成"没有 agent 就放行"。

### 3. Rust 侧已有同形状的登记表可作范本（不要另起炉灶）

阶段 3A 的流式工作已经建好 `src-tauri/src/agent_bridge/turns.rs`：

```rust
pub struct Turns(Mutex<HashMap<String, (String, Option<String>)>>);
//                          ACP session id -> (前端 sessionId, requestId)
```

它已经具备本文需要的几乎全部性质：按 ACP 会话 id 索引、**在模型执行前**登记
（`begin()`）、以及严格的"迟到事件绝不重新贴标签"校验（`emit_stream` 里三重比对后才放行）。
`begin()` 还会拒绝同一会话的并发轮次。

**新增的项目绑定应当照这个模式写，并放在它旁边**，但注意一个关键差异：

| | `Turns` | 项目绑定 |
| --- | --- | --- |
| 生命周期 | **每轮**（begin / end） | **每会话**（跨多轮存活） |
| 失效时机 | 轮次结束 | 会话删除或用户改绑 |

不要把项目绑定塞进 `Turns` 复用——生命周期不同，会互相破坏。

### 4. 会话身份已持久化，绑定也必须持久化

阶段 3A/3B：会话映射存在本机独立的 `ai-sessions.sqlite`，重启后经 `session/resume`
恢复原 ACP 会话 id。

→ **项目绑定如果只放内存，应用一重启，被恢复的会话就没有锁了**，而用户以为它还锁着。
这是初版写作时不存在的新约束（当时映射还只在内存里）。绑定必须与会话映射一起持久化。

### 5. 现有工具完全信任模型传来的 projectId

`runBenefitCalculation.ts` 的 `execute` 直接把 `args.projectId` POST 给桥接
（`postBridge(CALCULATE_ROUTE, { projectId: args.projectId, ... })`），
Rust 侧拿到就查库，没有任何归属校验。**这就是本文要堵的洞**，也是第一个改造对象。

### 6. cwd 已经是用户工作区

阶段 1 已把 `session/new` 的 cwd 从仓库根改成当前打开的 workspace，
未打开工作区时在启动子进程前明确拒绝。初版写作时还是仓库根，相关描述已过时。

### 7. "id collision" 是过时信息

初版要求排查"复用同一 sessionId 会报 id collision"。那是 SDK 协议时期的坑，
ACP 下会话 id 由 dsh 生成，**该坑不存在**（`agent-bridge/README.md` 已更正）。不要再查。

## 要做的事

1. **新增 session → project 绑定登记表**（照关键事实 3 的模式，注意生命周期差异），
   并与会话映射一同持久化到 `ai-sessions.sqlite`（关键事实 4）。
2. **新增 Tauri 命令**（例如 `ai_bind_session_to_project(session_id, project_id)`），
   在用户为会话选定项目之后、发第一条消息之前调用。
3. **让工具把可信 sessionId 带到桥接**：工具 `execute` 里取 `exec.agent?.session.id`
   （关键事实 2），随请求发给桥接服务；**取不到就直接失败**，不要发送。
4. **桥接侧做归属校验**：处理业务逻辑之前，用 sessionId 查绑定表拿到锁定的 projectId，
   与请求里要操作的 projectId 比对，不一致**直接拒绝**。
   fail closed——参照现有令牌鉴权、审批超时自动拒绝的一贯风格，
   不要设计成"默认放行、异常才拒绝"。
5. **先在 `run_benefit_calculation` 上把校验跑通**（关键事实 5）。
   它是只读的，拿来验证"绑定 A、请求 A 放行"和"绑定 A、请求 B 拒绝"两条路径最安全。
6. **前端会话创建流程**：新建会话时从**已有 lamber 项目列表**里选一个绑定
   （不是任意文件系统文件夹选择器），选完调用第 2 条的命令，再允许开始对话。
   已存在的历史会话没有绑定时如何处理，需要一个明确策略（见"需要先定的事"）。

## 本轮执行口径（2026-09-06）

1. **允许通用聊天，但禁用全部工具**；新会话必须显式选择已有项目或通用聊天后再发送。
   > ⚠ **本条已被 ③ 修订（2026-09-06 用户确认）**：通用聊天**允许调用聚合只读工具**
   > （路线图 ③ 的跨项目查询），且**只允许这一类**；其余工具、尤其是所有写工具仍全部禁用。
   > 修订理由与校验二维表见
   > [TASK_BOOK_cross_project_query.md](./TASK_BOOK_cross_project_query.md) 的"已定口径"。
   > 本行保留原文，是为了让后来者看到口径变过、以及为什么变。
2. **历史会话只读保留**；没有后端绑定的历史不补绑，继续工作时创建新身份。项目绑定不可改绑；用户清空/删除后同时撤销绑定。
3. **本轮没有跨项目豁免**。路线图③限定为聚合值与基本信息，禁止其他项目报价、成本明细；工具参数和返回字段须在③任务书列明后单独实现。

以上采用本轮提出的默认方案；并未把前端 `projectId` 元数据当成用户已经选择绑定的证据。

实现与验收见 [会话项目限权验证](../verification/session-project-binding.md)。

## 不要做的事

- **不做写业务数据的工具。** `create_project`、填模板字段等等都在本文之后
  （路线图 ④⑥），本次工具集保持不变，只加校验。
- **不碰审批机制**：`ApprovalGate`、`agent_approval_log`、`AgentApprovalDialog`、
  `GATED_TOOLS` 一律不动。跨项目请求被拒绝是**硬性权限校验**，不是需要用户点确认的审批场景，
  两者不要混着做。
- **不给旧 `AiRuntime` 回退链路加绑定**（关键事实 1），它即将删除。
- **不做任意文件系统文件夹浏览选择器**。绑定范围只在已有 lamber 项目列表里选。
- **不改 `turns.rs` 的现有语义**去兼容项目绑定（关键事实 3 的生命周期差异）。
- 不做多 dsh 子进程 / 进程池改造。保持"整个 App 一个常驻子进程、靠 sessionId 区分会话"。
- 不改 `calculator.rs` / `docfill.rs` / 测算引擎 / NPV / 现金流 / 税额 / 0 容差校验。

## 验证要求

1. **两条核心路径，带真实 key 走完整链路**（不能只测校验函数本身）：
   会话绑定项目 A → 请求项目 A（放行）；会话绑定项目 A → 请求项目 B（拒绝）。
2. **fail closed 路径**：`exec.agent` 取不到时、绑定表查不到时、绑定表与请求都缺
   projectId 时，各验一次，确认都是拒绝而不是放行。
3. **重启后绑定仍在**（关键事实 4）：绑定 → 重启应用 → 恢复会话 → 跨项目请求仍被拒绝。
   这条最容易漏，而漏了就等于绑定形同虚设。
4. **并发**：两个会话分别绑定不同项目，交替发起工具调用，各自的限制互不串台。
5. `cargo test`、`cargo test agent_bridge -- --ignored`（真实 key）、
   `npm run lint/build --prefix src-ui`、`npm run typecheck --prefix agent-bridge/dsh-tool-lamber` 全过。
6. 结果记入 `docs/verification/`，照实区分"已验证"与"未验证 / 已知限制"。
