# ACP 下审批链路的真人点击验证

> **状态：四条路径已由真人点击验证通过。** 验证日期 2026-09-04，macOS 15.6，
> `dsh --profile acp`（`@deepseek-ai/dsh@0.1.2-alpha.5`），真实 DeepSeek 模型
> `deepseek-official` / `deepseek-v4-flash`，通过应用内 `#/agent-lab` 联调台操作。
>
> [`approval-channel-manual-check.md`](./approval-channel-manual-check.md) 记录的是
> `--profile sdk` 时期、经由 HTTP 审批通道触发的验证。ACP 下触发机制完全不同
> （`session/requestPermission` 是 dsh 主动向客户端发起的请求），**那份记录不是本次的
> 证据**，本文不引用它的任何结论。它按要求保留未动，仅作历史。

## 为什么必须重验

| | SDK 协议时期 | ACP 之后 |
| --- | --- | --- |
| 谁发起问题 | 插件的 `approval/request` 监听器 | `dsh-acp` 转成 ACP 请求发给客户端 |
| 走哪条路 | `POST /lamber-bridge/approval`（lamber 起的回环 HTTP 服务） | ACP 的 stdio 连接，反方向 |
| 谁在等 | 桥接服务的一个工作线程 | ACP 连接上 `spawn_blocking` 出去的一个阻塞任务 |
| 工具名/参数从哪来 | 请求体里带着（插件的 `pendingCalls` 表） | 更早那条 `session/update` 的 `tool_call` 通知（Rust 的 `ToolCallIndex`） |
| 弹窗文案从哪来 | 守卫写的 `reason`，随请求体传过来 | Rust 侧的镜像表（`dsh-acp` 把守卫的 `reason` 丢掉了） |

被复用的部分（`ApprovalGate`、`agent_approval_log`、`ai://approval-request`、
`ai_resolve_approval`）确实没改，但**它们上游的每一环都换了**。编译通过、`cargo test` 全过、
带 key 的集成测试通过，都只说明机器认为链路是通的；弹窗里显示的东西对不对、点下去有没有反应，
只有人点了才知道。

## 前置条件

```bash
cd agent-bridge && npm install && (cd dsh-tool-lamber && npm install) && npm run provision
```

`provision` 现在默认链的是 **`acp`** profile。确认这一步的输出里有
`linking … into profile "acp"`。

启动应用（需要真实的 `DEEPSEEK_API_KEY`）：

```bash
DEEPSEEK_API_KEY=<你的 key> ./run.sh
```

然后在应用里访问 `#/agent-lab`（Agent 联调台）。这是真实应用中唯一能驱动 Agent、
进而触发审批弹窗的入口。

## 路径一 / 二：确认与拒绝

联调台默认指令即可（`请调用 write_test_marker 工具，note 参数填「真实点击联调」。`）。

1. 点「发送」。
2. 等待弹窗出现。
3. 分别做一次「确认执行」和一次「拒绝」（两次独立的发送）。

| 检查项 | 确认执行 | 拒绝 |
| --- | --- | --- |
| 弹窗是否出现 | ✅ | ✅ |
| 弹窗显示的工具名是否为 `write_test_marker` | ✅ | ✅ |
| 弹窗显示的参数是否为模型真实传的 `note` | ✅ `真实点击联调` | ✅ |
| 弹窗文案是否为镜像表里那句「该工具会写入文件……」 | ✅ | ✅ |
| 倒计时是否在走 | ✅ | ✅ |
| 点击后弹窗是否立即关闭 | ✅ | ✅ |
| 审计表 `decided_by` | ✅ `user` | ✅ `user` |
| 审计表 `approved` | ✅ 已批准 | ✅ 已拒绝 |
| 系统临时目录下是否新增标记文件 | ✅ 有 | ✅ 无 |
| 联调台事件流里 `tool_call_update` 的 `status` | ✅ `completed` | ✅ 未执行 |

审计日志原文（联调台右栏，2026-09-04）：

```
已批准 write_test_marker · user · 用户已确认 · 2026-09-04T08:57:40.733350+00:00
已拒绝 write_test_marker · user · 用户已拒绝 · 2026-09-04T08:57:44.821740+00:00
```

事件流原文（确认那一轮，节选；证明工具名与参数确实来自 `tool_call` 通知）：

```
session/update · tool_call {"sessionId":"70e58328-…","update":{"sessionUpdate":"tool_call",
  "toolCallId":"call_00_yyVXrY1hv9kEamuJC47y2124","title":"write_test_marker",
  "status":"in_progress","rawInput":{"note":"真实点击联调"}}}
session/update · tool_call_update {…,"toolCallId":"call_00_yyVXrY1hv9kEamuJC47y2124",
  "status":"completed","content":[{"type":"content","content":{"type":"text",
  "text":"已写入测试标记文件: /var/folders/…/lamber-agent-marker-N8Q6mT/
  marker-2026-09-04T08-57-40-743Z.txt (212 字节, 2026-09-04T08:57:40.743Z)"}}]}
```

落盘旁证：`marker-2026-09-04T08-57-40-743Z.txt`，212 字节，写入时刻与审计表里那条
「已批准」相差 **10 毫秒**——先有决定、后有执行，顺序正确。拒绝那一轮没有产生任何新文件。

> 参数一致性还有一条机器侧的断言：`dsh_gated_tool_runs_only_after_the_user_confirms`
> 会比对弹窗拿到的 `args` 与 `tool_call` 通知里的 `rawInput` 是否相同。人工这里要看的是
> **显示出来的东西是否对人有意义**，不是它们是否相等。

## 路径三：超时

1. 发送同一条指令。
2. 弹窗出现后**什么都不点**，等满 90 秒（或用 `LAMBER_APPROVAL_TIMEOUT_SECS` 调短）。

| 检查项 | 结果 |
| --- | --- |
| 倒计时归零后弹窗是否自行关闭 | ✅ |
| 审计表 `decided_by` | ✅ `timeout` |
| 审计表 `approved` | ✅ 已拒绝 |
| 是否新增标记文件 | ✅ 无 |
| dsh 是否继续正常响应下一条指令（没有卡死在这一轮） | ✅ |

审计日志原文：

```
已拒绝 write_test_marker · timeout · 等待用户确认超时（90 秒），按拒绝处理
  · 2026-09-04T08:59:24.007406+00:00
```

理由文本里的「90 秒」来自 `ApprovalGate` 自己的超时分支，不是 dsh 那边的兜底——ACP 下
dsh 根本不给权限请求设上限，这条记录就是这个唯一兜底真的会触发的证据。

> ACP 下 dsh **不给权限请求设上限**，就等客户端应答。SDK 时期插件答复器那道 180 秒兜底
> 已经不存在，所以这 90 秒是**唯一**能结束一次无人应答的机制——这条路径比以前更重要。

## 路径四：无工作区

1. 不打开任何工作区，直接进 `#/agent-lab`。
2. 发送指令，弹窗出现后点「确认执行」。

| 检查项 | 结果 |
| --- | --- |
| 弹窗是否照常出现并可点击 | ✅ |
| 决定是否照常生效（标记文件是否写出） | ✅ |
| `agent-approval-spool.jsonl` 是否出现该条记录 | ✅ |
| 随后打开一个工作区，记录是否回填进 `agent_approval_log` | ✅ |
| 回填后缓冲文件是否被删除 | ✅ |

## 证据来源，以及本文没有证明什么

分清楚哪些是硬证据、哪些是操作者的观察，免得日后把两者当成一回事：

* **有留痕的硬证据**：审计日志四条记录（含时间戳、`decided_by`、批准与否、理由文本）、
  联调台事件流里的 `tool_call` / `tool_call_update` 原文、系统临时目录下的标记文件
  及其字节数与 mtime。这些在本文里都贴了原样。
* **由操作者当场观察、无逐项留痕的**：弹窗的渲染、倒计时是否在走、点击后是否立即关闭、
  路径四缓冲文件的出现与消失。这些属于"人看着屏幕确认"，没有截图逐项对应。

**本文没有证明的事**（避免被过度引用）：

1. 没有覆盖 `write_test_marker` 以外的工具。`GATED_TOOLS` 目前只有它一个。
2. 没有验证 dsh 自带工具（bash、文件编辑等）触发权限请求时的表现。ACP 下 lamber 是唯一的
   权限应答方，那条路径存在但未走过。
3. 没有做并发审批（同时挂起两个问题）的验证。
4. 超时那条按真实的 90 秒走的，没有验证 `LAMBER_APPROVAL_TIMEOUT_SECS` 覆盖后的行为。

## 结论

**四条路径（确认 / 拒绝 / 超时 / 无工作区）在 ACP 的 `session/requestPermission` 触发入口
下全部通过。** 审批机制在 ACP 下的迁移可以认为完成。

值得单独记一笔的是，事件流证实了这次改动里最关键的那个结构性判断：工具名与参数确实是从
更早那条 `tool_call` 通知（`title` / `rawInput`）里取到并显示在弹窗上的——权限请求本身
只带 `toolCallId`。`tool_calls.rs` 这个索引不是可有可无的补充，它是弹窗能说清"要执行什么"
的唯一来源。
