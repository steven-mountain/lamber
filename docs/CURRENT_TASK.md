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
                              dsh 子进程（--profile sdk）
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
2. **多会话 / 多项目 Agent 面板** —— 目前只有 `#/agent-lab` 这个单会话联调台（`LAMBER_AGENT_LAB` 门控，不是给终端用户的）。真正接进 `AiChatPanel`、支持多会话与项目上下文切换尚未开始。
3. **参数抽取二次确认** —— 模型从自然语言里抽出的参数（金额、项目名、年限等）在执行前让用户核对修正。当前审批弹窗只做"批准 / 拒绝"，不能改参数。

## Scope Boundary

- 未改动 `calculator.rs` / `docfill.rs` / 测算引擎 / NPV / 现金流 / 税额 / 甄选费 / 反算 / 0 容差校验。
- 未改动 `AiRuntime.ts` / `AiChatPanel.tsx` 既有行为。
- 插件包内除 `run_benefit_calculation` 与无害的 `write_test_marker` 外无其它工具；**没有**任何会写 lamber 业务数据的工具。
- 未做 SEA 单文件打包 / 瘦身。
