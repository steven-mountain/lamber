# dsh 阶段 2：产品接入与能力对照

日期：2026-09-05。主机：macOS。dsh：锁定的 `0.1.2-alpha.5`。

## 结论

聊天面板已接入默认关闭的 dsh 试用开关，旧 `AiRuntime.ts` 未修改。
取消、跨进程恢复、视觉图片均已真实模型验证；**逐 token 流式展示仍缺失**。
因此不能宣称 dsh 已完整覆盖旧 Chat，不能进入阶段 3。
Gate 1A-2b 已在独立 macOS 测试包验证模型保存与进程替换；Gate 1A 整体仍有未验证项，Windows Gate 1B 继续保留。

## 对照表

| 能力 | dsh 实现 / 已知缺失 | 验证与边界 |
| --- | --- | --- |
| 正文、思考、工具增量 | 已实现事件适配；**缺失逐 token 实时展示** | `DshRuntime.ts` 按类型累加正文/思考、按 toolCallId 合并稀疏工具更新。真实 ACP 事件回放通过。当前 dsh 在 durable `assistant/message` 后才生成 chunks；前端无法提前获得 token |
| 停止生成 | 已实现 | `ai_cancel_prompt → Command::Cancel → session/cancel`。透明回环代理观察到官方 API 的真实 delta 后立即取消，实测 11–18ms 收到 `Cancelled`，同会话下一轮成功 |
| 多会话隔离 | 已实现 | 发送前固定前端 sessionId/requestId；Rust 在排队前关联 ACP id，先订阅再 invoke。测试覆盖早到事件、切换可见会话、迟到片段及旧轮次取消 |
| 重启恢复 | 已实现 | 本机独立 `ai-sessions.sqlite` 保存映射，`session/resume` 恢复原 id。真实测试重启后回忆随机暗号通过；错误 cwd、缺失 id 明确拒绝，不新建兜底 |
| 上下文注入 | 已实现 | 复用 `buildAiChatContext + PromptRenderer`，每轮通过 ACP text block 注入已保存状态、未保存草稿、单位规则与用户问题。纯函数测试检查两层金额、规则及用户意图保留；未以真实业务项目验证回答质量 |
| 图片输入 | 已实现 | inline ACP `ImageContent`，最多四张 PNG/JPEG/WebP、每张 5MB。真实 `deepseek-v4-flash-vision-exp` 对测试红色 PNG 回答“红色。”；普通模型在投递前明确拒绝 |
| 模型配置 | 已实现 | 复用设置卡片；恢复的会话通过标准 `session/set_config_option` 同步当前模型，避免沿用持久化旧选型。产品设置从普通模型切到视觉模型后，旧 Node 92732 退出，新 Node 97422 识图成功；详细 PID 与凭据边界见 Gate 1A 记录 |
| baseURL 配置 | 已实现；不宣称通用兼容 | DeepSeek 官方端点实测；透明代理验证真实请求进入配置端点；阶段 0 严格 OpenAI 端点拒绝 dsh 自有字段的结论仍有效 |
| 旧 Chat 回退 | 保留 | 开关默认关闭；旧 endpoint/model/key 界面及 `AiRuntime.ts` 保留。两条链路不自动迁移对方已有的服务端历史；对照时使用新会话 |
| 审批 | 既有实现保持 | 自动化模拟审批通过；本轮另用 Computer Use 在产品浮窗点击确认/拒绝，观察自然超时及无工作区拒绝，见下表；不冒充用户真人验收 |

## 实流事实与失败实验

1. 三个阶段 1 真实模型用例均带 key 执行成功，无“无 key 跳过”。测试只使用合成项目数据与临时标记，不读取真实业务项目。
2. 第一次长文实验等待 ACP 首个 chunk 后取消，得到 `EndTurn`，**该实验失败**。源码确认 `assistantUpdates` / `onSessionEvent` 仅投影已提交消息，首块出现时整条消息可能已生成完成。
3. 未修改 dsh 源码，也未用前端打字动画伪装实时流。通过只转发到 DeepSeek 官方 API 的透明回环代理，观察到真正的 upstream delta 后取消，证明模型请求运行中能停止。
4. 取消实验收到一个 `agent_thought_chunk`（部分思考内容）后 `Cancelled`；下一轮收到完整中文正文、`usage_update`、`EndTurn`。本次未观察到思考/正文交错，也未观察到拆开的 UTF-8 字符；不能据此宣称不存在这种边界情况。
5. 实际事件载荷保存在 `src-ui/scripts/fixtures/dsh-real-stream.json`，均为合成提示词输出，不含密钥或客户数据。适配器另以构造事件验证交错、稀疏更新和竞态。

## 验证命令

凭据仅通过测试进程环境注入，不写入仓库、fixture、日志或本文件。不要把 key 写在命令历史中。

- `cargo test --manifest-path src-tauri/Cargo.toml`
- `cargo test --manifest-path src-tauri/Cargo.toml agent_bridge -- --ignored --nocapture --test-threads=1`（真实 key）
- `python3 scripts/verify-dsh-stage2.py`（真实 key；专门验证模型生成中的取消）
- `npm run test:dsh --prefix src-ui`
- `npm run lint --prefix src-ui`
- `npm run build --prefix src-ui`
- `npm run typecheck --prefix agent-bridge/dsh-tool-lamber`

`stage2_real_stream_cancels_and_accepts_next_turn` 的测试观察代理使用 Python 3 标准库，
属于开发验证设施，不随安装包发布。代理不记录请求体和 Authorization。

## 最终自动化结果

- Rust 常规测试：77 passed、0 failed、13 ignored。
- 全量 agent_bridge ignored 测试带真实 key：13 passed、0 failed、0 ignored；无缺 key 跳过。
- 前端 `test:dsh`、lint、build、插件 typecheck 全部通过。
- macOS 独立测试 `.app` 构建成功；这不能代替 Windows NSIS Gate 1B。

## 产品界面验证（Computer Use）

使用独立 bundle id `com.cmcc.benefitcalc.stage2`，不覆盖正式应用配置。
合成工作区 `/private/tmp/lamber-stage2-ui-workspace`；凭据只在 debug 测试进程环境中，
设置文件未保存 key。本表是工具实际点击与观察，不是用户真人签收。

- dsh 开关初始关闭；启用后从产品聊天取得“产品聊天验证通过。”。
- 普通模型显示不支持图片；保存视觉模型后可添加图片，实际上传仓库测试 PNG 并回答“红色。”。
- 应用重启后保留前端历史；模型切换前后前端 id 与 ACP id 映射不变，cwd 为合成工作区而非仓库。
- 浮窗原先未挂载审批组件，导致提示不可见。现在浮窗挂载既有 `AgentApprovalDialog`，后端只向一个承载窗口发送审批事件（优先 AI 浮窗，否则主窗口），避免双窗重复。
  `ApprovalGate`、审计实现、对话框本体及 `GATED_TOOLS` 未改。

| 路径 | 产品界面 / 审计结果 | 证明边界 |
| --- | --- | --- |
| 确认 | 浮窗显示 `write_test_marker` 与参数「阶段2确认」；点击确认后工具已完成；审计 `approved=1, decided_by=user` | 只写系统临时标记，未修改业务项目 |
| 拒绝 | 浮窗显示参数「阶段2拒绝」；点击拒绝后工具失败并回复未写入；审计 `approved=0, decided_by=user` | 未绕过原审批机制 |
| 超时 | 修复后浮窗显示「阶段2超时」及倒计时，不点击；自然 90 秒后关闭、工具失败，审计 `approved=0, decided_by=timeout` | 请求 06:04:47.158 UTC，决定 06:06:17.160 UTC；后续普通消息可继续 |
| 无工作区 | 取消临时工作区关联后发消息，明确返回 `NotReady` / “请先新建或打开 Lamber 工作区”，无新进程、无审批请求 | 阶段 1 已将此路径前置拒绝；旧联调台“无工作区仍弹审批并 spool”的流程已不适用，spool 机制本轮未改 |

仍待：用户真人验收、Gate 1A 其它未验证项、Windows 1B、真实业务项目上下文回答质量。
逐 token 输出是明确的上游缺失，不能凭本表勾成完整覆盖。

审批审计留痕（2026-09-05，UTC；测试 SQLite 只读提取）：

```text
05:50:24.816304  阶段2确认  approved=1  user     用户已确认
05:51:28.151142  阶段2拒绝  approved=0  user     用户已拒绝
06:06:17.159999  阶段2超时  approved=0  timeout  等待用户确认超时（90 秒），按拒绝处理
```

修复前另一次不可见弹窗的 timeout 记录保留在测试库中，不计作修复后超时 UI 证据。
