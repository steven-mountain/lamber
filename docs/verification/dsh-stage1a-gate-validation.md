# dsh 阶段 1 · Gate 1A 验证记录

> 本轮补充 1A-1、1A-2b 与 1A-5 的实际结果；2a 保存密钥全链路、3、4 的空项仍待验证。
> 每项都要写清是**自动化测试**验的、还是**产品界面人工/Computer Use** 验的——
> 两者的证明力不同。没跑过的一律写"未验证"，不要因为"应该没问题"就打勾。

- 验证日期：2026-09-05（阶段 2 补跑）
- 主机 / 架构：macOS / Apple Silicon
- 构建形态：macOS debug `.app`，独立 bundle id `com.cmcc.benefitcalc.stage2`；不覆盖正式安装版和配置。
- 凭据处理：用户本轮提供；仅注入测试进程环境，未写入仓库、日志或本文件。

## 背景

原阶段 1 gate 是"干净 Windows NSIS 机器跑通"。因开发机为 macOS，
按[任务书](../tasks/TASK_BOOK_dsh_full_integration.md)的 Gate 拆分决定分为：

- **1A** 平台无关，阻塞阶段 2 —— 本文件。
- **1B** 真 Windows-only，挂起，但**必须在阶段 3 之前补上** ——
  见 [dsh-stage1-local-packaging-validation.md](./dsh-stage1-local-packaging-validation.md) 的待记录表。

---

## 1A-1 · 未打开 workspace 时的拒绝路径

- 验证方式：Computer Use 在产品数据管理界面取消临时工作区关联，再从 AI 浮窗启用 dsh 发送合成消息；只读检查进程与审计表。
- 期望：在启动 dsh 子进程**之前**拒绝；返回结构化错误；不回退仓库根。
- 实际：立即返回 `{"code":"NotReady","message":"请先新建或打开 Lamber 工作区"}`；没有新进程/审批记录，既有空闲 Node PID 97422 不变。Rust 每次发送均先检查 workspace，再处理会话恢复/启动。
- 结论：☑ 通过 ☐ 未通过 ☐ 未验证（本次为已有空闲进程时的无 workspace 拒绝；干净安装首发仍属于 Windows 1B）

## 1A-2 · 设置保存后 dsh 子进程重启并实际生效

**这项分两段，两段都过才算通过。** 前半段只能证明"新配置能用"，
证明不了"改配置会让旧进程换掉"——重启逻辑有 bug 时前半段一样会绿。

### 2a · 保存的配置确实进入 dsh 运行链路

- 验证方式：Computer Use 在产品聊天发消息并观察真实回复；本轮只保存模型，key 通过 debug 测试进程环境传入。
- 期望：产品界面发消息，收到 `agent_message_chunk` 与终结事件。
- 实际：真实回复与结束状态已观察；“从界面保存 key 后进入 dsh”的完整路径未执行，不能用环境 key 代替证明。
- 结论：☐ 通过 ☐ 未通过 ☑ 未验证（保存 key 的完整链路）

### 2b · 旧进程被终止、新配置被新进程读到

- 前置 PID（保存前进程树）：Node 92732，父进程为独立测试 App 92137。
- 验证方式：在子进程存活时再次保存配置 → 核对旧 PID 已终止 → 发下一条消息 → 记录新 PID。
- **配置必须改成一个可观测不同的值**，否则"生效"不可证伪：
  只看到新 PID，无法判断它是否真的读到了新值。
  建议探针：把模型切到 `deepseek-v4-flash-vision-exp`，
  核对 `ai_agent_status` 的 `supportsImagePrompts` 由 `false` 翻为 `true`
  （该值随模型变化已在阶段 0 实测确立）。
- 旧 PID 是否终止：是；在模型设置由 `deepseek-v4-flash` 改为 `deepseek-v4-flash-vision-exp` 并点击保存后，系统查询旧 PID 为不存在。
- 新 PID：97422，父进程仍为 92137；发送下一条图片问题时启动。
- 可观测探针：产品提示由“不支持图片”改为“支持图片”，附件按钮启用；向新进程发红色 PNG，真实回复“红色。”。同时确认配置文件模型已保存为 vision-exp、未保存 key，ACP 会话映射保持不变。这里没有把模型目录静态提示单独当成握手证据，实际图片发送还通过后端握手能力检查。
- 验证方式：Computer Use 点击保存、上传和发送；OS PID 查询、配置/映射文件只读核对。
- 结论：☑ 通过 ☐ 未通过 ☐ 未验证（模型配置更新与进程替换）

## 1A-3 · 安装资源缺件时的用户文案

- 验证方式：（如从 staging 树删除某个运行文件触发）
- 期望：只提示重新安装；**不出现** npm / provision / `agent-bridge` 目录等开发者话术。
- 实际文案原文：
- 结论：☐ 通过 ☐ 未通过 ☐ 未验证

## 1A-4 · 使用分发内 Node 而非系统 Node

- 验证方式：（如将测试进程 PATH 置空后运行）
- 期望：ACP 握手仍完成，进程树使用分发内 Node。
- 实际：
- 结论：☐ 通过 ☐ 未通过 ☐ 未验证

## 1A-5 · 补跑 3 个真实 LLM 用例

阶段 1 记录中这 3 个用例因未设置 `DEEPSEEK_API_KEY` 被跳过。
**它实为阶段 2 的前置**：流式适配器与 cancel 的正确性只能靠真实 chunk 流验证。

| 用例 | 覆盖内容 | 结果 |
| --- | --- | --- |
| `a_finished_turn_reports_its_stop_reason` | 终结事件 | 通过，真实模型 |
| `dsh_tool_call_reaches_the_calculator_and_returns_real_numbers` | 只读测算工具闭环 | 通过，合成项目数据库 |
| `dsh_gated_tool_runs_only_after_the_user_confirms` | 审批确认 / 拒绝闭环 | 通过，测试线程模拟点击，非人工 UI |

- 执行命令：三个用例分别 `cargo test <用例名> -- --ignored --nocapture`；随后全量 `cargo test agent_bridge -- --ignored --nocapture --test-threads=1`。
- 是否仍出现"无 key 跳过"：否，全部真实执行。全量 13 passed、0 failed、0 ignored。
- 验证方式：自动化测试。
- 结论：☑ 通过 ☐ 未通过 ☐ 未验证

### 真实 chunk 流的观察（阶段 2 会直接用到，顺手记下来）

写适配器前需要知道真实流长什么样，这里记录实际观察到的、而不是推测的：

- `agent_thought_chunk` 与 `agent_message_chunk` 是否交错出现：本次采样未观察到交错；dsh ACP 从完整提交消息按块投影，并非逐 token 推送。
- 是否观察到单个 UTF-8 字符被切分到两个 chunk：未观察到；JSON 载荷中的中文保持完整。
- 最后一个 chunk 与 `session/turn-ended` 的先后顺序：本次真实采样先内容，再 usage（若有），最后 turn-ended。
- 其它异常：等待 ACP 首块再取消时已经 EndTurn；改为透明代理观察官方 API delta 后取消，11–18ms 获得 Cancelled，详见阶段 2 验证。

---

## 结论

- Gate 1A：☐ 全部通过，可进入阶段 2 ☑ 尚未全部通过：2a 的保存 key 链路、3 和 4 尚无本轮完整界面/安装证据。
- Gate 1B：仍挂起。阶段 3 之前必须补上，不得因为进入阶段 2 就当它不存在。

## 项目决定：带缺项进入阶段 2（2026-09-05）

由项目负责人拍板，在 **1A-2b 未完成**的情况下开始阶段 2。

- **理由**：2b 属于设置保存后的子进程重启链路，与阶段 2 的 chunk 适配器、cancel、
  会话映射正交。它真正拦的是**发版**，不是开发；带着它做阶段 2 不会产生返工。
- **这不等于 Gate 1A 通过。** 当时 2b 一栏保持未验证；本轮已补真实模型保存/PID 证据后才勾选，不能因为阶段 2 已经开始就打勾。
- **偿还期限**：与 Gate 1B 一并，**阶段 3 切换与下线之前必须补完**。
  阶段 3 的第 1 步是"默认切到 dsh 路径"——那一刻起用户改设置改不动 dsh 就是线上故障。
- 1A 其余四项（1、3、4、5）的结果仍需按上面各节如实填写，不受本决定影响。

## 已知限制（照实写）

- 使用 debug 进程环境 key，本轮没有保存用户密钥；不证明正式发布版的设置密钥路径。
- macOS 独立测试包不能替代 Windows 1B；UI 由 Computer Use 操作，不冒充用户真人签收。
- dsh ACP 的逐 token 推送缺失见阶段 2 对照表。
