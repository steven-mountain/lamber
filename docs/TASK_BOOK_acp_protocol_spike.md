# 任务书（小/验证性质）：确认 Rust `agent-client-protocol` crate 与 dsh-acp 协议版本能否握手

这是 ACP 协议层重写正式开始前的**唯一前置验证**，排在其它所有 ACP 相关任务书之前。只回答一个问题：Rust 官方 crate `agent-client-protocol`（当前查到 2.0.0，schema 1.7.0）跟 dsh 的 `dsh-acp`（内部绑定 `@agentclientprotocol/sdk@1.4.0`）能不能真正握手、走完一轮完整的请求-响应。这两个版本号体系不确定是否 1:1 对应，此前没有做过真实验证，纯靠读代码/文档判断不出来。

## 要做的事

1. 新增一个 `acp` profile，跟现有 `sdk` profile 平行、不冲突：新建 `.dsh-home/profiles/acp/` 目录，`package.json` 里 `dsh.profile.bundles` 设为 `["@deepseek-ai/dsh-base", "@deepseek-ai/dsh-acp-app"]`（对照现有 `sdk` profile 的写法），跑一遍现有的 `agent-bridge/scripts/provision-profile.mjs`（传 `--profile acp`）确认能正常装起来，`dsh --profile acp` 能启动不报错。
2. 在 Rust 侧加 `agent-client-protocol`（crates.io 官方 crate）依赖，写一个**完全独立、不接入现有 `AgentRuntime`/`dsh_session.rs`/主应用**的小型验证程序（放 `src-tauri/examples/` 下的一个 example binary，或者一个带 `#[ignore]`、需要真实 `DEEPSEEK_API_KEY` 才跑的 `cargo test`），只做这几件事，顺序不能跳：
   - spawn `dsh --profile acp` 子进程；
   - 用这个 crate 的 client builder 完成 `initialize` 握手，确认协议版本协商**不报错**；
   - 调 `session/new` 建一个会话；
   - 发一条真实的 `session/prompt`（简单文本即可，比如"回复 ok 两个字"），确认真的拿到模型回的内容，走完一次完整的请求-响应闭环；
   - 顺手确认一下这个 crate 是否提供类似 `on_receive_request` 的方式注册 `session/requestPermission` 的 handler——只要确认这个 API 存在、能注册成功即可，不用真的触发一次权限请求走完整流程（这条是加分项，不是这次的硬指标）。
3. 把结果如实记下来，放进 `docs/verification/`：crate 版本、dsh-acp 实际协商到的协议版本号、握手是否成功、完整 round trip 是否成功，任何报错的原始信息全部照抄，不要美化、不要说"大概能行"。最后给一句话结论：**兼容 / 不兼容 / 部分兼容（说明具体卡在哪一步）**。

## 不要做的事

- 不要把验证程序接入现有 `AgentRuntime`/`dsh_session.rs`/`mod.rs` 主链路，不改动任何现有生产代码路径——这次完全独立，跟主应用零耦合。
- 不要动 `dsh-tool-lamber` 的 TS 插件、`approval.ts`、审批弹窗、`agent_approval_log`——这些是协议重写正式开始后才涉及的东西，这次不碰。
- 不要现在就实现真正的 `session/requestPermission` 权限应答逻辑，也不要现在就解析 `session/update` 流式输出——这次只验证"能不能连上、能不能握手、能不能走完一轮请求响应"，别的都不用做。
- 不要删除或修改现有 `sdk` profile 和现有闭环 A/B 相关代码，两套 profile 要能并存、互不影响。
- **不要在没有真实跑通的情况下宣称"协议兼容"**——必须是代码真的跑起来、看到真实的握手成功和模型返回内容，才算通过；只是编译通过、没有实际运行不算数。如果握手失败或版本不兼容，如实报告失败原因和完整报错信息，不要尝试绕过/降级/换个没验证过的版本硬凑出一个"能跑"的假象。
