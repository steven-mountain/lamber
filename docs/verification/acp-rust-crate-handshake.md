# ACP 协议兼容性验证：Rust `agent-client-protocol` ↔ dsh `dsh-acp`

> **结论：兼容。** 四步全部在真实进程间跑通，模型返回内容已拿到（`ok`）。
> 验证日期 2026-09-04，macOS 15.6 / rustc 1.94.1 / Node `dsh@0.1.2-alpha.5`，真实 DeepSeek 模型
> `deepseek-official` / `deepseek-v4-flash`。

回答的唯一问题：crates.io 官方 crate `agent-client-protocol` 跟 dsh 的 `dsh-acp`
（内部绑定 `@agentclientprotocol/sdk@1.4.0`）能不能真正握手、走完一轮完整的请求-响应。

---

## 版本事实

| 项 | 值 | 来源 |
| --- | --- | --- |
| Rust crate `agent-client-protocol` | **2.0.0** | `src-tauri/Cargo.lock:23-24` |
| 其传递依赖 `agent-client-protocol-schema` | **1.5.0** | `src-tauri/Cargo.lock:57-58` |
| dsh CLI | `@deepseek-ai/dsh@0.1.2-alpha.5` | `agent-bridge/package.json` |
| `@deepseek-ai/dsh-acp` | `0.1.2-alpha.5` | 其 `package.json` |
| 其绑定的 TS SDK | `@agentclientprotocol/sdk@1.4.0`（精确固定，非 `^`） | `dsh-acp/package.json` |
| **实际协商到的 ACP 线上协议版本** | **`1`**（整数） | 见下方原始报文 |

### 关于「2.0.0 / 1.7.0 / 1.4.0 这几个版本号对不对得上」

任务书里的疑问在这里可以明确回答：**这三个号根本不是同一套体系，不需要对得上。**

* **ACP 线上协议版本是一个整数**，跟包版本号无关。schema 里写得很清楚：
  > `"description": "Protocol version identifier.\n\nThis version is only bumped for breaking changes.\nNon-breaking changes should be introduced via capabilities."`，`"type": "integer", "format": "uint16"`
* TS SDK 1.4.0 的根导出 `PROTOCOL_VERSION = 1`
  （`node_modules/@agentclientprotocol/sdk/dist/schema/index.js:51`）。
  同包另有一个 v2 命名空间导出 `PROTOCOL_VERSION = 2`（`dist/v2/schema/index.js:42`），但 `dsh-acp`
  是从**根**导入的（`dsh-acp/lib/index.js:9`），所以它说的是 **v1**。
* Rust crate 2.0.0 里对应的是 `ProtocolVersion::V1` 与 `schema::v1::*`；它也带
  `unstable_protocol_v2` feature，本次**未启用**。
* 另外一处与任务书不符：任务书写「crate 2.0.0，schema 1.7.0」，但 crate 2.0.0 的清单把
  schema **精确锁死**在 `=1.5.0`（`agent-client-protocol-2.0.0/Cargo.toml:139-140`），
  所以实际解析到的是 1.5.0，拿不到 1.7.0。这不影响握手，但记录在案。

**一句话：双方都在讲 ACP v1，握手时交换的是整数 `1`。**

---

## 验证程序

`src-tauri/examples/acp_handshake_probe.rs` —— 一次性探针，**与主应用零耦合**：

* 不 `use` 本 crate 任何模块，不碰 `AgentRuntime` / `dsh_session.rs` / `mod.rs`；
* 不碰 `dsh-tool-lamber` 插件、`approval.ts`、审批弹窗、`agent_approval_log`；
* `agent-client-protocol` 与 `tokio` 放在 **`[dev-dependencies]`**，
  发布的 `benefit-calculator` 二进制**不增加任何依赖**。

新增的 `acp` profile（`agent-bridge/.dsh-home/profiles/acp/`）与现有 `sdk` profile 平行，
两者互不影响；`sdk` profile 与闭环 A/B 相关代码未改动。

```bash
# 一次性准备
cd agent-bridge && node scripts/provision-profile.mjs --profile acp

# 跑探针
cd src-tauri && DEEPSEEK_API_KEY=<key> cargo run --example acp_handshake_probe
```

---

## 逐步结果

| # | 步骤 | 结果 |
| --- | --- | --- |
| 0 | `acp` profile 装起来、`dsh --profile acp` 能启动 | ✅ 通过 |
| 1 | spawn `dsh --profile acp` 子进程 | ✅ 通过 |
| 2 | `initialize` 握手、协议版本协商 | ✅ 通过，协商到 `1` |
| 3 | `session/new` | ✅ 通过 |
| 4 | `session/prompt` 真实模型返回 | ✅ 通过，模型回 `ok`，`stopReason: end_turn` |
| + | `session/requestPermission` handler 可注册（加分项） | ✅ 通过（编译+注册，未触发真实权限请求） |

### 0. `acp` profile

`agent-bridge/.dsh-home/profiles/acp/package.json`：

```json
{
  "name": "dsh-profile-acp",
  "private": true,
  "dsh": {
    "profile": {
      "bundles": ["@deepseek-ai/dsh-base", "@deepseek-ai/dsh-acp-app"],
      "patchReload": "startup"
    }
  }
}
```

`node scripts/provision-profile.mjs --profile acp` 原样输出（含一条既有的无害告警）：

```
[provision] DSH_HOME=/Users/hermesjang/Documents/CMCC/tools/lamber/agent-bridge/.dsh-home
[provision] linking /Users/hermesjang/Documents/CMCC/tools/lamber/agent-bridge/dsh-tool-lamber into profile "acp"…

dependencies:
+ dsh-tool-lamber link:/Users/hermesjang/Documents/CMCC/tools/lamber/agent-bridge/dsh-tool-lamber

Already up to date
Done in 192ms using pnpm v10.34.5

dsh: warning: dsh-tool-lamber declares no dsh.bundle — installed as a plain dependency, not a profile layer (a later update that gains one activates it automatically)
```

`dsh --profile acp --help` 正常（证明 profile 能启动、bundle 装配无误）：

```
Usage: dsh --profile acp [options]

Serve automation clients over Agent Client Protocol stdio.

Options:
  -h, --help  show this help

Example:
  dsh --profile acp     serve ACP until the client disconnects
```

> 注：provision 脚本会把 `dsh-tool-lamber` 也链进 `acp` profile（脚本行为如此，未改脚本）。
> 但 `acp` profile 的 `bundles` 里没有它，探针也没传 `--patch patch.yml`，所以本次验证中
> **该插件未被挂载**，握手结果不受它影响。

### 2. `initialize` —— 握手成功

原始报文（探针的 `with_debug` 回调逐行抄录，`>>` 为发出，`<<` 为收到）：

```
>> {"jsonrpc":"2.0","id":"9301b542-5f0c-4619-b56d-beb037b66d7d","method":"initialize","params":{"protocolVersion":1,"clientCapabilities":{"fs":{"readTextFile":false,"writeTextFile":false},"terminal":false}}}
<< {"jsonrpc":"2.0","id":"9301b542-5f0c-4619-b56d-beb037b66d7d","result":{"protocolVersion":1,"agentInfo":{"name":"deepseek-harness-acp","version":"0.0.1"},"agentCapabilities":{"mcpCapabilities":{"http":true},"promptCapabilities":{"image":false,"audio":false,"embeddedContext":false},"sessionCapabilities":{"close":{},"list":{},"resume":{}}},"authMethods":[]}}
```

Rust 侧解析结果：

```
协商到的 protocolVersion : ProtocolVersion(1)
agentInfo               : Some(Implementation { name: "deepseek-harness-acp", title: None, version: "0.0.1", meta: None })
agentCapabilities       : AgentCapabilities { load_session: false, prompt_capabilities: PromptCapabilities { image: false, audio: false, embedded_context: false, meta: None }, mcp_capabilities: McpCapabilities { http: true, sse: false, meta: None }, session_capabilities: SessionCapabilities { list: Some(SessionListCapabilities { meta: None }), delete: None, additional_directories: None, resume: Some(SessionResumeCapabilities { meta: None }), close: Some(SessionCloseCapabilities { meta: None }), meta: None }, auth: AgentAuthCapabilities { logout: None, meta: None }, meta: None }
authMethods             : []
```

**无告警、无降级、无 unknown-field 报错**：crate 2.0.0 的 `InitializeResponse` 把 dsh 返回的
capabilities 完整解析成了结构体（包括 dsh 侧没发的字段，都落到了 `None` / `false` 默认值）。

一个值得记下的实现细节：**dsh 侧不做版本协商，是硬编码回 `1`**。
`dsh-acp/lib/index.js:1143-1146` 的 `initialize(_params)` 直接忽略入参，
无条件返回 `protocolVersion: PROTOCOL_VERSION`。也就是说客户端请求任何版本它都答 `1`——
本次刚好一致，但这意味着**将来 dsh 升到 v2 时不会有握手期的报错来提示我们**，
版本不匹配只会在后续某个字段上炸开。接协议层时值得加一条客户端侧的断言。

### 3. `session/new` —— 成功

```
>> {"jsonrpc":"2.0","id":"d2964da2-5efa-4be3-9ef5-8a136261ede4","method":"session/new","params":{"cwd":"/Users/hermesjang/Documents/CMCC/tools/lamber/src-tauri","mcpServers":[]}}
<< {"jsonrpc":"2.0","id":"d2964da2-5efa-4be3-9ef5-8a136261ede4","result":{"sessionId":"b6b80f9e-943e-481b-9863-cdc2e24ed5f3","configOptions":[...]}}
```

`sessionId : SessionId("b6b80f9e-943e-481b-9863-cdc2e24ed5f3")`

`configOptions` 里 dsh 报了可选模型（`deepseek-v4-flash` / `deepseek-v4-pro` /
`deepseek-v4-flash-vision-exp`，当前 `deepseek-official`/`deepseek-v4-flash`）与
reasoning effort（当前 `high`）。crate 2.0.0 能吃下这个字段，未报错。

### 4. `session/prompt` —— 完整 round trip 走通

发出提示词「回复 ok 两个字」，收到模型真实返回。原始报文全文：

```
>> {"jsonrpc":"2.0","id":"8f992195-9199-45c8-b65f-827bffe6aa8c","method":"session/prompt","params":{"sessionId":"ea82b13e-1825-4678-b9f1-f55757364203","prompt":[{"type":"text","text":"回复 ok 两个字"}]}}
<< {"jsonrpc":"2.0","method":"session/update","params":{"sessionId":"ea82b13e-1825-4678-b9f1-f55757364203","update":{"sessionUpdate":"agent_message_chunk","messageId":"f40a5e49-2f1e-45a7-975a-42156f6800ea","content":{"type":"text","text":"ok"}}}}
<< {"jsonrpc":"2.0","method":"session/update","params":{"sessionId":"ea82b13e-1825-4678-b9f1-f55757364203","update":{"sessionUpdate":"usage_update","used":11509,"size":1000000}}}
<< {"jsonrpc":"2.0","id":"8f992195-9199-45c8-b65f-827bffe6aa8c","result":{"stopReason":"end_turn"}}
```

Rust 侧（crate 2.0.0）解析结果：

```
[session/update] AgentMessageChunk(ContentChunk { content: Text(TextContent { annotations: None, text: "ok", meta: None }), message_id: Some(MessageId("f40a5e49-2f1e-45a7-975a-42156f6800ea")), meta: None })
[session/update] UsageUpdate(UsageUpdate { used: 11509, size: 1000000, cost: None, meta: None })
stopReason : EndTurn

== 完整 round trip 走通：收到 2 条 session/update 通知 ==

结果：全部四步通过。
```

**模型返回的正文是 `ok`**，与提示词「回复 ok 两个字」一致——是真实模型输出，不是桩。
`usage_update` 报 11509 tokens，也证明真的打到了上游。

按任务书要求，本次**没有实现 `session/update` 的结构化解析**：探针只对通知做 `{:?}` Debug
打印。但这已足够证明 crate 2.0.0 能把 dsh 发的两种 `sessionUpdate`
（`agent_message_chunk` / `usage_update`）反序列化成强类型 enum 变体而不报 unknown-variant。

> 复现记录：第一次跑（未设 `DEEPSEEK_API_KEY`）在这一步失败，报
> `Internal error: turn failed: llm-deepseek: no API key for provider route "deepseek-official"; store DEEPSEEK_API_KEY through the credentials service (the web Models page writes it), or export DEEPSEEK_API_KEY in the launching environment`。
> 该错误来自 `llm-deepseek` 模型适配层而非 JSON-RPC 编解码；设置真实 key 后同一份代码一次通过。
> 记在这里是为了说明：**这一步的通过与否只取决于凭证，与协议兼容性无关。**

### 加分项：`session/requestPermission` handler

crate 2.0.0 提供 `Builder::on_receive_request`，配合 `on_receive_request!()` 宏做类型注册。
探针里已注册：

```rust
.on_receive_request(
    async move |request: RequestPermissionRequest, responder, _cx| {
        match request.options.first().map(|opt| opt.option_id.clone()) {
            Some(id) => responder.respond(RequestPermissionResponse::new(
                RequestPermissionOutcome::Selected(SelectedPermissionOutcome::new(id)),
            )),
            None => responder.respond(RequestPermissionResponse::new(
                RequestPermissionOutcome::Cancelled,
            )),
        }
    },
    agent_client_protocol::on_receive_request!(),
)
```

编译通过且连接建立时注册未报错 —— **该 API 存在、能注册**。按任务书要求，本次
**未触发真实权限请求**，也未实现真正的应答逻辑（当前是无脑选第一项，仅为让类型完整）。

---

## 结论

**兼容。**

Rust 官方 crate `agent-client-protocol@2.0.0` 与 dsh `dsh-acp@0.1.2-alpha.5`
（绑定 `@agentclientprotocol/sdk@1.4.0`）**能够正常握手并走完完整的请求-响应闭环**：
四步（spawn / `initialize` / `session/new` / `session/prompt`）全部在真实进程间跑通，
协商到的 ACP 协议版本为整数 `1`，模型真实返回内容 `ok`、`stopReason: end_turn`。
加分项 `session/requestPermission` 的 handler 注册 API 存在且可用。

全程无版本协商报错、无字段解析失败、无降级、无绕过。

### 接协议层前需要注意的两点

1. **dsh 不做版本协商。** `initialize(_params)` 忽略客户端请求的版本，无条件回 `1`
   （`dsh-acp/lib/index.js:1143-1146`）。本次刚好一致，但这意味着将来 dsh 升到 ACP v2 时
   **握手期不会报错来提示我们**，不匹配只会在后续某个字段上炸开。
   正式重写时建议在客户端侧对 `init.protocol_version` 加一条显式断言。
2. **schema 版本与任务书预期不同。** crate 2.0.0 把 `agent-client-protocol-schema`
   精确锁死在 `=1.5.0`，拿不到任务书说的 1.7.0。本次不影响，但升级 crate 时要重新核对。

### 复现

```bash
cd agent-bridge && node scripts/provision-profile.mjs --profile acp
cd ../src-tauri && DEEPSEEK_API_KEY=<key> cargo run --example acp_handshake_probe
```
