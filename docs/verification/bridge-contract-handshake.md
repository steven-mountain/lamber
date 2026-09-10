# 桥接契约握手验收（2026-09-06）

状态：实现、新版应用验证及自动回归完成；**任务书第5项要求的旧版人话显示未通过**，
原因及实测结果见下方。Windows NSIS 和③真人核对仍未验收。

## 根因与修复范围

运行中的旧 Rust 与新插件可以分别更新。原工具直到调用业务路由时才发现不兼容，
404 被转述成数据访问失败。进一步实际追踪发现，dsh 的 ACP 入口还能在其他插件
异步加载时提前接受 initialize，因此只给插件加一次 HTTP 请求不足以消除失败路径。

现在使用一份 `agent-bridge/bridge-contract.json`：生成插件编译常量，同时编入 Rust。
插件 apply 在注册任何工具/hook 前 POST 版本和必需路由；Rust 不匹配拒绝。
插件注册完发明确就绪回执；Rust 在接到该回执前不进行 ACP initialize，不开放会话。
启动错误沿原 Result 通道返回。没有重试、跳过工具或空数组降级。

原路由、未知路由404、审批策略、项目权限和聚合白名单未改。业务错误保留后端 error
正文，仅移除 URL/HTTP 状态码包装；未更改测算引擎、财务数据和正式工作区数据库。

汇总机器记录：[bridge-contract-handshake-evidence.json](./bridge-contract-handshake-evidence.json)。

## 自动验证

| 检查 | 实测结果 |
| --- | --- |
| `cargo test` | 82 passed / 17 ignored / 0 failed |
| `cargo test agent_bridge -- --ignored` | 17 passed / 0 ignored / 0 failed，真实 Key 未跳过 |
| 插件 typecheck | 通过 |
| 插件 test:contract | 实际 HTTP：同代、旧404、409、令牌401、畸形响应、断连；失败零注册、无重试 |
| 插件 test:scope | 缺身份、伪造身份、越权拒绝、审批前拒绝、执行时复查通过 |
| 前端 lint/build | 通过；保留已有大 chunk 提示 |
| 前端 test:dsh / test:session-scope | 通过 |
| `npm run test:packaging` | 6 passed |
| `npm run test:dsh-stream` | 6次提交前增量，最终显示与ACP正文一致 |

新增真实启动测试不使用 LLM：

- Rust 真实 handler 接受同代、拒绝高版本及未知必需路由，令牌仍必需，业务403仍返回原说明。
- 实际 dsh 启动连接404旧服务、409错配服务、401服务及缺少连接配置，`AcpRuntime::start`
  直接返回预期中文错误，无法取得可创建会话的 runtime。
- 在独立 home 修改**实际编译插件**，增加 Rust 不支持的 `/lamber-bridge/future-tool`，
  真实 dsh + Rust 握手在 session/new 前失败。源码、正式插件和业务数据没有被改变。
- 原真实回归覆盖测算、单项目/通用会话、聚合工具、越权、重启恢复、取消、视觉模型、审批。
  常规审批测试覆盖确认/拒绝/超时/关闭，原策略不变。

真实 Key 从现有应用配置读取后仅传入测试进程环境，未输出或复制进测试配置/仓库。

## 实际构建门禁

不只测纯函数：

1. 将 staging 中实际 `lib/contract.generated.js` 的 version 从1改为2，执行真实
   `cargo check`，构建脚本退出失败并显示“AI 组件契约不匹配，请先构建插件并重新准备分发资源”。
   测试以 finally 恢复原编译文件，随后正常测试和 app 构建通过。
2. 新 macOS 实际二进制与 staging 插件一致，`assertBinaryContract` 通过。
3. 开工前保留的实际旧 debug 二进制没有编译契约，检查拒绝，不能进入 bundle。
4. Windows 脚本现在先 build --no-bundle，再比对**实际目标二进制**与 staging lib，
   通过后才 bundle。目标二进制无需运行；本轮未执行 Windows NSIS。

## Computer Use：新版 / 旧版分开记

使用独立 `com.cmcc.benefitcalc.scope-test` 应用及其既有合成工作区，没有冒充正式工作区。

**新版 + 不同代插件：通过。** 从独立资源目录加载 version2 插件，在界面发送前触发
启动，实际可见：

> AI 组件版本不匹配，无法启动。请完整重新构建 Lamber 和 AI 组件，或重新安装最新版本。

没有路由名、404或内部 dispatcher 名称。恢复同代插件后重新构建独立 macOS app。

**新版 + 同代插件：界面真实调用通过。** 在通用测试会话发送 query_projects 请求，
实际 ACP tool_call_update 状态为 completed，模型报告命中1个“3A流式验收项目”。
本次使用已有凭据的进程环境注入，未保存到独立测试配置；首次无凭据尝试只有握手通过，
未将其误记为真实问答通过。

**旧 Rust + 新插件：启动阻断通过，严格文案不通过。** 克隆旧 app，实际运行旧二进制，
让它加载新插件。握手请求遭旧服务404，dsh插件树启动失败，stderr包含准确的中文
版本错配原因。但旧 Rust 已抢先完成 ACP initialize，界面收到：

> ACP session/new 失败: Internal error: the ACP bridge has been disposed

没有进入模型工具轮次，也没有继续逐次调用业务路由；不过该文案仍不是任务书要求的人话。
旧二进制没有新增的就绪同步和错误识别逻辑，**不能靠替换插件给旧调用方补上这些代码**。
需完整重建/重装，新版路径已实测。此项不伪记通过，也不通过修改旧路由或返回空值掩盖。

## ③裸问题复验

真实模型先查询完整合成工作区，看到其他乙项目87%及当前甲项目13%；随后仅问：

> 我这个项目毛利率多少？

回答当前甲项目13%，自动断言通过，未附加保存汇总或项目名引导。
[实际模型回答](./cross-project-query-human-review.md)、[原始结果](./cross-project-query-real-evidence.json)。
**真人核对仍待用户，Computer Use/自动测试均不替代真人验收。**

## 构建交付与限制

已重新构建 `src-tauri/target/debug/bundle/macos/Lamber Scope Test.app`（独立测试身份）。
未覆盖 `/Applications` 中的已安装版本，未发版。真实 Windows 路径、杀软与NSIS安装验收
保持原 Gate 1B 状态。既有工程 warnings 保留，未扩张为无关清理任务。
