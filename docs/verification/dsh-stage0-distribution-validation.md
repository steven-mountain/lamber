# dsh 完全融合 · 阶段 0 前提验证

- 验证日期：2026-09-05
- 主机：macOS arm64
- Node：v24.14.0
- dsh：`0.1.2-alpha.5`
- 结论：阶段 0 唯一阻塞项（脱离仓库目录启动）已通过，可以进入阶段 1。

## 1. ACP 图片能力

Rust 的 `AgentHandshake` 已保留
`agentCapabilities.promptCapabilities.image`，`ai_agent_status` 也会返回
`supportsImagePrompts`。

实测命令：

```bash
cd src-tauri
cargo test agent_bridge::tests::acp_handshake_reports_image_capability_for_each_catalog_model -- --ignored --nocapture
```

| ACP 配置模型 | `promptCapabilities.image` | 预期 | 结果 |
| --- | ---: | ---: | --- |
| `deepseek-v4-flash` | `false` | `false` | 通过 |
| `deepseek-v4-flash-vision-exp` | `true` | `true` | 通过 |

能力值确实由模型目录中的 `inputModalities` 决定，普通模型的 `false` 不是 dsh
缺少图片能力。阶段 1 必须让用户可选择模型，阶段 2 再按这个握手能力决定是否允许发送
ACP image block。

## 2. `baseURL` 指向严格 OpenAI 兼容端点

本机没有安装或运行 Ollama、LM Studio、vLLM。为排除具体服务实现差异，本次启动了一个只接受
OpenAI Chat Completions 标准字段、按标准 SSE 返回的本地临时端点，再通过 dsh ACP 发起完整一轮。

实测命令：

```bash
cd src-tauri
cargo test agent_bridge::tests::strict_openai_endpoint_rejects_dsh_deepseek_extension_fields -- --ignored --nocapture
```

结果：

- `llm-deepseek.config.baseURL` 覆盖成功，请求到达本地 `/chat/completions`。
- 请求字段为 `model`、`messages`、`stream`、`stream_options`、`thinking`、
  `reasoning_effort`、`tools`、`max_tokens`、`dsh_plugin_packages`。
- 严格端点以 HTTP 400 拒绝 `thinking`、`reasoning_effort`、`dsh_plugin_packages`。
- 因此当前 dsh 版本下，“只改 baseURL”不是通用 OpenAI-compatible 服务商支持。

阶段 1 仍提供 baseURL 配置，但 UI 必须标明仅验证过 DeepSeek 官方端点；不得把它宣传为
Ollama 或任意 OpenAI-compatible 服务已受支持。若将来需要通用服务商，按任务书另开 adapter
任务解决这三个字段的投影问题。

## 3. 脱离仓库目录与系统 Node 的握手

在 `/tmp/lamber-dsh-stage0-clean.IdTbyN` 构造独立分发目录：

- 复制 `agent-bridge/node_modules`、`.dsh-home`、`patch.yml`、
  `dsh-tool-lamber` 和 Node 可执行文件；
- 不复制瞬态 `profiles/node_modules.lock`；
- 把 profile 模块链接全部重建为临时分发目录内的链接；
- 扫描并确认没有任何链接仍指向 lamber 源码仓库；
- PATH 仅含随分发 Node、Rust 工具目录和系统基础命令；
- 通过 `LAMBER_REPO_ROOT` 指向临时根运行真实 ACP `initialize` 和 `session/new`。

实测结果：握手通过，`session/new` 返回非空会话 id。独立目录中的组成体积为：

| 组成 | KiB (`du -sk`) | 约 MiB |
| --- | ---: | ---: |
| 完整 `node_modules` | 316,016 | 308.6 |
| 初始化后的 `.dsh-home` | 620 | 0.6 |
| Node v24.14.0 arm64 | 116,344 | 113.6 |

首轮复制测试曾把 `profiles/node_modules.lock` 一并复制，dsh 因等待旧锁超时退出；排除瞬态锁后
通过。这说明模板打包和首次运行初始化必须显式排除 lock 文件，不能原样复制开发机
`.dsh-home`。

## 4. 分发形态体积与可用性

| 形态 | 实测体积 | 可用性 | 结论 |
| --- | ---: | --- | --- |
| 完整 npm `node_modules` | 316,016 KiB（308.6 MiB） | 握手通过 | 可用但包含开发依赖 |
| dsh 改为 production dependency 后 `npm prune --omit=dev` | 293,980 KiB（287.1 MiB） | `dsh --version` 返回 `0.1.2-alpha.5` | 阶段 1 采用 |
| ESM bundle + Node SEA | 115,828 KiB（113.1 MiB） | 进程退出码 139；不可握手 | 不采用 |

生产 prune 说明：当前 `package.json` 把 dsh 放在 `devDependencies`，按现状执行 production
prune 会删除运行时。临时把 dsh 的精确版本移动到 `dependencies` 后，基于已锁定的 npm
安装树执行 `npm prune --omit=dev`，得到上表 287.1 MiB。

不能直接把 `pnpm prune --prod` 用作发布步骤：当前目录是 npm 布局且没有 `pnpm-lock.yaml`，
实测时 pnpm 先把现有 alpha 依赖移入 `.ignored`，再从注册表重新解析当时的 rc 依赖，既不离线、
也不保持锁定版本；该尝试已终止，未把它伪装成有效体积结果。阶段 1 应使用与
`package-lock.json` 一致的 npm production 安装/校验流程。

SEA 说明：esbuild 能产出 380 KiB 的 ESM 入口并生成 SEA blob，但 dsh 在运行时仍通过包名
动态装载 Cordis/profile 插件，并读取包清单与原生资产；这些不会自动进入单脚本 bundle。
注入 Node v24.14.0 后的 113.1 MiB 可执行文件启动即崩溃（退出码 139），不能执行
`--version`，更不能完成 ACP 握手。要使 SEA 可用需要重写/虚拟化 dsh 的动态模块与资产加载，
与“不改 dsh 源码”的边界冲突，因此本阶段不采用。

## 阶段 1 分发决策

采用以下有证据支撑的形态：

1. dsh 固定为 production dependency；
2. 发布准备使用 npm 锁文件生成 production-only `node_modules`；
3. 随包携带目标平台 Node；
4. 安装资源只带只读 DSH_HOME 模板，首次运行复制到应用数据目录；
5. 模板不包含 session、credential、匿名 id、瞬态 lock 或开发机绝对链接。

该决策尚未通过干净 Windows NSIS gate。只有在无仓库、无系统 Node、无环境变量的 Windows
机器上从界面配置 key 并拿到首条回复后，阶段 1 才能标记完成。
