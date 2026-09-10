# dsh 阶段 1：本地分发与待执行 Windows Gate

日期：2026-09-05

## 本地已验证

- `scripts/package-windows.mjs` 的运行资源准备步骤使用锁文件执行 `npm ci --omit=dev`，随后复制当前打包主机的 Node、清洁 DSH_HOME 模板、基础 patch 与 `dsh-tool-lamber` 的 `package.json + lib`。
- 本机生成资源树字节合计为 333.3 MB；`node.exe <staged dsh entry> --version` 输出 `0.1.2-alpha.5`。
- `prepared_release_resource_tree_completes_acp_handshake` 从 staging 资源构造独立用户 home，未使用开发 `.dsh-home`，成功完成 ACP initialize 与 `session/new`。
- 清洁模板复制会拒绝符号链接并忽略瞬态 `node_modules.lock`；模型/baseURL 写入单独 patch，API Key 未写入该 patch。
- 旧版 `config.json` 不含 `aiAgent` 时可反序列化并得到默认模型与官方 baseURL。
- 设置页提供明确的 dsh 联调台入口，联调台可返回主界面；Windows gate 不需要设置 `LAMBER_AGENT_LAB` 才能从 UI 发消息。

## 自动化结果

```text
cargo test
72 passed; 0 failed; 10 ignored

cargo test agent_bridge -- --ignored --nocapture
10 passed; 0 failed
```

ignored 套件中 7 个无需 key 的用例实际完成；3 个真实 LLM 问答/工具闭环因本机未设置 `DEEPSEEK_API_KEY` 而按用例约定提前跳过。该结果不能替代真实 key 验证。

前端 lint、前端生产构建、`dsh-tool-lamber` typecheck 与 Windows 打包脚本测试均通过。

## 干净 Windows Gate（待执行）

此门禁尚未通过，因此阶段 2 不得开始。验证机必须同时满足：无 Lamber 仓库、无系统 Node、无 `DEEPSEEK_API_KEY` / `LAMBER_REPO_ROOT` 等开发环境变量。

待记录：

- NSIS 文件名、SHA-256 与安装包体积：未验证
- 首次启动至设置页可操作耗时：未验证
- 设置页保存真实 API Key、模型和 baseURL：未验证
- 打开用户 workspace 后发送消息并取得首条模型回复：未验证
- 未打开 workspace 时的明确拒绝：未验证
- 安装资源缺件时的用户文案：未验证

通过标准：上述步骤均在干净机完成，且进程树使用安装包内 Node、用户数据目录承载 DSH_HOME；不得依赖源码目录或系统环境变量。
