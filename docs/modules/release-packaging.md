# Windows Release Packaging

## Standard command

Run packaging from the repository root:

```powershell
npm run package:windows
```

This command increments the patch version before building. For example, `1.1.0`
becomes `1.1.1`.

Before version files are changed, the script prepares the bundled dsh runtime
under `src-tauri/resources/agent-runtime/`. It installs the exact production
dependency tree from `agent-bridge/package-lock.json`, builds and copies only
the distributable `dsh-tool-lamber` files, copies the clean DSH_HOME template
and patch, and bundles the Node executable used by the packaging host. The
staging directory is generated and ignored by Git; only its README placeholder
is tracked.

The selected distribution shape is the npm production tree plus Node. Stage 0
measured 308.6 MiB for the complete dependency tree and 287.1 MiB after a valid
production prune. The 113.1 MiB SEA experiment crashed before ACP startup and
cannot support dsh's dynamic package/plugin loading without changing upstream
dsh, so it is not a release option. Do not replace `npm ci --omit=dev` with a
direct `pnpm prune --prod` on this npm lock/layout: the measured attempt changed
the pinned package graph instead of pruning it in place.

## Version policy

- `patch` is the default for routine fixes and packaging iterations:
  `npm run package:windows`
- `minor` is used for backward-compatible features:
  `npm run package:windows:minor`
- `major` is reserved for incompatible product or data-contract changes:
  `npm run package:windows:major`

The release script requires all version sources to match before it starts. It
updates:

- root `package.json` and `package-lock.json`
- `src-ui/package.json` and `src-ui/package-lock.json`
- `src-tauri/tauri.conf.json`
- `src-tauri/Cargo.toml` and the application package entry in `Cargo.lock`

If the Tauri build fails, all version files are restored to their original
contents. A successful build keeps the new version and prints each generated
bundle path, size, and SHA-256 checksum.

## Bundled dsh runtime

Tauri installs the prepared tree as the `agent-runtime` resource. At runtime,
the backend resolves dsh in this order:

1. installed Tauri resource directory;
2. `LAMBER_REPO_ROOT` for explicit development runs;
3. a repository layout found beside a development executable.

The installed resource is immutable. On first AI launch, the clean
`dsh-home-template` is copied to `<app_data_dir>/dsh-runtime`; sessions and all
other dsh state remain there. The application plugin is refreshed from the
installed resource on every launch, while model and baseURL are written to a
separate generated patch. API keys remain in Lamber's application config and
are passed to the child process only at launch.

The ACP `session/new` cwd is always the currently opened user workspace. An AI
launch without a workspace is rejected explicitly; it must never fall back to
the source repository or installation directory.

## Output

The Windows NSIS installer is generated under:

```text
src-tauri/target/release/bundle/nsis/
```

The installer filename contains the incremented version. Build outputs remain
ignored by Git; commit the synchronized version files with the release changes.

## Validation

The versioning rules have a dependency-free Node test:

```powershell
npm run test:packaging
```

For local distribution validation, run `prepareAgentRuntime()` and the ignored
Rust test `prepared_release_resource_tree_completes_acp_handshake`. This proves
the staged Node + production dependency tree + clean user home can complete an
ACP handshake and `session/new` without using the repository runtime tree.

This does not satisfy the release gate by itself. Under the 2026-09-05 gate split,
Windows Gate 1B remains mandatory before a release to real users (the later 2026-09-05 decision permits Stage 3 development). Install
the resulting NSIS package on a clean Windows machine with no repository, Node,
or relevant environment variables; configure a key in the UI and record the
first successful model reply, installer size, and first-start timing.


## Stage 2 application data and release boundary

The bundled runtime layout and dependencies are unchanged. Lamber now stores the
frontend-to-ACP mapping in `<app_data_dir>/ai-sessions.sqlite`, separately from
business databases, while dsh history stays in `<app_data_dir>/dsh-runtime`.
Preserve both locations across upgrades; a mapping without its dsh history must
fail explicitly rather than opening an empty replacement conversation.

The product dsh switch remains off by default. Stage 2 verifies cancellation,
resume, and vision with real models, but alpha.5 ACP only publishes committed
assistant messages, so token-by-token streaming is a known gap. macOS local app
checks do not replace Windows path, bundled node.exe, NSIS, or antivirus checks.
See [Stage 2 verification](../verification/dsh-stage2-product-integration.md).
The Python relay used for real cancellation tests is development-only and is
not included in packaging resources.

## Stage 3A incremental display (2026-09-05)

The compiled plugin now includes `lib/stream.js`. Existing recursive plugin-copy
steps include it automatically; no runtime dependency or packaging shape changes.
The product launch sets `LAMBER_STREAM_DISPLAY=1`; independent bridge consumers
leave it disabled unless they implement the display route. Always rebuild the
plugin before preparing resources; user-home plugin refresh must include this file.

Run `npm run test:dsh-stream` (Node + Python 3 development environment) when
upgrading dsh or preparing a release. This launches the actual packaged-version
plugin against a local SSE fixture and fails if pre-commit deltas disappear.
It complements, and does not replace, Windows NSIS installation validation.

## 业务文档模板：有意不随包分发（2026-09-05 决定）

**这不是遗漏，是刻意为之——出于控制安装包体积的考虑，模板不进 `bundle.resources`。**
后续排查时不要把它当成打包缺陷去"修复"。

### 现状事实

- 运行时读的是 `<工作区>/templates/`（`docfill.rs` 的 `get_available_templates`，
  `module_path` + `templates`）。
- 工作区创建流程**不创建也不填充** `templates/`：`workspace.rs` 只建 `.backups`、
  `.exports` 和项目下的 `assets` / `documents` / `analyses`。
- 仓库里的 `项目全生命周期文件模版/` **仅被测试引用**（`docfill.rs` 的 `#[cfg(test)]` 区，
  1029 行之后）。它不是分发源，也不是运行时来源。
- 模板由用户手工放入工作区。

### 由此产生的两个已知后果（照实记录，不要当成 bug 报）

1. **测试与用户用的是两份不同的模板副本，会漂移。**
   2026-09-05 实测：8 个文件中 2 个 MD5 不同。
   - `【2025版】ICT项目立项签批表…docx`：结构性差异（工作区版 25 个占位符、按科目名+金额分列；
     仓库版 18 个、改用 `PROJECT_INVESTMENT_SITUATION` / `PROJECT_REVENUE_SITUATION` 两段叙述）。
   - `效益分析表 .xlsx`：文件清单与共享字符串完全一致，仅保存格式/元数据差异，内容无变化。
   → 含义：`cargo test` 通过**不代表**用户实际填写的那份模板是对的。改模板时两边都要看。

2. **新装用户的模板页是空的，且静默。**
   `get_available_templates` 在 `templates/` 不存在时走
   `if !template_dir.exists() { return Ok(vec![]); }`，返回空数组、不报错。
   → **这一条与体积决定无关，仍应修**：给出明确空状态提示
   （"未找到模板，请将模板放入〈工作区〉/templates/"）或提供导入入口，
   而不是让用户面对一个没有解释的空列表。

### 模板版本的权威来源

**以工作区 `templates/` 为准**（2026-09-05 用户确认）。仓库副本仅供测试，
与工作区不一致时以工作区为准，不要反向覆盖工作区。

### 待定的业务规则（不要自行拍板）

工作区版立项签批表的投入/收入明细只有 IT、CT 两个槽位
（`SUBJECT_IT_COST` / `SUBJECT_CT_COST` 为科目名称，`IT_INVESTMENT` / `CT_INVESTMENT` 为金额），
而 `totalCost = itCost + ctCost + nonItCost + mixCost`（`TemplateForms.tsx:1526`）含四类。
→ 项目存在非IT/CT或综合类投入时，明细之和小于总额，差额在文档中无交代；收入侧同理。
采购甄选费可写入这两类科目，故该情形可达。
**并入哪一槽、或该表是否仅适用于无这两类投入的项目，需业务方确定后再改代码；
在此之前不要为了让数字看起来对而自行凑数。**

## 桥接契约构建门禁（2026-09-06）

- 唯一版本/能力源：`agent-bridge/bridge-contract.json`。插件生成并编译快照，Rust
  编入同一源文件；不能通过运行环境覆盖两端版本来“制造一致”。
- Cargo 在编译前检查开发插件及存在的 staging 插件的 `lib/contract.generated.js`。
  缺失/不同代直接停止。开发者先构建插件；存在旧 staging 时重新准备分发资源。
- Windows 脚本先 `tauri build --no-bundle`，再检查实际 `.exe` 中的契约与 staging
  插件输出一致，之后才执行 `tauri bundle`。失败恢复版本文件，不生成本次安装包。
- 目标二进制带有可读取的契约字节标记，检查不启动目标程序；可覆盖 Windows GUI
  子系统和跨平台目标。实际旧 macOS 二进制与实际新二进制均已验证检查结果。
- 独立 macOS app 的验证不替代 Windows NSIS 安装 gate。旧二进制无法仅靠插件更新
  获得新版人话错误，需要完整重建/重装。[证据](../verification/bridge-contract-handshake.md)。

### 审批参数契约 v2（2026-09-07）

新增一次性审批参数交接路由后，bridge-contract.json 已升至2。完整构建须包含同代 Rust 二进制与插件 lib；开发和 staging lib 都要重建，不能仅替换插件。
本轮 macOS 独立验证包不代表 Windows NSIS 或真人 gate 已通过，相关发版阻塞继续保留。

### 模板写入契约 v3（2026-09-07）

fill-template-fields 业务路由纳入契约，当前版本3；完整重建开发/staging插件及应用后，已对独立 macOS app 的实际二进制核对契约。打包6项通过不替代 Windows NSIS 安装及 Gate B 真人验收。

### 主动读取契约 v4（2026-09-07 补充）

新增read-template-fields路由，当前契约v4。独立macOS应用与开发/staging插件已完整构建并核对实际二进制；打包6项通过。Windows安装及真人gate仍保留。
