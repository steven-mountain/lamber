# 会话项目绑定与工具限权验收（2026-09-06）

## 结果

路线图②已实现。绑定项目 A 的会话只允许调用 A 的测算；请求真实存在的项目 B 被硬性拒绝。
关闭 dsh 进程、重新打开会话数据库并 ACP resume 后，该限制仍生效。
此结论来自生产 Rust dispatcher + 实际 dsh 工具执行，不能仅靠模型口头拒绝来替代。

## 已实施规则

- 新会话显式选择已有项目或通用聊天。通用聊天禁用全部工具，并不读取业务上下文。
- 无后端绑定的历史保留记录；继续时创建新身份。绑定不可改绑，清空/删除同时撤销映射和绑定。
- `exec.agent.session.id` 是可信身份来源，模型 schema 中没有 sessionId 参数。
- 只有原来的测算和无害测试标记两项工具具备权限规则；未知工具、缺失身份/项目/绑定、工作区变化、已删除项目全部拒绝。
- 现有审批、财务计算、现金流、NPV、税额与文档生成逻辑不变。
- 本轮无跨项目豁免。③只允许聚合值与基本信息，须另写任务书列明字段白名单，禁止其他项目报价与成本明细。

## 自动验证

| 检查 | 结果 | 证据范围 |
| --- | --- | --- |
| `cargo test --manifest-path src-tauri/Cargo.toml` | 80 passed，14 ignored | 常规用例、HTTP 权限入口、持久化、并发与原财务回归 |
| `cargo test --manifest-path src-tauri/Cargo.toml agent_bridge -- --ignored --test-threads=1 --nocapture`（真实 Key） | 14 passed，0 skipped | 实际 dsh、模型、工具、审批、取消、视觉和恢复 |
| `dsh_project_scope_real_tool_denial_survives_restart` | passed；补充真实存在的 B 后再次通过 | A 正常放行；关闭 dsh、SQLite reopen、resume 后 B 返回 403 |
| `project_scope_http_isolation_persistence_and_fail_closed` | passed | 请求缺 sessionId / projectId、未知绑定、通用会话、删除项目、跨项目方案、未知工具、工作区切换、撤销绑定 |
| 双会话并发 | passed | A/B 两线程各循环 6 次，对自身返回 200、对对方返回 403，共 24 次交替请求 |
| 插件 `npm run test:scope` | passed | 无 agent 在网络/写文件之前拒绝；伪造参数 sessionId 不生效；拒绝不进入审批；未知工具单调 guard；执行体重验 |
| 插件 `npm run typecheck` | passed | 工具运行时上下文类型契约 |
| 前端 `npm run lint` / `npm run build` | passed | TypeScript、ESLint、Vite；存在原有大 chunk 提示 |
| 前端 `npm run test:dsh` | passed | 原流式投影、取消、历史持久化和身份隔离 |
| 前端 `npm run test:template-state` | passed | 前项模板状态与图片补齐回归 |
| 前端 `npm run test:session-scope` | passed | 绑定 A 只读 A；指名 B 不代用 A；不同页面草稿不混入；通用聊天不读业务；临时旧链路保持原路由 |

常规 Rust 编译仍有既有 unused/dead-code 警告。没有增加依赖。

## 独立 macOS 应用实测

应用：`Lamber Scope Test.app`，独立 identifier `com.cmcc.benefitcalc.scope-test`。
工作区：沿用上一项的 `/private/tmp/lamber-template-20260906-workspace` 合成项目，未改正式业务工作区。
使用本机已配置的官方服务凭据，测试数据是合成项目；凭据不写入本记录、仓库或工具结果。

1. 打开 AI 面板，未选择权限时不显示消息输入；选择器仅列出已有 Lamber 项目。
2. 选择“3A流式验收项目”并绑定后输入开放；顶部显示“已绑定：3A流式验收项目”。
3. 从正式聊天 UI 请求 `run_benefit_calculation`，工具状态“已完成”，回复 NPV `37914.69`、利润率 `0.40`，与合成测算输入一致。
4. 新建并选择“通用聊天”，顶部显示“通用聊天 · 工具已禁用”；完成选择不再额外生成空白会话。
5. 切回项目会话，原绑定和工具结果保持。完全退出应用再启动，历史和绑定项目名称仍在。
6. 恢复后从 UI 请求其他 projectId 的报价与成本，模型明确拒绝并提示另建目标项目会话，没有执行工具；独立测试库仍只有原来的一条 ACP 映射，确认续聊使用 resume。该 UI 提示与上方 Rust 真实工具 403 测试分别记录，不混为同一证据。
7. 已结束测试应用并移除测试配置中的临时凭据副本。结构化结果见 [evidence JSON](./session-project-binding-evidence.json)。

## 验证边界

- 并发限权是在生产 HTTP dispatcher 上的双线程验证；未宣称两个 GUI 窗口同时生成已完成真人验收。
- 真实重启越权验证使用 dsh 子进程重启 + SQLite reopen + ACP resume；macOS 完整应用重启另外核验了显示和续聊。
- Windows 打包/真人稳定性观察未在本轮重新验收。
- 本轮只完成②，未实现③聚合、④填模板或⑥项目创建工具。默认策略可按后续产品决定调整，不能通过提示词或审批直接突破现有权限。

## 复验入口

- Rust：`src-tauri/src/agent_bridge/tests.rs` 的 `project_scope_http_isolation_persistence_and_fail_closed` 与 `dsh_project_scope_real_tool_denial_survives_restart`。
- 插件：`agent-bridge/scripts/test-project-scope.mjs`。
- 前端：`src-ui/scripts/test_session_scope.cjs`。
- 设计：[ai-session-workspace.md](../modules/ai-session-workspace.md#项目权限路线图②)。
