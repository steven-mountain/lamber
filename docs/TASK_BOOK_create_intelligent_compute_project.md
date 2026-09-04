> **排序更新**：这份任务书排在 `docs/TASK_BOOK_cowork_session_project_binding.md`（Cowork 会话绑定项目 + 工具调用硬性限权）**之后**执行。本文里的工具都是"绑定会话里可用的工具"，前置的会话绑定/限权机制没做完之前不要开始这份。

# 任务书：模块 1 — 智算项目创建 + 缺项校验（对话式修改默认预设）

范围：仅 `project_type: "intelligent_compute"`。ICT 项目类型不在本次范围内。

## 已确认的关键事实（写代码前不用重新调研）

- `create_project_in_workspace(name, customer_name, project_type)`（`src-tauri/src/benefit/commands.rs:150-279`）已经做好校验和回滚，直接包一层当 Agent 工具用即可，不用重新设计创建逻辑。`project_type="intelligent_compute"` 时会自动调用 `ensure_project_state` + `ensure_default_amount_source`，失败会连 DB 行带目录一起回滚。返回完整 `Project`。
- 智算项目的"金额来源包"读写命令都已存在，不用新增：
  - 读：`get_intelligent_compute_project(project_id) -> IntelligentComputeProjectData { state, amount_sources }`（`intelligent_compute.rs:322-353`）。
  - 写金额来源：`save_intelligent_amount_source(project_id, { source: IntelligentAmountSource, expected_version })`（`intelligent_compute.rs:437-540`），整体覆盖式保存 + 乐观并发（`source_version` 对不上会报 `IntelligentAmountSourceVersionConflict`）。
  - 写项目设置：`save_intelligent_compute_project_state(project_id, { expected_version, active_amount_source_id, project_years, discount_rate })`（`intelligent_compute.rs:355-435`），同样乐观并发（`state_version`）。
- **默认预设的坑**：`ensure_default_amount_source` 只在 DB 里插入一条 `name="H200 标准智算金额来源"` 的空壳记录——`parameters/revenueItems/costItems/mappings` 全是空数组，`calculationSnapshot` 是 `{}`。真正的"64 台 H200"数值（`device_count=64`, `gpu_service_price=90000`, `machine_price=3300000` 等全套参数/收入项/成本项）只存在于前端 `src-ui/src/features/ai-compute-quote/presets.ts:88-160` 的 `createH200Blueprint()`，且只有前端页面打开时才会因为"检测到没有业务数据"而回填。**通过后端命令创建的项目，如果不主动调用 `save_intelligent_amount_source` 把这些数字写进去，DB 里就只有一个名字对但内容为空的假预设。**
- **`project_years` 默认坑**：`create_project_in_workspace` 建项目时 `project_years` 默认是 `1`，不是 5。`5 年服务期`这句描述文字目前只是摆设，除非显式调用 `save_intelligent_compute_project_state` 把它设成 5。
- **计算链路不通**：`run_benefit_calculation`（闭环 A 已有工具，`agent-bridge/dsh-tool-lamber/src/runBenefitCalculation.ts`）读的是 `benefit_schemes`/`benefit_snapshots` 表（`agent_bridge/calculation.rs`），这两张表只有 `save_benefit_scheme`（ICT 流程）会写。智算项目走的是完全不同的另一条路：前端 `AiComputeQuoteView.tsx` 调 `sync_intelligent_compute_to_ict`（`project_state/mod.rs:891-931`）直接调 `calculate_ict_benefit`，结果写进 `project_lifecycle_states`/`project_cashflow_states`，从不碰 `benefit_schemes`。**结论：对刚创建的智算项目直接调用现有的 `run_benefit_calculation` 会报错"项目「X」还没有任何测算方案"，这一步现在是断的，必须先接上。**
- 审批机制原样复用闭环 B：`agent-bridge/dsh-tool-lamber/src/approval.ts` 里的 `GATED_TOOLS` 是唯一需要登记新工具名的地方；Rust 侧审批弹窗事件、`agent_approval_log` 落库、超时/失败关闭都不用碰。
- 前端审批弹窗 `AgentApprovalDialog.tsx` 已经是通用 JSON 展示（`JSON.stringify(current.args, null, 2)`），不需要为了展示新工具参数而改前端，除非发现某个工具传的 args 大到不可读（见下面"要做的事"第 4 条）。
- `write_test_marker` 的执行模式（execute() 里在 dsh 进程本地直接写文件）**不适用**于本次新工具——本次工具要真的写 lamber 的项目数据，dsh 进程本身连不到 Tauri，必须走新的 Rust bridge route（模式参照现有 `CALCULATE_ROUTE`/`APPROVAL_ROUTE`，在 `src-tauri/src/agent_bridge/mod.rs:63-91` 的 `match path` 里加）。

## 要做的事

1. 新增三个 dsh 工具（`agent-bridge/dsh-tool-lamber/src/`），各自对应一个新的 Rust bridge route：
   - `create_intelligent_compute_project`（需审批）：调用 `create_project_in_workspace(name, customer_name, "intelligent_compute")`，创建成功后**在同一次已批准的操作里**依次调用 `save_intelligent_amount_source`（把 H200 预设的真实数值写进去）和 `save_intelligent_compute_project_state`（把 `project_years` 设为 5、`discount_rate` 设为 0.055）。工具的返回值里要带上实际写入的预设数值，让模型"念"给用户的是数据库里真实存在的数字，不是凭训练知识编的。
   - `update_intelligent_compute_amount_source`（需审批）：接收"要改哪几项"，先用 `get_intelligent_compute_project` 取当前 `amount_source`（拿到 `source_version`），把要改的字段合并进去，再调 `save_intelligent_amount_source` 传回 `expected_version`。同时支持通过参数捎带修改 `project_years`/`discount_rate`（走 `save_intelligent_compute_project_state`）。
   - `run_intelligent_compute_calculation`：新增（不要复用 `run_benefit_calculation`，那个读的表智算项目从不写）。参照 `sync_intelligent_compute_to_ict`（`project_state/mod.rs:891-931`）的路径写一个新 bridge route，直接调 `calculate_ict_benefit` 拿到结果返回。是否需要审批：参照 `run_benefit_calculation` 现有先例——它不在 `GATED_TOOLS` 里，本工具也不用挂审批（计算/同步不是用户数据的写入，是纯衍生结果）。
2. H200 预设数值只能有一份权威来源。把 `presets.ts:88-160` 的具体数字移植到 Rust 侧（或者做成一份两边都读的静态 JSON），两处代码互相加注释指向对方文件，避免以后前端改了预设、后端新建工具却还在用旧数字。
3. 把新增的两个写工具（create / update）加进 `agent-bridge/dsh-tool-lamber/src/approval.ts` 的 `GATED_TOOLS`，参照 `write_test_marker` 现有条目的写法。
4. `update_intelligent_compute_amount_source` 传给审批弹窗的 `args` 只放"实际要改的字段"，不要把整个几百行的 `IntelligentAmountSource` 对象原样塞进去——否则用户在弹窗里看到的 JSON 没法审阅，违背审批机制"让用户看清将写入什么"的目的。
5. 验证要求（延续闭环 B 已验证的标准，不能只测函数本身）：
   - `cargo test` 全过，无回归。
   - 带真实 `DEEPSEEK_API_KEY` 的集成测试，覆盖三个新工具从 `tool/call` → 桥接 → DB 真实写入/读出 → `tool/result` 的完整链路。
   - 针对 create / update 两个新审批点，四条路径（确认/拒绝/超时/无工作区）至少各验证一次；确认/拒绝要走真实点击，不能脚本模拟（沿用已知限制：本机 `osascript` 辅助功能授权无效）。
   - 结果记录进 `docs/verification/`（新建文件或追加到 `approval-channel-manual-check.md`），照实区分"真验证"和"未验证/已知限制"。
6. 完工后更新 `docs/CURRENT_TASK.md`：在最上面新增一节记录本次改动，并把现有"尚未开始的部分"第 1 项标成已完成，说明还剩什么（比如"更细粒度的字段级审批文案"之类如果没做完）。

## 不要做的事

- 不做 `project_type: "ict"` 的对话式创建/校验流程。
- 不改 `calculator.rs` / `docfill.rs` / 测算引擎本体、NPV、现金流、税额、甄选费、反算逻辑。
- 不做参数二次确认 UI（模型抽取参数后的人工核对修正）——这是 `CURRENT_TASK.md` 里"尚未开始的部分"第 3 项，单独任务书，不要在本次顺手做。
- 不接多会话 `harnessSessionId`（"尚未开始的部分"第 2 项），前端多会话 UI 已经做完，本次不碰。
- 不做图片输入相关改动（模块 3，可行性未确认，需要先单独调研 dsh 的 `contentBlocks` 协议）。
- 不改 `AgentApprovalDialog.tsx` 的展示逻辑，除非新工具的 args 经验证确实无法正常展示（当前判断它已经足够通用）。
- 不新增 DB migration，除非验证过程中发现现有表结构确实不够用——现有表（`intelligent_compute_amount_sources`、`project_intelligent_compute_states`、`agent_approval_log`）理论上够用，缺的是"怎么调用"，不是"缺列"。
- 不用 mock/假 API key 做验证，不能只跑 `execute()` 单元测试就算完——必须走真实 dsh 子进程 + 真实审批弹窗的端到端链路，参照闭环 A/B 已验证的标准。
- 不要在没有明确写明"已使用真人点击验证"的情况下，把审批四条路径记为"通过"。
