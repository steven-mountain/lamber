# 路线图：让 AI 真正能查、能写业务数据

> 背景：dsh 融合（`TASK_BOOK_dsh_full_integration.md`）已把**链路**打通——默认走 dsh、
> 流式、取消、多会话、上下文注入、图片识别都在。但**工具面还是空的**：
> 只有 `run_benefit_calculation`（只读 `benefit_schemes`）和 `write_test_marker`（无害测试）。
> 所以 AI 现在只能看上下文里塞给它的东西，不能自己查，更不能写。
>
> 本文排四项能力的顺序与前置，不替代各自的任务书。

## 架构决定：不给 AI SQL 接口（不要重新讨论）

**AI 调用的是业务操作，不是 `UPDATE`。** 三条理由：

1. `AGENTS.md` 明令 "No direct AI database writes"——写库必须由用户动作触发。
2. **19 张表里有一批是计算引擎的产物，不是数据**：`project_lifecycle_states`、
   `project_cashflow_states`、`benefit_schemes` 由 `calculator.rs` 算出。直接改它们等于绕过
   测算引擎，0 容差校验、税额闭合、NPV 全部失效——**而且不报错，只会安静给出错的财务数字**。
3. **已有 132 个注册好的 Tauri 命令**，每个都带参数校验、事务、回滚。
   正确做法是**把选定命令包一层当工具**，先例见
   `TASK_BOOK_create_intelligent_compute_project.md`："`create_project_in_workspace`
   已经做好校验和回滚，直接包一层当 Agent 工具用即可，不用重新设计创建逻辑。"

同理，"跨项目查询"也**不是**开放 SQL，而是参数化的检索工具
（按客户、年份、金额区间、指标区间过滤，返回结构化摘要行）。

## 三块前置机制的状态

| 机制 | 状态 | 谁需要它 |
| --- | --- | --- |
| 审批弹窗 + 审计落库 | ✅ 已做，四条路径真人验证 | 所有写工具 |
| 会话绑定项目 + 硬性限权 | ✅ 已实现并通过真实链路与重启限权验证 | 所有工具（读也要，否则能翻别的项目） |
| 审批弹窗能读能改（原⑤） | ❌ 完全没做 | **所有写工具**（不只是写财务数据的） |

**"审批弹窗能读能改"是最容易被低估的一块**，而且它的位置**原先排错了**。
现在弹窗只会 `JSON.stringify` 参数、只能批准/拒绝。原路线图以为它只是 ⑥ 的前置
（数字抽错要能改），**但 ④ 写几百字中文散文时它同样是硬前置**：
JSON 里转义换行的一坨散文没法审阅，用户只会盲批——
那"任何写操作都需要人工审核"就成了名义要求。
已并入 ④ 作为其阶段 A，见下。

---

## 执行顺序

### ✅ ① 图片与附件入库　＋　模板 FormData 缺陷（已完成）

2026-09-06：实现及核心产品验证完成，模型真实口述/系统图片粘贴端到端待验收。见[验证记录](../verification/template-state-and-chat-assets.md)。

- 任务书：[TASK_BOOK_demand_analysis_image_completion.md](./TASK_BOOK_demand_analysis_image_completion.md)（解冻）
  ＋ [TASK_BOOK_template_form_state_generation.md](./TASK_BOOK_template_form_state_generation.md)
- **为什么能马上做**：图片写入由**用户点击**触发既有 `save_template_asset`，
  不需要新写工具、不需要审批、不需要会话绑定。
- **为什么要合并**：两者是同一个方向的改动——把 `TemplateForms` 的取值统一到保存态、
  不再依赖当前挂载的 DOM。分开做要动同一个文件两遍。
  图片任务书第 1 步"缺项规则抽成纯函数"与 FormData 缺陷的修法是同一件事。
- 这也是用户第一轮提出的原始目标，做完就闭环。

### 接下来（按前置顺序，不可颠倒）

**② 会话绑定项目 + 硬性限权（已完成）**

- 任务书：[TASK_BOOK_cowork_session_project_binding.md](./TASK_BOOK_cowork_session_project_binding.md)。
- 可信 ACP 身份 + 会话登记表 + SQLite 持久化；允许绑定项目、拒绝跨项目的真实链路已通过，重启 resume 后仍拒绝。
- 新会话选择项目或通用聊天；通用聊天禁用全部工具；未绑定历史保留并另建会话。
- 本轮未开放跨项目豁免。验证：[session-project-binding.md](../verification/session-project-binding.md)。

**③ 跨项目查询与汇总（实现及自动验证完成，待真人核对）**

- `query_projects` 已实现，81项常规/15项真实集成通过。[验收与真人核对](../verification/cross-project-query.md)。

- 任务书：[TASK_BOOK_cross_project_query.md](./TASK_BOOK_cross_project_query.md)
  —— ✅ **2026-09-06 已写好，口径已定，可直接开工**。
- **口径（已确认）**：通用聊天**允许**调用聚合只读工具，且**只允许这一类**；
  其余工具、尤其所有写工具仍全部禁用。此条**修订了 ② 的"通用聊天禁用全部工具"**，
  ② 的任务书已就地标注。理由：否则用户为问一句全局问题就得先随便绑个项目，会诱导乱绑。
- **风险模型已摆正**：这里的白名单目的是**控上下文体积 + 防模型混淆**，**不是防泄露**
  （整个工作区都是同一用户的数据，模型已能看到当前项目全部成本）。
  真正高发的风险是"拿别的项目的数字回答当前项目的问题"。
- 绑定 A / 访问 A 放行、访问 B 拒绝已在②的原测算工具上验证；新增查询工具必须复用同一硬性边界。
- 不需要审批（先例：`run_benefit_calculation` 不在 `GATED_TOOLS` 里）。
- 设计约束：参数化检索，不是 SQL；返回结构化摘要而非整表。
- **实现规则已定为结构性的**：只读 `projects` 表的行 + 解析 `summary_metrics`，
  **不 JOIN 任何明细表**。这比逐字段维护白名单可靠——报价、供应商、科目明细都在别的表里，
  不 JOIN 就天然进不来，将来加字段也不会漏网。
- ②的校验不认识"没有 projectId 的工具"，需新增"聚合只读"分类，
  **必须按工具名白名单 fail closed**，不能写成"没有 projectId 就放行"。

**④ 帮我填模板字段（写，中风险）＋ ⑤ 审批弹窗升级（已合并，⑤ 提前）**

- 任务书：[TASK_BOOK_ai_write_template_fields.md](./TASK_BOOK_ai_write_template_fields.md)
  —— ✅ 2026-09-06 已写好，可开工。分两阶段，**A 不过不做 B**。
- **路线图修订**：原先把 ⑤（参数二次确认）排在 ④ 之后、只当 ⑥ 的前置。**那是错的。**
  ④ 写的是几百字中文散文，而弹窗现在是 `JSON.stringify`（`AgentApprovalDialog.tsx:99-102`）
  ——转义换行的一坨，旁边还有倒计时。**看不懂的东西点确认，不叫审核**，
  用户"任何写操作都需要人工审核"的要求会名义成立、实质落空。
  而且散文比数字更需要"改后再批"：只能批准/拒绝会逼用户反复拒绝重说。
- 阶段 A：弹窗按写入意图结构化展示、**显示新旧对照**（覆盖 vs 填空是两个决定）、
  支持修改后批准、审计同时记模型原值与用户改后值、写类工具倒计时延长但**超时仍拒绝**。
- 阶段 B：`fill_template_fields` 复用既有保存链路；**只写文本字段**，
  金额/税率/测算结果**服务端拒绝**；走 ② 绑定校验，通用聊天一律拒绝。

**⑥ 新建项目 / 改测算参数（写，最高风险）**

- 任务书：[TASK_BOOK_create_intelligent_compute_project.md](./TASK_BOOK_create_intelligent_compute_project.md)
  （已写好，覆盖智算项目；ICT 项目需另写）
- 依赖 ②③ 与 ④ 的阶段 A（审批弹窗升级）全部就位。
- 硬边界：不改 `calculator.rs` / 测算引擎本体、NPV、现金流、税额、甄选费、0 容差校验。

---

## 依赖关系一览

```
①图片+FormData ──（独立，可立即开工）

②会话绑定 ──→ ③只读查询 ──→ ④A 审批弹窗升级 ──→ ④B 填模板字段
                                     │                    ↑
                                     │                    ①（共用缺项规则）
                                     └──→ ⑥建项目/改参数
```

## 每步都适用的既有约束

- 写业务数据必须由用户确认触发，不得静默写入（`AGENTS.md`）。
- 新增写工具必须登记进 `dsh-tool-lamber/src/approval.ts` 的 `GATED_TOOLS`。
- 审批机制本体（`ApprovalGate`、`agent_approval_log`、`AgentApprovalDialog`）不要改，
  它已通过四条路径真人验证，是目前最稳的一块。
  > **唯一例外：④ 的阶段 A**，且扩展范围严格限定为——展示层、
  > 决定内容（bool → bool + 可选的修改后参数）、审计字段（同时记模型原值与用户改后值）。
  > **不得改动** `ApprovalGate` 的超时与失败关闭语义、`gated_tool_names_match_the_plugin`
  > 这类守卫、`GATED_TOOLS` 的判定位置，以及"未决即拒绝"的默认方向。
  > 除此之外仍然一律不碰。
- 每项完工后按任务书惯例落 `docs/verification/`，照实区分"已验证"与"未验证/已知限制"。
