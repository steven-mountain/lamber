# AI 长文本审批（阶段 A / B）

## 限定扩展

保留 `GATED_TOOLS` 镜像与权限守卫位置，当前为 `write_test_marker` 与 `fill_template_fields`。阶段 B 只接入需求导入表文本，没有财务字段写入或新数据库迁移。

`ApprovalPrompt.intent` 由 Rust 提供项目名、表单名、实际目标和字段列表，字段含 key、label、previousValue、proposedValue。前端不凭工具参数猜测旧值。测试标记每次创建新文件，所以原值为空；独立诊断演练从实际演练文件读取旧值。

前端长文本对照使用普通字体与保留换行的正文；可就地修改、查看原始建议。没有可编辑意图的操作仍按普通文本只读展示参数。提交失败保留弹窗，可修正或拒绝。

## 决定与真实执行

`ai_resolve_approval(requestId, approved, modifiedArgs?)` 扩展决定内容。Rust 只允许已展示工具的文本参数；允许 marker 的 note 或 fill_template_fields 本次已展示的文本字段（单项最多20000字）。未知字段/工具、非文本、拒绝携带编辑参数均报错，不能借编辑改变项目身份。

ACP 协议本身仍返回一次性 grant/reject。为了让修改值进入真实工具，Rust 在批准时登记 `{可信 sessionId, callId, toolName, originalArgs, approvedArgs}`，插件执行前从鉴权的 `/lamber-bridge/reviewed-arguments` 取回批准参数。它不是第二套审批，只是决定内容的单次交接：

- 先经既有 AUTHORIZE_ROUTE 核对会话与工作区，再精确匹配工具及原始参数；只消费一次。
- 未找到、不匹配、过期或绑定已撤销均拒绝，不能回退执行模型原值。
- 交接最长保存10分钟、最多128项；运行时关闭清空，条目不持久化。
- 路由加入 bridge-contract.json，阶段 B 契约版本3；同代构建检查照常执行。

原 Condvar 等待、拒绝/超时/关闭/无工作区审计流程保持。默认审核由90秒延长至600秒，UI明确说明；`expiresAt` 让排队不重新获得10分钟，Rust 的单调时钟 deadline 最终判定。修改校验失败不会自动批准。

## 审计

沿用 `agent_approval_log.args_json`，不增加表或迁移。新版对象 `auditVersion=2` 含：

- `modelArgs`：模型原参数。
- `userArgs`：用户修改值；未修改为 null。
- `approvedArgs`：最终批准参数；拒绝为 null。它表示审批决定，不冒充执行成功。
- `intent`：当时显示的旧值、字段及目标。

既有 approved/decidedBy/decisionReason/requestedAt/decidedAt 不变，无工作区时继续 spool。联调台兼容旧版裸参数审计，并支持展开三版正文。

## 诊断与人工 gate

设置 → 打开 dsh 联调台 → 长文本审批演练，可无需模型触发同一 gate。命令只由用户点击调用，不登记成模型工具；固定写 app_data_dir/approval-rehearsal.txt，不接受路径参数，不接触业务项目。
再次演练显示上次实际保存内容；批准写修改值、拒绝保留原文件，回执读回实际文件；并发演练拒绝，写入用临时文件替换，防止半写覆盖。

演练内置389字合成样例用于自动化，不能替代任务书要求的300字以上真实需求正文与真人判断。任务书要求 Gate A 真人验收；2026-09-07 用户明确放行阶段 B，保留此前证据的实际验证性质。 验收见 [验证记录](../verification/ai-runtime-retirement-stage-query-approval-a.md)。

## 阶段 B 执行与窗口同步

模板写入由 Rust 在保存事务前直接消费单次交接，并复核审批旧值及模板版本。普通 reviewed-arguments 路由仅允许 marker。详见 [ai-template-write.md](./ai-template-write.md)。

审批请求通过 `emit_to` 投递承载窗口，前端必须显式以当前窗口 label 订阅，不能使用默认 Any 目标，否则多个窗口会重复显示同一请求。成功决定后全窗口接收 `ai://approval-settled` 清理对应队列；校验失败保留待审弹窗。审批策略及默认拒绝方向不变。
