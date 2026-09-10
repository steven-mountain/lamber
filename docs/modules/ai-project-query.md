# AI 跨项目查询（路线图③）

## 权限与数据流

风险目标是控制上下文体积、防止把检索到的其他项目数字误作当前项目结论。
`query_projects` 是唯一聚合只读白名单工具，绑定会话与显式通用会话均可使用。
未知/未登记会话仍拒绝；缺少 projectId 不构成授权。原测算仍需绑定且项目匹配，通用会话不能使用测算或测试写入工具。

插件 `queryProjects.ts` 从运行时的 `exec.agent.session.id` 获取身份，POST `/lamber-bridge/query-projects`。
Rust `ProjectBindings.authorize` 返回明确的 `AggregateRead` / `Project` 分类；桥接按路由固定实际工具名，不能用请求体伪造工具分类。
绑定机制、持久化、工作区一致性、原审批机制均保持不变。

`project_query.rs` 调用 `ProjectService.get_projects`，与既有 `get_projects` Tauri 命令复用同一入口。
没有新 SQL 或明细 JOIN。现有仓储解析 projects.summary_metrics；按 default_scheme_id 调用既有 get_schemes，仅取对应方案的阶段、名称和保存时间元数据，不读快照、模板或财务明细。
投影为独立 DTO 后才序列化；不会直接序列化整个 Project，因此路径、note、logs 和将来新加字段都不会自动进入模型。
插件输出也使用封闭 schema，拒绝返回结构漂移。

## 查询契约

所有参数可选，不接受 SQL、表达式或 projectId。未知参数/排序名/错误类型拒绝。

| 参数 | 规则 |
| --- | --- |
| customerName | 不区分大小写的字面包含匹配，非正则/SQL LIKE |
| status / benefitStatus | 精确匹配；文本筛选 1–128 字符 |
| createdFrom / updatedFrom | 包含起点 |
| createdBefore / updatedBefore | 不包含终点；可表示完整自然年/月份 |
| 时间格式 | RFC3339 可携带时区；YYYY-MM-DD、无时区日期时间按 UTC 解释 |
| totalRevenueIncl / totalCostIncl | 含税元；范围对象支持 gte / gt / lte / lt |
| marginRate / npvRate / irr | 比例小数，20%=0.2；低于20%应传 `{lt:0.2}` |
| npv / dynamicPayback | 元 / 年，同一范围对象 |
| sortBy | name、createdAt、updatedAt、totalRevenueIncl、totalCostIncl、marginRate、npv、npvRate、irr、dynamicPayback |
| sortOrder | asc / desc，默认 updatedAt desc；同值按 id 稳定排序，缺失值始终排最后 |
| limit | 默认20，硬上限50，超大请求压至50；0/负数/小数拒绝 |

区间无边界、上下限颠倒或严格边界相交为空时拒绝。带时间/指标过滤时，无法解析的值不参与命中。
指标支持数字字符串；比例兼容带 `%` 的存储值。缺失、非数字或非有限指标返回 null，不按零处理。
项目名、客户名、状态与风险标签限制展示字符数，超长加省略号；项目 id 保留原值用于识别。

## 输出与防混淆

每条仅含 id/name/customerName/status/benefitStatus、createdAt/updatedAt/progress、projectYears/discountRate、totalRevenueIncl/totalCostIncl、六个 summaryMetrics 指标、isBoundProject，以及 defaultSchemeId/stage/stageLabel/schemeName/schemeUpdatedAt。
返回 boundProjectId，通用会话为 null，所有条目的 isBoundProject 为 false。

固定 notice 说明这是跨项目检索、不得替代当前项目测算结论；指标来自已保存摘要而非重新测算，保留 benefitStatus 供说明过期状态。
`matchedCount` 是全体命中；`returnedCount`、`appliedLimit`、`truncated` 与中文 message 明示返回数量，禁止静默截断。
`stageTotals` 按阶段覆盖全体命中而非截断后的列表，包含收入/成本及有效金额条数；金额按已有 Decimal 类型累加，溢出明确报错。
`mixedStages=true` 时 `totals=null`，`comparisonNotice` 明说不能横向比较；财务排序先按阶段分组，再在组内排序，不产生跨阶段排名。
单一标签可返回 `totals`，但所有行都未标注也不等于口径已确认。缺失默认方案、找不到所指方案或未知阶段返回 `unlabeled/未标注`，绝不从其他方案猜测。

`summary_metrics` 的既有含义就是最后保存的默认方案汇总；本次不改保存与默认方案语义，不重算指标。阶段只接受 `pre_selection`（甄选前/限价口径）、`post_selection`（甄选后/中标口径），缺失保留未知。
回答指标时必须同时给阶段标签、方案名与保存时间。默认方案切换只改变读取的元数据，不把另一个方案快照的数字拼入原汇总。

前端 `sessionScopePolicy.ts` 统一选择说明、通用聊天标签与模型指导。自动上下文仍只提供绑定项目；通用聊天不自动装入业务状态，由工具按参数读取。
模型回答“我这个项目”时只能使用绑定 id / isBoundProject=true 的行，未命中不能借其他行或合计替代。
该行为需观察真实多轮模型回答；权限测试本身不能替代真人核对。

## 验证

[验收记录](../verification/cross-project-query.md)；真实回答与工具值见 [evidence](../verification/cross-project-query-real-evidence.json)，真人核对样本单独留档。
