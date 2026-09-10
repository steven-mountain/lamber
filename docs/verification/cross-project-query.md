# 跨项目查询验收（2026-09-06）

实现及自动验证完成；任务书要求的防混淆真人核对已提供实际回答，尚待用户确认。

## 已验证

- `cargo test`：81 passed，15 ignored，0 failed。
- `cargo test agent_bridge -- --ignored --test-threads=1 --nocapture`：真实 Key，15 passed，0 skipped。
- 前端 lint/build、`test:dsh`、`test:template-state`、`test:session-scope` 通过；插件 typecheck、`test:scope` 通过。
- 原 financial/calculator/docfill 与审批代码未改；原项目越权、通用会话单项目测算拒绝、重启绑定与审批集成用例保持通过。

| 要求 | 证据 |
| --- | --- |
| 绑定会话聚合放行 | 生产 bridge + 实际 dsh + 官方模型调用成功 |
| 通用会话聚合放行 | 同一生产链路返回真实两项目汇总，boundProjectId=null |
| 绑定 A 请求 B 测算拒绝 | HTTP 403；原真实越权/重启测试继续通过 |
| 通用会话请求任何单项目测算拒绝 | A/B 均 HTTP 403；测试写入也拒绝 |
| 未知无 projectId 工具拒绝 | HTTP authorize 403，未登记会话、无身份也拒绝 |
| 字段边界 | 测试项目路径/note/logs 放置 FORBIDDEN 标记；返回无标记，行字段数量固定14；插件封闭输出 schema |
| limit | 65 条命中，limit=9999，返回50条，并明示“命中65条，仅返回50条”；合计仍覆盖65条 |
| 金额合计 | 65×1000.1=65006.5；65×500.2=32513，按 Decimal 累加 |
| 参数边界 | 客户字面包含、状态、benefit_status、时区等价边界、日期范围、严格小于20%、金额相等范围、未知参数/排序、空或错误区间、limit0/负数 |
| 缺失指标 | IRR `--` 输出 null，`irr > 0` 不命中；百分数字符串20%解析成0.2 |
| 只读 | 测试连接 total_changes 查询前后相同；生产实现仅调用已有 get_projects，无新 SQL 或 JOIN |

## 真实模型防混淆样本

合成数据库中，“当前甲项目”毛利率13%，另一个“其他乙项目1”为87%。
实际模型先按毛利率倒序返回两个项目，甲项目被标为当前绑定项目；含税成本合计为1000.4元。
随后追问“我这个项目毛利率多少”，模型回答当前甲项目13%，并注明已保存摘要、没有重新测算。
测试使用生产 `sessionScopePrompt`，不是测试专用宽松提示词。通用聊天也实际执行查询，结果中两个项目均非绑定项目。

- [实际工具返回与模型回答](./cross-project-query-real-evidence.json)
- [已发给用户的真人核对样本](./cross-project-query-human-review.md)

自动断言验证了13%、甲项目名和不包含87%；这**不能代替任务书第5条要求的真人查看**。
真人核对状态目前保留为待确认，不能将本记录称为已完成所有验收。

## 限制及范围

- 仅读已保存项目级汇总，不重新计算，不返回科目/报价/供应商/模板等明细；过期摘要通过 benefitStatus 明示。
- 复用现有 get_projects，因此后端仍先读取工作区项目行，再过滤和限量返回；未为超大工作区另建索引或分页存储接口。
- 本轮没有新增可视化界面，仅调整已有选择说明、通用聊天状态和模型指导。未重新进行 Windows 或全量桌面人工交互验收。
- 真实测试仅使用合成数据，凭据只进入测试子进程环境；未改正式项目数据。
- 编译保留既有 Rust unused/dead-code 警告与 Vite 大 chunk 提示；无新增依赖。

复验：`agent_bridge/query_tests.rs`、`agent-bridge/scripts/test-project-scope.mjs`、`src-ui/scripts/test_session_scope.cjs`。
