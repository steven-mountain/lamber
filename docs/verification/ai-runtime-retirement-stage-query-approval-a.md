# 旧 AI 下线、查询阶段标签与审批 A 验证

日期：2026-09-07。执行者：Codex 自动化 + Computer Use。独立应用 Lamber Scope Test，独立配置/合成工作区；没有修改真实项目数据或替换 /Applications 的正式应用。

## 已实现

1. 删除 AiRuntime、useStreamingParser、chatTransition、旧配置 UI、临时回退入口/状态与无调用 telemetry/isolated 占位。仅保留 dsh 产品聊天。
2. query_projects 按默认方案补阶段、名称、保存时间；缺失标为未标注，混合阶段不提供统一总额，按阶段合计和财务排序。既有汇总保存语义不变。
3. 审批意图、新旧对照、编辑后批准、双版本审计、10分钟失败关闭已实现。ACP grant/reject 不变，新增同一决定的单次参数交接，契约升至2。
4. 保留联调台并提供独立文件长文本演练，供 Gate A 查看覆盖、编辑和实际文件回执。

## 3A 接替表复核

| 旧功能/入口 | 下线后的接替与验证 |
| --- | --- |
| 普通发送与快捷动作 | AiChatPanel.handleSend → 上下文 + PromptRenderer → DshRuntime；快捷按钮仍调用同一 handleSend；lint/build、前端 dsh 测试通过 |
| 上下文与历史 | boundProjectId 必填；通用 null 不读业务，缺失失败关闭；已绑定历史 resume，无绑定旧历史保留并另建会话；session-scope/历史持久化测试、真实重启恢复通过 |
| 模板图片间接入口 | TemplateForms.handleSendImageToAi → templateAssetSelection 的 DOM/storage/Tauri 事件 → AiChatPanel 附件 → loadAiTemplateAsset → dsh image block，代码链路保留；资产校验/前端图片回归、真实视觉模型识红图通过 |
| 停止/卸载/清空删除 | AbortController → ai_cancel_prompt → 等待轮次结束 → reset；前端取消竞态与会话隔离、真实停止续聊通过（本轮取消确认10ms） |
| 流式正文/思考 | DshMessageProjection；真实流1451个增量、1451次提交前更新，首delta1.760s/首commit13.244s，最终ACP一致 |
| endpoint/model/apiKey | 唯一 AiAgentSettingsCard；启动只删除旧四键，dsh配置及历史不动 |
| getTraces/clearTraces、invokeToolIsolated | 随旧实现删除；无外部调用，保留现有 dsh 有界 diagnostics 和联调台 |
| 临时回退入口本身 | 设置按钮、执行分支和内存回退状态同时删除；源码检索与编译无残留引用 |

没有把旧版本试用记录改写成这次重跑。本轮未再次完整点击“模板页图片→聊天发送”间接链路，已复核代码接线、相关确定性测试与 dsh 真图片终点；先前链路人工/Computer Use 证据保留在3A与模板资产记录中。

## 自动化结果

- Rust `cargo test --no-default-features`：86 passed，17 ignored，0 failed。
- 真实 key `cargo test agent_bridge -- --ignored --test-threads=1`：17 passed，0 failed/ignored；凭据仅进子进程环境。
- 真实审批批准/修改后批准/拒绝均经过 dsh + ACP + Rust gate + 插件实际执行，修改分支读取标记文件并断言保存用户值而非模型原文。
- 新增确定性测试：编辑禁止额外字段/非文本；单次交接拒绝错会话/错调用/错工具/错原参/重复/过期/撤销绑定；审计两版和旧值；长文演练拒绝不写、批准按修改值保存。
- 阶段测试：默认指针切换前/后、缺失指针、未知阶段、未标注不猜测；混合前/后阶段limit=1仍按全部命中分组汇总；金额与summary_metrics未重算。
- 前端 lint/build、test:dsh、test:session-scope、test:template-state、test:approval-review 通过。
- 插件 typecheck（构建TS检查）、test:scope、test:contract 通过；打包6项通过；独立 macOS .app 构建通过。

## 真实模型口径复验

先查两项目：当前甲13%、甄选前；其他乙87%、未标注。模型显式提醒混合口径并给分组合计。再裸问“我这个项目毛利率多少？”，无额外引导：

> 我这个项目是**当前甲项目**（绑定项目）：已保存毛利率为 **13%（0.13）**，属**甄选前（限价口径）**，对应方案“甄选前”，保存时间 2026-01-01。该值为已保存摘要，如需最新测算可告知。

这是合成项目的真实模型输出；[完整实际证据](./cross-project-query-real-evidence.json)。真人核对状态仍 pending。

## 审批界面验证

Computer Use 在独立测试应用实际操作：

1. 389字多段样例打开审批，标题/项目/表单/位置、空白标记可见；正文不转义换行；底部10分钟审核与拒绝/批准按钮可见。
2. 批准一次，再打开显示“覆盖已有内容”和完整原文。
3. 将“并保留回退方案”改为“并在切换前完成回退演练、保留原网络”，按钮变“修改后批准”。
4. 点击后实际文件回执为399字修订正文；展开持久化审计，模型原文、用户修改值和最终批准值均可查看。
5. 从独立工作区只读核验 args_json 与磁盘文本一致。详见 [审计与文件证据](./approval-a-review-evidence.json)。

## 明确未完成

- **Gate A 真人验收仍待用户**：任务书要求“300字以上真实需求分析文本”，本轮合成样例与自动化点击不代替真人判断。独立应用已打开联调台供用户操作。
- **阶段 B 未开始**：没有 fill_template_fields，也没有任何模板或财务字段写工具。
- 查询③真人核对、Windows NSIS Gate1B、既有密钥设置与发行剩余 gate 不冒充已通过。
- 无正式发版、安装覆盖或对外发布。

真人步骤：在独立应用展开“长文本审批演练”，换入真实正文→打开审批→核对覆盖/填空→改一处并批准→核对实际文件回执及右侧日志两版正文；通过后再进入阶段B。
