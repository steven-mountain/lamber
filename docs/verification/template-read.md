# read_template_fields 补充验证（2026-09-07）

执行依据：TASK_BOOK_ai_write_template_fields.md **末尾“补充（2026-09-07）”**。本轮补齐主动查询，原任务书 A/B 真人 gate 未冒充通过。

## 已实现

- `read_template_fields(templateId)`，只接收模板名。项目从可信 ACP 会话绑定取得；服务端拒绝额外 projectId/fields/path，通用聊天、缺失或失效绑定拒绝。
- 复用现有 get/list_template_state 和资产读取函数，不新增 SQL、表、迁移或写链路。读取不创建审批记录，不改变模板版本。
- 需求导入表 8 个文本值、技术清单、两个附件槽位存在状态及条件开关。valueSource 明确 saved/default/empty，不把界面默认值称为用户填写。
- 插件构建直接编译现有 demand.ts 纯函数及同一 JSON 目录，返回完整11项完成度与 missingFields，没有第二套缺项算法。为避免截断造成误判，缺项函数接收完整保存态的存在标志。
- 文本总计上限24000个 Unicode 字符，技术清单最多100行；返回原字数/实际字数、字段与总体 truncated、实际/返回行数和提示。未超限长文保留换行及全文。
- 白名单 DTO 不包含模板路径、输出目录、字段映射、图片路径或文件名；公示网址作为业务字段保留。其他模板不在当前需求表规则支持范围。
- 可用“需求导入表”定位绑定项目唯一已保存需求表；多候选要求指定名称，无候选不猜测。精确模板名无保存态时明确 hasSavedState=false 并标识默认值。
- 不改 buildAiChatContext 页面注入范围，不放行 bash/glob，不登记审批白名单。桥接契约同步升至 v4。

## 自动验证

1. Rust 常规 **88通过 / 19 ignored**：A/B 同名模板隔离，B独有模板不能泄露到A；额外参数/通用聊天/未知身份/项目删除/路径模板拒绝；未保存状态区分；原文换行保持；内部路径及任意额外字段排除；读取前后 total_changes 不变；Unicode超长截断仍保留完整缺项依据。
2. 插件端使用上述真实 Rust DTO 跑生产 execute，普通和超长正文均输出8/11及相同三项缺项；核对生成的完成度函数正文与 UI 原函数一致。
3. **19 项真实 Key ignored 集成通过**，包含②权限、③聚合及字段边界、审批原路径、阶段B写入和新增无页面上下文主动查询。
4. 前端 lint/build 与 dsh/template-state/session-scope/approval-review 四组回归通过；插件 typecheck/build/限权/契约通过；打包6项通过；独立 macOS debug app 构建及实际二进制契约 v4 校验通过。

证据：[Rust投影及截断](./template-read-projection-evidence.json)、[真实模型](./template-read-real-evidence.json)。投影证据为合成数据，包含长文截断样本。

## 原始场景真实复验

**真实 dsh 新会话，无页面注入**：只发送“需求导入表填了什么？还有哪些没填？”。模型实际读到只存于合成 SQLite 的“青鹭472”正文，区分默认与已保存字段，报告8/11及部署环境/附件1/附件2三项缺项；审批次数0。

**Computer Use 独立桌面应用**：重启 Lamber Scope Test，从项目看板直接打开 AI 面板（view=project_board），新建空白会话并绑定“3A流式验收项目”，期间没有进入模板页。发送同一句裸问，界面出现 `read_template_fields · 已完成`。模型正确返回模板版本12、337字服务正文概述、含“联合勘察并书面确认”的部署环境、7项技术清单及两附件已上传，完成度11/11。展开工具输入仅有 templateId，无 projectId。与先前模板页保存态一致。

[桌面复验记录](./template-read-ui-evidence.json)。测试使用合成工作区，未更换正式安装应用。

## 待真人验收

读工具已补齐，可恢复任务书要求的真实辅助填表 Gate A/B：不开模板页先询问现状，再让AI补缺项，核对审批旧值并改一处，批准后确认完成度+1及DOCX改文。Computer Use 和自动断言不替代真人可读性判断。Windows NSIS 安装及其他发版 gate 仍待完成。
