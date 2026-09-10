# AI 多会话工作区与 dsh 产品适配

## 状态所有权

`useAiSessionStore` 是前端会话和消息的唯一真相源，localStorage 快照
`lamber_ai_session_workspace` 保存前端历史。`harnessSessionId` 镜像 Rust 持久化的 ACP id；
它不代表项目绑定或写入授权。前端 `projectId` 只用于显示分组；权限以 Rust 的 `ai-sessions.sqlite` 为准。

2026-09-07 起产品只使用 `DshRuntime`。`AiRuntime`、旧 SSE/think 解析器、临时回退入口与内存状态、旧 endpoint/model/apiKey UI 均已删除。
启动仅清除四个已废弃的 localStorage 配置键，不迁移旧凭据；dsh 设置与聊天历史保持原存储。
没有后端项目绑定的历史保留可读，选择项目后创建新会话；已有绑定与 ACP 映射继续 resume。
`AgentLabView` 保留诊断入口，不是产品聊天的替代实现。

## 轮次与事件

每次发送固定前端 sessionId 和 requestId。Rust `Turns` 在 prompt 排队之前登记关联，
将它们附加到原有 `ai://session-event` 外层；ACP method/params 不变，联调台兼容。
前端先完成监听注册，再 invoke，并且只处理两种 id 都匹配的事件。
终结可先于 invoke 返回；结束 Promise 先建好，末块直接写 store 后才结束，不依赖 React effect 的刷新时机。

`DshRuntime` 接收 ACP 权威消息及插件的实时显示事件，由 `DshMessageProjection` 统一投影。
ACP 文本已完成 UTF-8/JSON 解码，不再解析 SSE 或 `<think>` 标签。
usage/config 与显示通道错误只进有上限的诊断列表，工具按 id 合并稀疏更新。

alpha.5 的 ACP 仍只投影已提交消息。阶段 3A 通过插件订阅 session/event 补齐增量：
模型前的 awaited agent/pre-step 从 `/lamber-bridge/stream` 获得固定请求绑定；
之后的 HTTP 事件携带该绑定，Rust 拒绝将迟到事件重标记到下一轮。
前端按 turn/step/block 缓冲字节，按 messageId 关联提交；提交替换对应步骤预览，
终结时以完整 ACP 输出替换整轮显示。工具参数草稿只能展示，执行结果仍来自 ACP。
预览关联与 decoder 只在单次请求内存中；渲染后的消息沿用原有 store 持久化。

插件串行发送、每请求 5 秒期限、最多积压 2048 事件；异常记录告警并保留 ACP 最终输出。
`npm run test:dsh-stream` 运行真实 dsh 和受控 SSE，再用生产前端适配器回放，
订阅失效、提交前不足三次显示或最终与 ACP 不一致都直接失败。
实现、真实模型数据和接替表见 [阶段 3A 验证](../verification/dsh-stage3a-streaming-and-handover.md)。

## 取消与生命周期

单个聊天面板只允许一轮请求；生成时可切换查看其他会话，输出始终回到发起会话。
`AbortSignal → ai_cancel_prompt → session/cancel` 只取消匹配的轮次，不停整个子进程。
启动期间点击停止会等投递结果后补发 cancel；停止后等待终结事件才释放输入。
清空或删除会话先等待停止，再显式关闭 ACP 会话、清除映射与前端记录。
组件卸载也触发停止并 flush；进程关闭会为仍活跃的轮次发出错误终结，保留已有输出。

## 持久化与工作区

Rust 在 `<app_data_dir>/ai-sessions.sqlite` 建独立映射表：前端 id、ACP id、canonical cwd。
它不修改项目数据库 schema 或 `projects_store.json`。
新会话先保存映射再发 prompt；运行进程只缓存已激活会话。
重启后用 `session/resume` 激活原 id，前端历史仍由 Lamber 展示（dsh 不重放历史）。
恢复失败、元数据冲突或 cwd 不匹配都明确拒绝，不悄悄新建会话。
设置模型会通过标准 ACP config option 同步到恢复的会话，避免旧模型选型覆盖新设置。
未打开工作区的每次发送均明确拒绝；工作区变化时只在旧运行时空闲后重启它。

前端创建、选择、重命名、删除、清空和映射回写立即持久化；消息更新节流 160ms。
图片 base64 不写 localStorage，保留附件元数据，避免挤爆会话存储配额。
删除/清空不清理 dsh 的磁盘历史文件，只解除活动会话与映射；磁盘历史清理不在本次范围。

## 业务上下文与图片

每轮复用 `buildAiChatContext + PromptRenderer`，将系统规则、已保存正式数据、
未保存草稿、加载备注和用户意图通过 ACP text block 发送。当前 ACP 未公开每会话 system/instructions
设置接口；进程级配置不适合多会话动态业务上下文，因此不采用全局补丁注入。

仅显式附加的图片进入 ACP image block；模板图片沿用既有按 projectId/assetId 解析服务。
普通模型界面明确提示不支持图片并引导选择视觉模型，后端再次按握手能力拒绝。
最多四张 PNG/JPEG/WebP，每张 5MB，数据不完整时明确要求重新添加，不改用文字替代图片。

所有业务上下文保持只读。计算、文档、0 容差核验不变；长文本审批的限定扩展见 [审批设计](./ai-approval-review.md)。

## 视觉与验收

宽屏约 216px 侧栏，小于 680px 使用覆盖抽屉。过渡入口、状态、工具明细使用低饱和 surface、
语义字体和现有圆角；设置卡片复用产品设置中心，面板内限制高度并可滚动。

实现/实测/缺失对照见 [阶段 2 验证](../verification/dsh-stage2-product-integration.md)。
阶段 3A 已补齐本地逐 token 显示；用户真人验收、Gate 1A 其余空项和 Windows 1B 仍不得隐去。
阶段 3B 默认切换及临时回退记录见 [3B 验证](../verification/dsh-stage3b-default-transition.md)。

产品浮窗挂载既有审批对话框；审批事件只投递到一个承载窗口，优先 AI 浮窗，否则主窗口。
守卫与失败关闭方向不变；默认审核时间为10分钟，决定可携带修改后文本，审计同时保存两版。窗口挂载不能代替审批授权。

## 项目权限（路线图②）

新会话在已有项目列表中选择项目，或显式选择通用聊天。选择成功、后端持久化完成之后才开放输入。
空白占位会话可完成选择；已有用户消息或 ACP 历史必须另建身份，不导入旧模型上下文。
项目不可改绑；切换页面、改名、前端篡改 `projectId` 均不能改变授权。顶部从后端显示当前绑定项目名。

`ai_bind_session_to_project` 只由显式用户操作调用，检查当前工作区中的项目存在。
绑定包含工作区 id、canonical cwd、可空项目 id，存于同一个 `ai-sessions.sqlite` 的 `project_bindings` 表。
显式 null 表示通用聊天，表中无记录表示未设置权限。③已修订：通用聊天仅允许白名单聚合只读工具；无记录仍拒绝全部工具。
`SessionStore.bind` 拒绝覆盖不同绑定，也拒绝给已有 ACP 映射的历史补绑。
重启后按持久化身份 resume，prompt 前注册到 `ProjectBindings`；清空/删除在同一事务删除绑定和映射，并撤销内存登记。
`Turns` 的 begin/end/late-event 语义完全独立，不清除会话权限。

所有插件工具从 `exec.agent?.session.id` 取得可信身份，缺失直接失败，不发送请求。
`projectScope` 在原审批 waterfall 前请求 `/lamber-bridge/authorize`；额外的单调 guard 拒绝缺失身份和未定义权限的工具，防止其他 listener 提前 allow。
项目类权限保留原有两个工具：测算必须匹配绑定 projectId；无害测试标记也必须有有效项目绑定且仍走原人工审批。
测算业务入口和标记执行体分别再次检查权限，不依赖审批结果来放行跨项目请求。
HTTP 桥接保留原 per-launch token 校验；项目类工具拒绝无绑定、缺项目、跨项目，所有类别都校验工作区与已登记身份并拒绝未知工具。
后端 `WorkspaceRuntime.require_context` 在同一锁顺序中取工作区身份与数据库连接；工作区切换同时更新两者，避免身份与连接配对错误。

dsh 的 `buildAiChatContext` 同样限于绑定项目：仅解析该项目名称；其他页面的草稿不附加；通用聊天不读取项目索引或业务数据。
显式模板图片必须属于绑定项目。`buildAiChatContext` 要求明确的绑定参数；缺失立即拒绝，null 通用聊天不读业务数据。
③已新增唯一聚合白名单 `query_projects`；绑定与通用会话均可调用，只读 projects 行和 summary_metrics，并按默认方案读取阶段/名称/时间元数据。不能借聚合类别调用其他业务工具。细节见 [ai-project-query.md](./ai-project-query.md)。

验收：[session-project-binding.md](../verification/session-project-binding.md)。

## 阶段 B 模板文本工具（2026-09-07）

fill_template_fields 只允许可信 ACP 会话的绑定项目；通用会话仍仅可聚合只读。审批前读取实际旧值，执行前再次复核工作区/绑定与模板版本，消费用户批准的字段文本。模板正文广播按目标定位并合并，不能以当前界面项目替代会话绑定。详见 [ai-template-write.md](./ai-template-write.md)。

### 主动模板读取

read_template_fields只接受templateId，项目由ACP绑定取得。新增工具沿现有项目范围授权，执行端再次核对；通用聊天与额外projectId拒绝。bash/glob继续禁止。该主动查询不依赖模板页面，也不扩大buildAiChatContext的被动注入范围。

## 浮窗生命周期恢复（2026-09-07）

AI入口通过独立窗口服务串行创建/复用，原生created后才定位显示；恢复最小化、核对当前显示器工作区并修复离屏位置。保持一个ai-assistant窗口，不销毁或重置会话来处理位置问题。位置偏好为v2物理像素，历史逻辑像素在恢复时转换并校验；失败反馈留在主窗口入口旁。详见[位置恢复验证](../verification/ai-window-recovery.md)。
