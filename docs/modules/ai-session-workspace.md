# AI 官方 WebUI 与 Lamber 业务适配

## 产品入口与状态归属

正常 AI 按钮经 `aiAssistantWindow.ts` 调用 `ai_open_webui`。Rust 定位分发资源，创建当前工作区的回环 Host，打开同源官方顶层 Webview。聊天、输入、流、工具过程、附件、会话日志和列表由完整 alpha.5 WebUI 管理；Lamber 通过官方 slots 和主题扩展提供绑定、业务入口、审批、旧历史和设置。实施及未完成验收见[升级验证](../verification/ai-webui-upgrade.md)与[任务书](../tasks/TASK_BOOK_ai_webui_upgrade.md)。

活动聊天不再双写 localStorage。`AiChatPanel`、`AiFloatingWindow`、旧侧栏/输入框/消息组件、`DshRuntime`、`DshMessageProjection` 和 `useAiSessionStore` 已退出并删除。Rust `ai-sessions.sqlite` 只保存可信绑定、必要映射、所选会话、迁移原文和业务动作回执。原 localStorage 源不会被清除。`sessionTypes.ts` 保留历史格式读取契约；AgentLab 及 ACP 后端仅供诊断，不是另一套产品聊天。

## 官方插件装配契约

`prepare_web_home_at` 生成 web profile 和部署补丁，资源版本由锁文件及打包检查固定为 alpha.5。官方 `applyEntryPatches` 的 `name` 是匹配断言，不能用它修改已有行的插件名称；不匹配会跳过。替换必须停用原行并插入不同 id 的部署行。

Gateway、SessionController、WorkspaceController 同时拥有 Host 和 browser 两端；只停用原行会删掉原浏览器模块。部署层在独立 `lamber-client-faces` 目录生成入口：Host 转发到 Lamber 子类，客户端保留原包身份、原 `dsh.client` 元数据和逐字节相同的官方 `client.js`。不修改已安装上游包，不维护客户端分叉。品牌包的 `inject` 只声明依赖，并不会自行激活被停用的模块。

Host 策略启动时检查四个实际服务的部署标记；集成测试启动真实官方 Host，核对返回页面中的三个原浏览器模块及 conversation，并比对客户端文件字节。不能用适配器单元测试代替实际组合验证。

部署子类须声明新增 Cordis 服务依赖。SRC Remote 从函数签名推导命名参数，必须保持官方 `_request` 等参数名；`inspect` 返回 `meta`，历史 `follow` 接受 `request.address`。这些契约已有回归测试，升级上游时须重新核对。

## 身份、权限与原业务服务

Web 会话以 `web:<dsh id>` 接入主窗口服务。绑定包含工作区 id、canonical cwd、可空项目 id，Rust `project_bindings` 为唯一授权来源。显式 null 表示通用聊天，无记录表示尚未选择权限。已有项目绑定不可改变；只有后端可信旧映射可恢复原身份。每次模型准入、工具执行和业务提交分别复核工作区及绑定。

Session 的 list/search/inspect/follow/control 和 Workspace 的 follow/归档/排序限定当前工作区；创建不能选择其他目录或其他预设。Gateway 通过公开 `invoke/stream` 限制管理、目录、凭据等接口；不另注册第二个 `/api` 所有者。官方设置描述只公开界面主题/语言，配置文件入口由 `hasDocument: false` 正常隐藏。隐藏按钮不能替代后端限制。

工具使用 `exec.agent.session.id` 的可信身份，沿原白名单、二维项目/通用权限及审批链执行。通用会话仅开放当前原白名单的聚合读取和甄选费计算；未知工具、跨项目和失效身份拒绝。详见[项目查询](./ai-project-query.md)、[模板文本](./ai-template-write.md)、[测算 A/B/C 与 D](./ai-benefit-calculation.md)。不开放 SQL、通用原生 invoke、直接项目金额或计划写入。

业务卡片通过固定高层操作交给主窗口原服务执行。原 Word 生成、技术/询价清单、需求表附件事务和 D 编辑器变换保持单一实现。D 确认只更改原编辑器，正式保存必须由用户点击原保存按钮；财务公式、原目标、容差、28 科目及全部年度计划规则不变。

## 单一审批与动作回执

官方 `tools/pre-execute` 前置钩子直接接原 ApprovalGate 和文本事务。Web 复用 `ApprovalReview` 的新旧对照及改文；Rust 审批仍拥有唯一决定、绝对截止时间、两版审计、版本复核及一次性许可。官方默认审批不能再独立放行同一写入。

停止绑定真实工具 AbortSignal，早到取消按 session/call 登记。停止或关闭承载窗口拒绝未完成审批；重开只读原结果，不恢复批准按钮。原生 AgentApprovalDialog 留给主窗口/诊断台。缺窗口、超时、冲突及重复确认维持失败拒绝。

SQLite 动作账本先登记，主窗口原子领取一次，再保存最终回执。未领取请求 15 秒到期；运行中操作不自动重试。异常退出只记结果不确定并提示核对原页面/文件。历史预览剥离令牌；临时读取结果消费后删除，图片字节不写回执。回执描述操作发生时的结果，后续状态以原页面为准。

## 模型上下文与设置

每轮 Host 准入后请求主窗口原 `buildAiChatContext` 序列化投影；只注入绑定项目的正式保存态、匹配的未保存草稿及加载备注，通用聊天不被动读取项目数据。当前用户意图与原业务知识/卡片邀请复用原纯函数。上下文及历史回执是被动数据，不能提升权限或触发动作重放。

alpha.5 在 `agent/pre-step` 前已完成 systemPrompt 组装；此时才登记 context provider 不会进入本轮。部署使用官方 `createUserMessage` 插件快照加入本轮 `decision.messages`，保留当前用户消息在后；旧插件快照不能重新作为用户意图。上下文读取受本轮取消和 30 秒截止控制。

模型/key/baseURL 归 Lamber config，密钥不回显、不入聊天或回执。设置页明确“新会话默认模型”；保存连接设置后重启 Host，保留原会话选择。已有会话保留自己的模型；官方输入框选型同时通过部署 `agentDefaultModel` 更新 Lamber 默认，不产生第二套设置持久源。不能因允许 baseURL 就宣称支持任意兼容供应商。

## 历史迁移与窗口生命周期

旧记录先原文备份，再校验/索引；格式损坏保留原文和上次有效索引，返回可见警告，不阻止新界面使用。只有 Rust 原会话映射和不可改绑记录同时成立，才允许恢复同一个原 dsh id；其他历史只读。图片失效明确显示元数据，不能把未附加图片伪装为已进入模型。迁移幂等，重复读取或恢复不重放旧动作。

所选 dsh 会话按桌面工作区持久化，不依赖每次随机端口的浏览器存储。关闭窗口回收 Host；父进程探针处理异常桌面退出。启动失败在主窗口入口旁提供可读错误/重试。入口使用标准按钮 click，拖拽只更新几何并抑制对应 click，支持键盘及辅助功能。

Tauri Context 只在 main 生成一次并移交产品或 debug 探针；重复 `generate_context!()` 会在 macOS dev 导致 `_EMBED_INFO_PLIST` 重定义。启动变更须实跑 `npm run tauri dev`，默认构建不能代替。探针只用于独立诊断，正式验收必须从普通 AI 按钮进行。

窗口位置沿用 v2 物理像素及显示器可见区域校验；保持单一 `ai-assistant` 窗口，不靠重建会话修复位置。品牌/业务层遵循[外观规范](./appearance.md)。Windows 安装、真实供应商及用户真人验收与本机工程验证分别记录。

## 原生工作区切换与审批归属

切换/关闭工作区、移除当前工作区均与 Web Host 创建共用原生互斥边界。在数据库身份切换前拒绝挂起审批、回收旧进程并隐藏旧页面；下一次普通 AI 点击加载新工作区的 Host 和选中会话，不等待前端轮询或子进程心跳。

每个 Web Host 固定原工作区数据库和专属 `.lamber-ai-approval-spool.jsonl` 故障缓冲；每个审批请求在开始时捕获记录器；切换时解除运行时对旧记录器的引用，挂起请求结束后即可释放旧数据库句柄。即使新 Host 已启动，旧回调也只向原数据库记审计。专属缓冲仅在同工作区打开时重放，不落到任意下一工作区。通用的无工作区诊断审批保留原应用级缓冲规则。
