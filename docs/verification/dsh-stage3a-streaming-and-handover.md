> 2026-09-06回填：本记录发现的模板确认页字段回退已修复，四模板跨页产物XML对照通过；见[验证记录](./template-state-and-chat-assets.md)。

# dsh 阶段 3A：实时显示与旧运行时接替表

日期：2026-09-05。执行者：Codex（自动化测试与 Computer Use），不是用户真人验收。
本轮范围为 3A-0 / 3A；产品开关默认仍关闭，未删除 `AiRuntime.ts`，未进行 3B 或发版。

## 订阅前置验证

先在临时复制的 `dsh-tool-lamber.apply(ctx)` 加只输出测试事件的观察器，运行真实
`@deepseek-ai/dsh@0.1.2-alpha.5` ACP profile，连接受控 SSE 端点。
收到 8 条 `assistant/chunk`，包含 `turn / step / chunk`；类型包含 block-start、
reasoning-delta、text-delta、block-end、finish；第一条早于 assistant/message。
验证通过后才接入产品路由。没有轮询会话文件，没有修改 dsh 源码。

## 修复范围与双通道契约

根因是 ACP alpha.5 仅投影提交后的 assistant/message，实时 delta 已存在于 session/event。
单纯把 delta 追加到原 reducer 会把正文重复一次，且第二个工具步骤与第一段回答混在一起。
因此保留 ACP 作为权威来源，新增有边界的预览投影：

- 插件 `stream.ts`：公开 `agent/pre-step` awaited hook 在模型执行前绑定当前前端
  sessionId/requestId；`session/event` 转发 text/reasoning/tool-call 增量及提交的
  messageId ↔ turn/step 关联。不插入 prompt 标记，不改变工具或审批规则。
- `LAMBER_STREAM_DISPLAY=1` 由产品运行时设置。独立插件/旧联调测试默认不启动显示桥接；
  订阅契约测试显式启用，不能以未启用模式冒充流式验证。
- `/lamber-bridge/stream` 复用 loopback 令牌鉴权、1 MiB 请求上限和 16 worker 上限；
  byte 数组、类型、序号进行校验。只发显示事件，不写项目数据。
- HTTP 请求顺序发送，单请求 5 秒期限，积压上限 2048；故障保留 ACP 权威输出，并记录
 插件告警及有上限的前端诊断。模型前绑定失败直接使请求报错，不默默绑定其他会话。
- 每条增量携带最初获得的绑定；Rust 在锁内检查仍与活动轮次相符，旧片段绝不重标记到新请求。
- 前端 `DshMessageProjection` 按 turn/step/block 暂存增量，按 messageId 关联 ACP。
  提交替换该步骤预览；终结事件将整轮显示替换为 ACP 累积正文/思考/工具结果，
  因此即使 HTTP 提交标记迟到，也不会重复、截断或留下未执行的工具草稿。
- 插件对尾部高代理项缓冲后编码字节；前端 `TextDecoder(..., {stream:true})` 缓冲不完整 UTF-8。
  LLM delta 本身是字符串；SSE 原始字节由 dsh 解码，桥接不再次解析 SSE。

## 可重复的验证

`npm run test:dsh-stream` 编译插件，启动临时 home/workspace 的真实 dsh + 受控 SSE，
验证至少三条增量早于 ACP commit，并把 HTTP/ACP 混合实录交给生产 DshRuntime 回放。
订阅失效、只有提交后更新、最终内容不一致均直接失败，不依赖 key，不允许“跳过即通过”。

真实模型命令（key 只在进程环境中使用；配置路径显式指定，不记录内容）：

```sh
python3 scripts/verify-dsh-stage3a.py --real --app-config /path/to/config.json
```

| 验证 | 结果 |
| --- | --- |
| 受控 SSE | 6 个增量；首 delta 0.781s，首 ACP commit 1.685s；6 次提交前显示；最终一致 |
| DeepSeek-V4-Flash 真实模型 | 590 个增量；首 delta 1.556s，首 ACP commit 7.865s；生产适配器 587 次提交前显示；最终正文/思考逐字等于 ACP |
| 中文与 emoji | SSE 字节流逐字节发送；前端遍历中文/emoji/金额字符串所有 UTF-8 切分位置，无乱码 |
| 跨通道竞态 | ACP 先到 / HTTP 先到、同消息多个块、多工具步骤、重复/迟到片段、取消时最终前缀替换通过 |
| Rust 轮次隔离 | 模型前绑定、无活动轮次拒绝、旧请求片段不能进入新轮次、非法 byte 拒绝通过 |
| Rust 常规 | 78 passed，13 ignored，0 failed |
| Rust ignored（带真实 key） | 13 passed，0 ignored，0 failed；含原审批、取消续聊、恢复、视觉模型与分发握手 |
| 前端 / 插件 | test:dsh、lint、build、插件 typecheck 通过；打包脚本 5 项测试通过 |

说明：原 13 项 ignored 回归未开启新的显示桥接；流式本身由上面的独立契约测试和产品试用验证。
真实模型实录是纯合成的“整理桌面文件”问答，生成在系统临时目录，不包含客户数据或密钥。

## AiRuntime 全部接替关系（3B 删除前置）

搜索 `src-ui/src` 中 AiRuntime 的 import/new/execute，以及其公开 trace/tool 方法：
只有 `AiChatPanel` 一处直接导入、一个实例、一个 execute 调用；没有另一个隐藏模型调用方。
面板由 `AiFloatingWindow` 挂载。以下是同一发送入口背后的不同业务链路。

| 原调用/能力 | dsh 接替位置 | 3B 注意事项 |
| --- | --- | --- |
| 文本输入、Ctrl+Enter、发送按钮 | AiChatPanel.handleSend → DshRuntime.execute → ai_send_prompt | 默认值本轮未改 |
| 快捷“分析当前项目效益”“推荐合适产品” | 同一 handleSend(action.label)，共享 dsh 分支 | 不另建 prompt 或运行时 |
| 当前页面、指定项目、已保存/草稿上下文 | buildAiChatContext → PromptRenderer → ACP text block | 继续只读、区分来源、元单位规则不删 |
| 上传图片 | resolveImagesForSend → imageBlocks → Rust prompt::blocks | 普通模型拒绝；视觉模型能力来自设置和握手 |
| 模板图片间接依赖 | TemplateForms.handleSendImageToAi → templateAssetSelection 的 DOM/storage/Tauri 事件 → AiChatPanel 附件 → loadAiTemplateAsset → dsh image block | 保留 projectId/templateId/assetId 校验和显式点击；不能只搜 AiRuntime import 就删附件路径 |
| SSE / think-tag parser | DshRuntime + DshMessageProjection；HTTP 字节预览 + ACP 正文/思考提交 | useStreamingParser 只属于旧链路；不能用它解析 dsh |
| 中断 / 关闭浮窗 / 清空删除 | AbortSignal → ai_cancel_prompt；清空删除先等待结束，再 ai_reset_session | requestId 防止旧取消误伤新轮次 |
| 本机多会话历史 | useAiSessionStore；Rust session_store；session/resume | 原 4000-token 前端裁剪不搬到 dsh；上下文历史预算改由 harness 管理，恢复不重放 UI 历史 |
| endpoint/model/apiKey UI | AiAgentSettingsCard → ai_get_settings/ai_save_settings → 启动配置 | 不宣称任意 OpenAI 服务兼容；不迁移不兼容的旧服务凭据 |
| getTraces / clearTraces | 无外部调用；dsh 有界 diagnostics + AgentLabView | 无需为了空调用方再造一套 trace 服务 |
| invokeToolIsolated | 旧类内占位，检索无调用方；实际工具由插件 + Rust bridge | 删除占位不会损失真实工具；工具数和审批策略不扩展 |

`AgentLabView` 保留为排障入口（设置页打开），不是产品路径，也不是 3B 删除旧运行时的接替者。

## 产品工作流试用

独立 debug `.app`：`com.cmcc.benefitcalc.stage3a`；工作区
`/private/tmp/lamber-stage3a-ui-workspace`；项目“3A流式验收项目”，客户“合成测试客户”。
所有业务录入、保存、开关与文档生成由 Computer Use 点击产品界面完成。
测试进程使用用户已配置的 key 作 debug 环境注入，不复制密钥到测试配置文件。

已录入系统集成收入含税 106,000 元（6%，不含税 100,000），设备投入含税 67,800 元
（13%，不含税 60,000），首年收付；科目计划一致、零容差检查通过后进入十年现金流表。
UI 显示 NPV 37,914.69 元、毛利率 40%、净现值率 66.67%，并保存方案。

手动打开 AI 入口及 dsh 开关后，真实模型思考内容在“停止生成”仍可操作时持续增长；
点击停止后可再次发送。指定项目续聊返回 NPV 37,914.69 元、毛利率 40%，没有把元误标为万元。
当上下文摘要没有提供完整收入/投入时，回答明确说缺少明细，没有猜数字。

只读 `run_benefit_calculation` 实际完成，工具调用卡片为“已完成”，返回上述 NPV/毛利率，
以及首年流入 100,000 元、流出 60,000 元、净现金流 40,000 元；工具前后回答没有重复。

需求导入表录入需求单位“合成测试分公司”和合成服务说明后保存；从“需求信息”页直接点击
“立即生成文件”，产物位于工作区项目目录：
`3A/ICT项目需求导入表-3A流式验收项目.docx`。
ZIP 完整性检查通过，提取 Word XML 的文本后确认项目名、需求单位、服务说明、106000 元预算均正确。
SHA-256：`9c874b7a3ebb8f90d6eb449e18a0e18ae09684cb85aa13f5578a4bda42d123ec`。
这是产品生成路径验证，不是文书排版真人验收。

### 试用暴露的既有模板问题（未改，不能记为通过）

同一已保存表单，切到“生成确认”页再生成，需求单位退回“XXX分公司”、服务内容退回默认组合，
而项目名和预算正常。已保存 SQLite formData 中自定义字段仍在，回到“需求信息”页可看到正确值；
在该页直接生成又能正确填充。

失败路径：`TemplateForms.handleGenerate` 使用 `new FormData(formRef.current)` 取当前挂载控件；
确认页未挂载的需求字段读到空值，随后使用默认值，未从已保存的表单状态补全。
这与 dsh 无关，本轮没有修改 TemplateForms 或文档引擎。留样：
`3A/ICT项目需求导入表-3A流式验收项目-生成确认页留样.docx`。
建议后续以统一表单状态构造生成输入，不能仅给两个字段增加默认值分支。

3A 试用结论：流式思考、停止续聊、真实工具与需求信息页文档生成已验证；
生成确认页的字段问题单独待修。不能据本轮宣称全部文档入口已通过。

## 边界

本记录不替代用户真人验收或 Windows NSIS Gate 1B。二者仍阻塞真实发版。
不作“用户已接受非流式降级”的结论：本轮验证的是自行补齐流式后的体验。
上游 issue 草稿保留，未代用户向外发布；上游未来支持实时投影后可移除本地显示 seam。
