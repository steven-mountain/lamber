# 任务书：需求分析对话 · 缺项校验 + 聊天内图片补齐落库

> **2026-09-05路线图已解冻方案A；2026-09-06实现及核心界面验证完成。**
> 与模板FormData任务合并执行，不依赖会话绑定；方案B仍不在范围内。
> [验证记录](../verification/template-state-and-chat-assets.md)：真实模型口述与图片粘贴端到端未验收，禁止标成全量验收完成。

## 产品目标（一句话）

用户在聊天框里说"帮我做一下这个项目的需求分析"，AI 读出《ICT项目需求导入表》当前还缺什么；
缺的是图片时给出上传槽位，用户点上传，图片落到**该项目 / 该模板 / 该 usage** 的资产库和项目文件夹里，
回到模板页立刻可见、生成 docx 时直接被用上。

---

## 已确认的关键事实（写代码前不用重新调研）

### 1. 缺项校验规则已经存在，但只活在 React 渲染期

`demandCompletionItems`（`src-ui/src/views/TemplateForms.tsx:2028-2040`）是组件渲染期的一个局部数组，
11 项，依赖十几个组件 state（`attach1Images`、`attach2Images`、`hasPublicUrl`、`techItems`、`getFormValue(...)`）。
后端和 AI 上下文都拿不到它。其中最后两项正是本次要用的图片缺项：

```
{ label: "附件1客户确认材料", filled: attach1Images.length > 0 },
{ label: "附件2招标材料",     filled: !hasPublicUrl || attach2Images.length > 0 },
```

**要让 AI 说出"缺附件1"，必须先把这套规则抽成不依赖 React 渲染的纯函数，不能在 AI 侧另写一套判断。**
另写一套的结果一定是聊天里说的缺项和模板页进度条对不上，用户会以两者中较严的那个为准，等于白做。

### 2. 图片入库链路是完整可用的，不用新建

`save_template_asset`（`src-tauri/src/project_files/commands.rs:172`）→
`save_template_asset_internal`（`src-tauri/src/project_files/assets.rs:98`），已经做完了：

- 按 `(project_id, template_name, asset_type, usage)` 写 `project_template_assets` 表；
- 物理文件写到 `<项目文件夹>/<项目名>-图片/assets/`；项目没绑文件夹、或路径落在工作区之外时，退到
  `<workspace>/.projects/<project_id>/assets/`；
- 写库失败会把已落盘的物理文件删掉再报错（`assets.rs` 尾部 `match result` 分支），不留孤儿文件；
- 限制：≤ 20MB（`assets.rs:113`），仅 PNG / JPEG / WEBP。

需求导入表现在用的 usage 就是字符串 `"attach1"` / `"attach2"`（`TemplateForms.tsx:835-836`、`3390`、`3433`）。

**"把图片存到数据库中项目对应的位置"这件事本身已经做完了，本次缺的是"从聊天里触发它"。**

### 3. 目前图片方向是单向的：模板页 → AI

`handleSendImageToAi`（`TemplateForms.tsx:1260`）通过 `publishTemplateAssetSelection` 事件，把**已入库**的图片
推给聊天面板做分析（`AiChatPanel.tsx:179-230` 监听 `AI_TEMPLATE_ASSET_SELECTED_EVENT`）。
反向——聊天里要图、图落库——没有任何链路。本次要补的就是反向这一段。

### 4. AI 已经看得见"有哪些图"，看不见"按业务规则还缺哪些图"

`AiTemplateDetailContext.assets`（`src-tauri/src/ai_context/dto.rs:142`，元素含 `assetId / fieldKey / exists`）
已经由 `buildAiChatContext.ts:97-113` 注入上下文。所以"已有资产清单"是现成的；
缺的是第 1 条那套"按模板业务规则该有而没有"的判定。

### 5. 聊天框里随手拖的图片不入库，刷新即丢

`AiInputBox`（`src-ui/src/components/ai/AiInputBox.tsx:63-98`）只做 `fileToDataUrl`，
上限 ≤ 5MB、≤ 4 张，且 `useAiSessionStore` 持久化时明确剥掉大体积 base64、只留元数据
（`src-ui/src/store/useAiSessionStore.ts:173-176` 及其注释）。

**即：用户现在在聊天框里贴一张图 = 图进模型、不进库、刷新就没。**这正是本次要补的那一段，
也是"不要把所有聊天图片都自动入库"这条边界的由来（见"不要做的事"）。

### 6. 两条聊天链路的图片能力不一样，开工前必须先把这条查清

- **Chat 模式**（`src-ui/src/ai/AiRuntime.ts:172-186`）自己拼 OpenAI 的 `image_url` 块，
  能不能收图取决于用户在设置里配的 endpoint / model。
- **Cowork / dsh 模式（ACP）能不能收图，取决于配置的是哪个模型**：`dsh-acp` 的 `initialize` 返回
  `agentCapabilities.promptCapabilities.image`，它由 `supportsAcpImagePrompts()` 算出
  （`agent-bridge/node_modules/@deepseek-ai/dsh-acp/lib/index.js:78-88`），要求配置模型的
  `inputModalities` 含 `"image"`。不满足时，任何带 image block 的 prompt 会被
  `admitAcpPrompt` 直接拒掉（同文件 `110-112`）。
- dsh 自带的默认模型目录里**有**视觉模型：`deepseek-v4-flash-vision-exp`，
  `inputModalities: ["text", "image"]`（`@deepseek-ai/dsh-llm-deepseek/lib/index.js:1839-1846`）。
  lamber 现在硬编码的 `deepseek-v4-flash`（`dsh_session.rs:130`）没声明 image——
  **所以这是模型选型问题，不是 dsh 缺能力。** 详见
  `docs/tasks/TASK_BOOK_dsh_full_integration.md` 的"关键事实 5"，模型可配也在那份任务书里做。
- Rust 侧现在把握手答案丢掉了：`AgentHandshake::from_response`（`dsh_session.rs:182-190`）只留了
  `protocol_version` 和 `agentInfo`，没读 `agent_capabilities`。

### 7. 审批机制不用碰

图片是用户点出来的，不是模型写的库。AGENTS.md 的 "No direct AI database writes" 要求写库必须由用户动作触发，
本方案天然满足，`GATED_TOOLS` 不需要新增条目，审批弹窗、`agent_approval_log` 全部不动。

---

## 要做的事

### 0. 先确认"关键事实 6"那个前提

确认当前配置的模型声明了图片输入（`promptCapabilities.image` 为 `true`）。
若为 `false`，说明配的还是非视觉模型——去 `TASK_BOOK_dsh_full_integration.md` 阶段 1 把模型改成可配、
选上视觉模型即可，**不是本任务的阻塞项，也不要停工**。

注意：本次第 1-5 步（缺项校验 + 上传落库）**完全不依赖模型看不看得懂图**，
即使模型是纯文本的也照常做得完，只是 AI 无法对图片内容发表意见。

### 1. 把需求表缺项规则抽成纯函数

新增 `src-ui/src/lib/templateCompletion/demand.ts`（目录名可调整，但必须是可被 AI 上下文侧复用的位置）：

- 输入：序列化后的模板表单状态 + 资产列表（不要接 React state，不要接 `formRef`）。
- 输出：`Array<{ key, label, filled, kind: 'field' | 'image', templateName, usage? }>`。
- `kind: 'image'` 的项**必须**带 `templateName` 和 `usage`，否则第 3 步不知道该把图存到哪个槽位。
- `TemplateForms.tsx` 改成调用它来渲染现有进度条。**权威规则只有这一份。**

### 2. 把缺项清单接进 AI 上下文

在 `buildAiChatContext.ts` 现有 `templateDetail` 分支（`:97-113`）里加一段 `completion`，用第 1 步的函数算出来，
字段级缺项和图片级缺项分开列。

这样"发现缺图"是**确定性的规则判定**，模型只负责把结果读出来、组织成人话——不是让模型自己猜哪里缺图。

### 3. 新增"图片补齐"的一次往返（方向与现有 `templateAssetSelection` 正好相反）

- 聊天消息下方渲染上传卡片：缺项 label + 目标模板 + 目标槽位 + 选择文件 / 粘贴。
  展示复用 `ImageAttachmentPreview`，不要新造一套预览组件。
- 用户选完文件后，前端直接调既有
  `domainSaveService.saveTemplateAsset(projectId, templateName, { usage, ... })`（`domainSaveService.ts:190`）。
  **不新增后端命令，不新增 DB migration。**
- 写入成功后在会话里回填一条系统消息（"已存入：项目X / 需求导入表 / 附件1"），并刷新上下文里的 `completion`，
  让下一轮 AI 看到它已经补齐。
- **触发方式先做方案 A**：前端按第 2 步的 completion 结果自动挂卡片。它不依赖模型行为、不依赖会话绑定、
  不需要审批，能独立跑通。方案 B（模型调 `request_project_image` 工具主动要图）留到会话绑定做完后再评估，
  本次不做。

### 4. 落库位置的边界情况要如实回执

`save_template_asset_internal` 已经处理了"项目没绑文件夹"的分支（落到 `.projects/<id>/assets/`），
但聊天里的用户看不到这个差别。卡片回执必须说清图片实际落在哪儿，
**不要一律说"已存到项目文件夹"**——项目没绑文件夹时那句话是假的。

### 5. 模板页表现必须同步

从聊天里补的图，回到模板页要立刻能看到：同一张表、同一个 usage、缩略图出现、进度条从 `X/11` 变 `X+1/11`。

**这是本次最重要的验收标准**：证明两个入口写的是同一份数据，不是两套并行的图片存储。

---

## 不要做的事

- **不做"AI 自动填写需求导入表字段"，不做"AI 直接调 `generate_lifecycle_docs` 出文档"。**
  那是写业务数据的工具，要走会话绑定 + 审批，是下一份任务书。本次 AI 只读、只提示，落库动作全部由用户点击触发。
- **不给聊天框里随手拖的图片做自动入库。** 只有"针对某个明确缺项槽位"的上传才落库；
  无槽位的图片维持现状（进模型、不进库）。否则用户随手贴的截图会污染项目资产库。
- **不新增 DB 表 / 迁移。** `project_template_assets` 够用，缺的是"怎么从聊天调用"，不是"缺列"。
- 不改 `save_template_asset_internal` 的存储路径规则、命名规则、20MB 上限和三种格式的限制。
- 不改 `docfill.rs` 的图片占位符机制（`ATTACH1_IMAGE` / `VENDOR_SCREENSHOT_LIST` 那套）。
- **不在 AI 侧重写第二套缺项判断规则**（见"已确认的关键事实"第 1 条）。
- 不碰审批弹窗、`GATED_TOOLS`、`agent_approval_log`。
- **本次只做需求导入表的 `attach1` / `attach2` 两个槽位。** 会审纪要的 vendor 截图、甄选结果签批表的图片
  不在范围内——先把一条路跑通，再按同一模式扩，不要一次铺开四张表。
- 不做 OCR、不做"从图片里自动抽字段回填表单"。
- 不改 `AiRuntime.ts` 的模型调用链路、不改 dsh 的 profile 配置。第 0 步只是**读出并记录**握手能力，不是去改它。

---

## 验证要求

1. **第 0 步的真实值**记录进 `docs/verification/`（新建文件），`true` / `false` 照实写，不要因为结论不方便就略过。
2. **端到端主链路**：打开一个 ICT 项目 → 聊天里问需求分析 → AI 说出缺附件1 → 上传 → 检查三处：
   - `project_template_assets` 有新行，且 `template_name` / `usage=attach1` / `project_id` 都对；
   - 项目文件夹下 `<项目名>-图片/assets/` 有对应物理文件；
   - 模板页需求 Tab 进度条 `X/11 → X+1/11`，缩略图可见。
3. **未绑文件夹分支单独验一次**：图片落到 `.projects/<id>/assets/`，且回执文案说的位置与实际一致。
4. **生成一次需求导入表 docx**，确认从聊天补的图真的出现在附件1的位置——这是"两个入口写同一份数据"的最终证据。
5. **失败路径**：> 20MB、GIF/BMP 等不支持格式、写库失败回滚（确认物理文件没有残留）。
6. **抽函数的回归证据**：第 1 步抽出前后，需求表 11 项逐项对比一次，进度条数字一致；
   其余三张表（会审 / 签批 / 甄选结果）的进度条数字不受影响。
7. `npm run lint --prefix src-ui`、`npm run build --prefix src-ui`、`cargo test` 全过，无回归。
8. 完工后更新 `docs/CURRENT_TASK.md` 与 `docs/CHANGELOG_AI.md`；照实区分"已验证"和"未验证 / 已知限制"，
   不要把没跑过的路径记成通过。
