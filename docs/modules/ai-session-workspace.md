# AI 多 Session 会话工作区

## 1. 模块定位

Lamber AI 窗口使用前端 Session 容器管理多段相互隔离的对话。该结构只负责界面、会话消息和本机持久化，当前消息执行仍完整复用 `AiRuntime` 的 OpenAI-compatible SSE 调用链。

本阶段不接入 deepseek-harness，不启动 `dsh` 子进程，不增加 Agent Tool、Approval、Sub-Agent、JSON-RPC、Rust bridge 或服务端 Session 持久化。

## 2. 数据契约

`AiSession` 位于 `src-ui/src/ai/sessionTypes.ts`：

```ts
interface AiSession {
  id: string
  title: string
  projectId?: string
  harnessSessionId?: string
  createdAt: number
  updatedAt: number
  messages: AiChatMessage[]
  titleSource?: 'default' | 'manual' | 'generated'
}
```

- `projectId` 是 Lamber 项目归属元数据；新会话在存在当前项目时自动关联该项目。
- `harnessSessionId` 仅作为未来一对一映射接缝，当前不得参与任何执行逻辑。
- `titleSource` 为后续自动标题预留来源标记；当前不调用模型生成标题。
- `currentSessionId` 与 Session 列表一同保存在版本化 localStorage 快照 `lamber_ai_session_workspace` 中。

## 3. 状态所有权与流式隔离

- `useAiSessionStore` 是 Session 与消息的唯一前端真相源；`AiChatPanel` 不再维护独立的消息数组。
- 发送时先固定发起请求的 `sessionId`，再把该 Session 已有消息作为 `AiRuntime.execute(...)` 的 history。
- 流式解析器的输出始终按该固定 `sessionId` 定向写回。用户在生成期间切换到其他 Session 时，不得把流式内容写入当前可见 Session。
- 当前仍只有一个 `AiRuntime` 和一条前端生成通道；切到其他 Session 时可查看历史和停止当前生成，但在当前请求结束前不并发发起第二个请求。
- 停止生成继续使用原 `AbortController` 与 streaming parser，不改变 SSE、PromptRenderer 或业务上下文构建路径。

## 4. 持久化规则

- 创建、选择、重命名、删除、清空和标题更新立即持久化；流式消息使用短节流批量写入，避免每个 token 都同步写 localStorage。
- 删除会话必须经过用户确认；删除当前会话后自动选择最近更新的剩余会话，删除最后一个会话后由聊天面板创建新的空白会话，保证工作区始终可用。
- 如果被删除的会话正在流式生成，先复用既有 Abort / parser 停止链路，再移除会话，避免迟到的流式片段继续写入已删除记录。
- 页面退出时主动 flush 最新快照，重启后恢复 Session 列表、当前选择和聊天文本。
- 用户上传图片的 `dataUrl` 只保留在当前运行期，持久化时保留附件元数据并移除 base64；这是为了避免单张 5MB 图片超过 localStorage 配额并导致全部会话快照写入失败。
- 不写入 Workspace SQLite、`projects_store.json`、项目目录或 Rust 数据库 schema。

## 5. 界面与响应式规则

- 默认宽窗为 Session Sidebar + Chat Area；Sidebar 使用约 `216px` 的低对比 surface。
- 当前项目 Session 优先分组展示；其余项目与通用 Session 进入“其他 / 通用会话”。
- Session 按 `updatedAt` 降序展示，active 使用 `bg-card + shadow-sm`，hover 使用轻量 surface shift。
- 每个 Session 的轻量操作菜单提供重命名与删除。重命名在列表内完成，`Enter` 或失焦保存，`Escape` 取消；本阶段不提供项目归属移动。
- 小于 `680px` 时 Sidebar 改为覆盖式抽屉，不占用输入区宽度；主区域始终保留展开入口。
- 新布局继续消费现有 semantic token、主题、字体、圆角和 dark mode，不引入独立视觉体系或高饱和大面积背景。

## 6. 后续 Harness 接入边界

未来接入 Harness 时，可在现有 `AiSession.id` 与 `harnessSessionId` 之间建立一对一映射，并替换 Session 的执行适配器。不得让 Harness 接入反向污染本阶段的 UI store，也不得绕过用户操作直接写 Lamber 核心项目数据。
