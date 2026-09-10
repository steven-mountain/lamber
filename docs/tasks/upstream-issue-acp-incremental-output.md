# 上游 issue 草稿 · dsh ACP 层无法投影 in-flight 增量输出

> **草稿，未提交。** 目标仓库：`deepseek-ai/deepseek-harness`。
> 提交前请自行核对：本文引用的行号来自本机 `node_modules` 里的**已构建产物**
> （`@deepseek-ai/dsh-acp@0.1.2-alpha.5` 的 `lib/index.js`），
> 上游仓库源码（`packages/acp/acp`）的行号会不同，建议改引函数名而非行号。
>
> 提交后把 issue 链接回填到 `docs/CURRENT_TASK.md`，并在
> `docs/verification/dsh-stage2-product-integration.md` 里标注它是这条缺失的跟进入口。

---

## Title

ACP: `agent_message_chunk` is only projected from committed messages, so clients cannot render incremental output

## Body

### Summary

The ACP bridge projects assistant output **only after a message is committed as a durable
event**. As a result an ACP client receives the entire assistant message in one (or a few
block-sized) `agent_message_chunk` notifications, arriving after generation has already
finished. There is no way for a client to render token-by-token output, even though the
harness itself streams token-level deltas internally.

This makes ACP clients feel unresponsive on long answers: the user sees nothing at all for
the whole generation, then the complete text appears at once.

### Environment

- `@deepseek-ai/dsh` / `@deepseek-ai/dsh-acp`: `0.1.2-alpha.5`
- ACP schema negotiated: `1.5.0`
- Client: Rust, `agent-client-protocol` crate, `dsh --profile acp` over stdio
- Provider/model: `deepseek-official` / `deepseek-v4-flash`

### What we observe

Sending one `session/prompt` that produces a multi-sentence answer, the client receives no
`session/update` carrying assistant text until generation is complete, then receives the
whole message. We confirmed the model itself is streaming: a transparent loopback proxy in
front of the DeepSeek API shows real upstream SSE deltas arriving well before any ACP
notification is emitted.

(We also cancel mid-generation successfully — `session/cancel` returns `Cancelled` in
11–18 ms — so the turn genuinely is still in flight while the client has been shown nothing.)

### Where it comes from

**The harness already appends every stream chunk to the session as an `assistant/chunk`
event, in real time, before the message is committed.** In `@deepseek-ai/dsh-agent-loop`:

```js
for await (const chunk of stream) {
    signal.throwIfAborted();
    chunkSeqs.push(this.session.append("assistant/chunk", { turn, step, chunk }).seq);
    assembler.push(chunk);
}
```

And the ACP bridge subscribes to the **full** session event firehose —
`ctx.on("session/event", ...)` — so it *does* receive those `assistant/chunk` events.
It simply doesn't project them: `onSessionEvent` branches only on
`assistant/message` and `tool/call`, and everything else is dropped.

In other words the data is already on the bus and already reaching the ACP layer;
only the projection is missing.

For completeness, in `@deepseek-ai/dsh-acp`:

- `onSessionEvent(session, event)` — documented as "Process one durable event and enqueue
  its standard ACP projections" — acts on `event.type === "assistant/message"`, i.e. a
  **committed** message.
- `assistantUpdates(ctx, session, event)` — documented as "Convert one committed assistant
  message ... in block order" — iterates `event.data.message.content`, an already-assembled
  block list, and emits one `agent_thought_chunk` / `agent_message_chunk` per block.

So the projection boundary is the commit, not the delta. A client can therefore never see
partial output, regardless of how it subscribes.

There is also no configuration escape hatch: the ACP plugin's config schema exposes only
`provider`, `model` and `sessionListPageSize`.

### Why this matters for ACP specifically

`agent_message_chunk` is the update ACP clients use to render streaming assistant output.
Delivering it only post-commit means the notification name and the client-side rendering
contract disagree: what arrives is a completed message, not a chunk.

### What we're asking for

Some way for an ACP client to receive in-flight assistant deltas. Any of these would solve
it for us, roughly in order of preference:

1. Project the `assistant/chunk` events the bridge already receives as
   `agent_message_chunk` / `agent_thought_chunk` while the turn is running, keeping the
   existing commit-time projection as the durability/replay path.
2. The same, behind an opt-in config flag (e.g. `streamPartialOutput: true`), so the current
   post-commit behaviour stays the default for consumers that depend on it.
3. If projecting partial output conflicts with the durable-event design on purpose, a
   documented statement to that effect would still help — we would then stop looking for a
   configuration and plan around it.

We recognise there may be a deliberate reason for the current boundary (replay determinism,
attachment integrity re-checks on commit, ordering guarantees against `tool_call` updates).
If so we'd appreciate understanding it; we're happy to help with a PR if a maintainer can
point at the intended shape.

### What we are not asking for

Not asking to change the durable event model, and we are not patching dsh locally — we want
to stay on upstream releases.

---

## 提交前的自检

- [ ] 行号改成函数名，或核对上游 `packages/acp/acp` 的真实行号
- [ ] 确认 `0.1.2-alpha.5` 仍是最新版；若已有新版，先在新版上复现一次再提
- [ ] 搜一遍现有 issue，避免重复提
- [ ] 移除任何本项目的业务信息（当前草稿已不含）
