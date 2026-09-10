import type { AiChatMessage, AiToolCall } from './types';

/** ACP text is already decoded UTF-8. Never parse it as SSE or think-tag markup. */
export function reduceDshUpdate(message: AiChatMessage, update: Record<string, unknown>): AiChatMessage {
  const content = update.content as { type?: string; text?: string } | undefined;
  switch (update.sessionUpdate) {
    case 'agent_message_chunk':
      return content?.type === 'text' && typeof content.text === 'string'
        ? { ...message, content: message.content + content.text } : message;
    case 'agent_thought_chunk':
      return content?.type === 'text' && typeof content.text === 'string'
        ? { ...message, think: (message.think ?? '') + content.text } : message;
    case 'tool_call':
    case 'tool_call_update': {
      if (typeof update.toolCallId !== 'string') return message;
      const calls = [...(message.toolCalls ?? [])];
      const index = calls.findIndex(call => call.id === update.toolCallId);
      const prior = index < 0 ? undefined : calls[index];
      const call: AiToolCall = {
        id: update.toolCallId,
        title: typeof update.title === 'string' ? update.title : prior?.title ?? '工具调用',
        status: typeof update.status === 'string' ? update.status : prior?.status ?? 'pending',
        input: update.rawInput === undefined ? prior?.input : update.rawInput,
        output: update.content === undefined ? prior?.output : update.content,
      };
      if (index < 0) calls.push(call); else calls[index] = call;
      return { ...message, toolCalls: calls };
    }
    default: return message;
  }
}

interface Block {
  type: string;
  text: string;
  decoder: TextDecoder;
  id?: string;
  name?: string;
}
interface Step {
  turn: number;
  step: number;
  seq: number;
  messageId?: string;
  blocks: Map<number, Block>;
}

/** One request owns the correlation/decoder state; only rendered messages leave it. */
export class DshMessageProjection {
  private authoritative: AiChatMessage = { role: 'assistant', content: '' };
  private messages = new Map<string, AiChatMessage>();
  private steps = new Map<string, Step>();

  update(update: Record<string, unknown>): AiChatMessage {
    this.authoritative = reduceDshUpdate(this.authoritative, update);
    if (typeof update.messageId === 'string') {
      const prior = this.messages.get(update.messageId) ?? { role: 'assistant', content: '' };
      this.messages.set(update.messageId, reduceDshUpdate(prior, update));
    }
    return this.render();
  }

  stream(frame: Record<string, unknown>): AiChatMessage {
    const { turn, step, seq } = frame;
    if (!Number.isSafeInteger(turn) || !Number.isSafeInteger(step) || !Number.isSafeInteger(seq)) return this.render();
    const key = `${turn}:${step}`;
    const current = this.steps.get(key) ?? { turn: turn as number, step: step as number, seq: -1, blocks: new Map<number, Block>() };
    if ((seq as number) <= current.seq) return this.render();
    current.seq = seq as number;
    this.steps.set(key, current);
    if (frame.kind === 'commit' && typeof frame.messageId === 'string') {
      current.messageId = frame.messageId;
    } else if (frame.kind === 'delta' && !current.messageId) {
      const chunk = frame.chunk as Record<string, unknown> | undefined;
      if (!chunk || !Number.isSafeInteger(chunk.index) || !Array.isArray(chunk.bytes)
        || !chunk.bytes.every(byte => Number.isInteger(byte) && byte >= 0 && byte <= 255)
        || !['text-delta', 'reasoning-delta', 'tool-call-delta'].includes(String(chunk.type))) return this.render();
      const index = chunk.index as number;
      const block = current.blocks.get(index) ?? { type: String(chunk.type), text: '', decoder: new TextDecoder('utf-8', { fatal: true }) };
      if (block.type !== chunk.type) return this.render();
      // Keep incomplete code points in the decoder until the following chunk.
      block.text += block.decoder.decode(new Uint8Array(chunk.bytes), { stream: true });
      if (typeof chunk.id === 'string') block.id = chunk.id;
      if (typeof chunk.name === 'string') block.name = chunk.name;
      current.blocks.set(index, block);
    }
    return this.render();
  }

  finish(): AiChatMessage {
    // ACP precedes the terminal event. Replace the whole preview, including
    // orphaned tool argument drafts, even when HTTP commit metadata arrives late.
    this.steps.clear();
    this.messages.clear();
    return this.authoritative;
  }

  private render(): AiChatMessage {
    if (this.steps.size === 0) return this.authoritative;
    let content = '', think = '';
    const previews = new Map<string, AiToolCall>();
    for (const step of [...this.steps.values()].sort((a, b) => a.turn - b.turn || a.step - b.step)) {
      let text = '', reasoning = '';
      for (const block of [...step.blocks.entries()].sort(([a], [b]) => a - b).map(([, block]) => block)) {
        if (block.type === 'text-delta') text += block.text;
        else if (block.type === 'reasoning-delta') reasoning += block.text;
        else if (block.id) previews.set(block.id, { id: block.id, title: block.name ?? '工具调用', status: 'pending', input: block.text });
      }
      const committed = step.messageId ? this.messages.get(step.messageId) : undefined;
      content += committed ? committed.content : text;
      think += committed ? committed.think ?? '' : reasoning;
    }
    // Actual execution status and parsed arguments only ever come from ACP.
    for (const call of this.authoritative.toolCalls ?? []) previews.set(call.id, call);
    return { role: 'assistant', content, think, toolCalls: [...previews.values()] };
  }
}
