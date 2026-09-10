/** Temporary alpha.5 display seam. ACP remains authoritative; never executes tools. */
import type { Context } from '@deepseek-ai/cordis';
import { postBridge } from './bridge.js';

export const STREAM_ROUTE = '/lamber-bridge/stream';
type Binding = { lamberSessionId: string; requestId: string | null };
type Stream = { turn: number; binding: Binding; tail: Promise<void>; pending: number; failed: boolean; surrogates: Map<string, string> };

/** Preserve a trailing UTF-16 high surrogate before encoding the byte transport. */
function bytes(state: Stream, key: string, text: string): number[] {
  let value = (state.surrogates.get(key) ?? '') + text;
  state.surrogates.delete(key);
  const last = value.charCodeAt(value.length - 1);
  if (last >= 0xd800 && last <= 0xdbff) {
    state.surrogates.set(key, value.slice(-1));
    value = value.slice(0, -1);
  }
  return Array.from(Buffer.from(value, 'utf8'));
}

export function applyStreaming(ctx: Context): void {
  if (process.env.LAMBER_STREAM_DISPLAY !== '1') return;
  const streams = new Map<string, Stream>();
  // Await the binding before model execution. A cancelled old turn can never
  // obtain the binding of a later prompt, even if its HTTP queue drains late.
  ctx.on('agent/pre-step', async ({ agent, turn, signal }, next) => {
    if (streams.get(agent.session.id)?.turn !== turn) {
      const binding = await postBridge<Binding>(STREAM_ROUTE,
        { kind: 'bind', sessionId: agent.session.id }, AbortSignal.any([signal, AbortSignal.timeout(5000)]));
      streams.set(agent.session.id, { turn, binding, tail: Promise.resolve(), pending: 0, failed: false, surrogates: new Map() });
    }
    return next();
  });
  ctx.on('session/event', (session, event) => {
    const state = streams.get(session.id);
    if (!state || !('turn' in event.data) || event.data.turn !== state.turn || state.failed) return;
    let payload: Record<string, unknown>;
    if (event.type === 'assistant/chunk') {
      const chunk = event.data.chunk;
      if (chunk.type === 'text-delta' || chunk.type === 'reasoning-delta') {
        payload = { kind: 'delta', step: event.data.step, chunk: { type: chunk.type, index: chunk.index,
          bytes: bytes(state, `${event.data.step}:${chunk.index}`, chunk.text) } };
      } else if (chunk.type === 'tool-call-delta') {
        payload = { kind: 'delta', step: event.data.step, chunk: { type: chunk.type, index: chunk.index,
          id: chunk.id, name: chunk.name, bytes: bytes(state, `${event.data.step}:${chunk.index}`, chunk.argumentsDelta) } };
      } else return;
    } else if (event.type === 'assistant/message') {
      payload = { kind: 'commit', step: event.data.step, messageId: event.data.message.id };
      state.surrogates.clear();
    } else return;
    // Sequential requests preserve event order without racing 16 HTTP workers.
    // Bound backlog and fail visibly instead of retaining unlimited generated text.
    state.pending++;
    if (state.pending > 2048) state.failed = true;
    const envelope = { ...payload, sessionId: session.id, ...state.binding, turn: state.turn, seq: event.seq };
    state.tail = state.tail.then(async () => {
      if (state.failed) return;
      await postBridge(STREAM_ROUTE, envelope, AbortSignal.timeout(5000));
    }).catch(() => {
      state.failed = true;
      ctx.logger.warn('lamber: incremental display bridge failed; ACP output remains authoritative');
      // A small diagnostic, never the user's content or bridge credential.
      return postBridge(STREAM_ROUTE, { ...envelope, kind: 'error', chunk: undefined,
        error: '实时显示通道失败，当前回答将以最终提交内容显示' }, AbortSignal.timeout(5000)).then(() => {}, () => {});
    }).finally(() => { state.pending--; });
    if (state.failed) {
      ctx.logger.warn('lamber: incremental display backlog exceeded 2048 events');
      void postBridge(STREAM_ROUTE, { ...envelope, kind: 'error', chunk: undefined,
        error: '实时显示积压超限，当前回答将以最终提交内容显示' }, AbortSignal.timeout(5000)).catch(() => {});
    }
  });
  ctx.on('session/disposed', session => { streams.delete(session.id); });
}
