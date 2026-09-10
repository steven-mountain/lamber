import { invoke } from '@tauri-apps/api/core';
import { DshMessageProjection } from './DshMessageProjection';
import { listen } from '@tauri-apps/api/event';
import type { AiChatMessage, AiImageAttachment } from './types';

export interface DshEvent {
  method: string;
  lamberSessionId?: string;
  requestId?: string;
  params: { [key: string]: unknown; sessionId?: string; update?: Record<string, unknown>; error?: string; stopReason?: string };
}

export { reduceDshUpdate } from './DshMessageProjection';

export interface DshTransport {
  listen: (handler: (event: DshEvent) => void) => Promise<() => void>;
  invoke: <T>(command: string, args: Record<string, unknown>) => Promise<T>;
}
const transport: DshTransport = {
  listen: handler => listen<DshEvent>('ai://session-event', event => handler(event.payload)),
  invoke,
};

export function imageBlocks(images: AiImageAttachment[]) {
  return images.map(image => {
    const match = image.dataUrl?.match(/^data:(image\/(?:png|jpeg|webp));base64,([A-Za-z0-9+/=]+)$/);
    if (!match) throw new Error(`图片「${image.name}」数据不可用，请重新添加`);
    return { mimeType: match[1], data: match[2] };
  });
}

export interface DshTurn {
  sessionId: string;
  requestId: string;
  harnessSessionId?: string;
  text: string;
  images?: AiImageAttachment[];
  onUpdate: (message: AiChatMessage) => void;
  onSession: (id: string) => void;
  signal: AbortSignal;
}

export class DshRuntime {
  // Diagnostic metadata only, deliberately bounded and excluded from chat/persistence.
  readonly diagnostics: Array<Record<string, unknown>> = [];
  constructor(private readonly bridge: DshTransport = transport) {}

  async execute(turn: DshTurn): Promise<void> {
    const images = imageBlocks(turn.images ?? []);
    const projection = new DshMessageProjection();
    let settled = false;
    let resolveEnd!: () => void;
    let rejectEnd!: (error: Error) => void;
    const ended = new Promise<void>((resolve, reject) => { resolveEnd = resolve; rejectEnd = reject; });
    // Attach immediately: a terminal event may arrive before invoke resolves.
    void ended.catch(() => {});
    const unlisten = await this.bridge.listen(event => {
      if (settled || event.requestId !== turn.requestId || event.lamberSessionId !== turn.sessionId) return;
      if (event.method === 'session/update' && event.params.update) {
        const update = event.params.update;
        if (update.sessionUpdate === 'usage_update' || update.sessionUpdate === 'config_option_update') {
          this.diagnostics.push(update);
          if (this.diagnostics.length > 100) this.diagnostics.shift();
        }
        turn.onUpdate(projection.update(update));
      } else if (event.method === 'session/stream') {
        if (event.params.kind === 'error') {
          this.diagnostics.push({ error: event.params.error });
          if (this.diagnostics.length > 100) this.diagnostics.shift();
        } else turn.onUpdate(projection.stream(event.params));
      } else if (event.method === 'session/turn-ended') {
        if (!event.params.error) turn.onUpdate(projection.finish());
        settled = true;
        if (event.params.error) rejectEnd(new Error(event.params.error)); else resolveEnd();
      }
    });
    let queued = false;
    const cancel = () => {
      if (!queued || settled) return;
      void this.bridge.invoke('ai_cancel_prompt', { sessionId: turn.sessionId, requestId: turn.requestId })
        .catch(error => rejectEnd(new Error(`停止生成失败：${String(error)}`)));
    };
    turn.signal.addEventListener('abort', cancel, { once: true });
    try {
      if (turn.signal.aborted) return;
      const id = await this.bridge.invoke<string>('ai_send_prompt', {
        sessionId: turn.sessionId, requestId: turn.requestId, text: turn.text,
        harnessSessionId: turn.harnessSessionId, images,
      });
      queued = true;
      turn.onSession(id);
      if (turn.signal.aborted) cancel();
      await ended;
    } finally {
      settled = true;
      turn.signal.removeEventListener('abort', cancel);
      unlisten();
    }
  }
}
