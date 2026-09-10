import type { AiChatMessage, ContextNode } from './types';

/** Harness history excludes local card receipts. Supply only this session's application events. */
export function appReceiptContext(messages: readonly AiChatMessage[]): ContextNode[] {
  const receipts = messages.filter(message => message.role === 'assistant' && message.appReceipt === true).slice(-8);
  return receipts.length ? [{
    type: 'json',
    title: '本会话应用操作回执（按发生顺序；历史成功不代表最近失败的操作成功）',
    content: receipts.map(message => ({ result: message.content })),
    metadata: { module: 'app_receipts' },
  }] : [];
}
