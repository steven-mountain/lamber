export interface ReviewField { key: string; label: string; previousValue: string | null; proposedValue: string }
export interface WriteIntent { projectName: string; templateName: string; targetDescription: string; fields: ReviewField[] }
export interface ApprovalPrompt {
  requestId: string; toolName: string; callId: string | null; sessionId: string | null;
  reason: string | null; args: unknown; timeoutSeconds: number; expiresAt: string; intent: WriteIntent | null;
}
export function remainingSeconds(expiresAt: string, now = Date.now()): number {
  const deadline = Date.parse(expiresAt);
  return Number.isFinite(deadline) ? Math.max(0, Math.ceil((deadline - now) / 1000)) : 0;
}
export function amendedArguments(prompt: ApprovalPrompt, values: Record<string, string>): Record<string, unknown> | undefined {
  if (!prompt.intent) return undefined;
  const original = prompt.args && typeof prompt.args === 'object' && !Array.isArray(prompt.args) ? prompt.args : {};
  if (prompt.intent.fields.every(field => values[field.key] === field.proposedValue)) return undefined;
  const fields = Object.fromEntries(prompt.intent.fields.map(field => [field.key, values[field.key] ?? field.proposedValue]));
  return prompt.toolName === 'fill_template_fields' ? { ...original, fields } : { ...original, ...fields };
}
/** Read-only fallback for tools without an editable intent. Preserve prose/newlines. */
export function readableArguments(value: unknown, path = '参数'): { label: string; text: string }[] {
  if (value && typeof value === 'object') return Object.entries(value).flatMap(([key, item]) => readableArguments(item, path === '参数' ? key : `${path} · ${key}`));
  return [{label:path, text: value == null ? '（无）' : String(value)}];
}
