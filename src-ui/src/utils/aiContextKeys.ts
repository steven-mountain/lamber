export type AiContextScope = 'hub' | 'ict';
export type AiContextKind = 'core' | 'template';

export const AI_CONTEXT_KEY = {
  HUB: 'hub',
  ICT_CORE: 'ict.core',
} as const;

const TEMPLATE_SEPARATOR = '.template.';

export function sanitizeAiContextId(id: string): string {
  const sanitized = id
    .trim()
    .replace(/[\\/\s:.]+/g, '_')
    .replace(/[^\w\u4e00-\u9fff-]+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '');

  return sanitized || 'untitled';
}

export function buildAiContextKey(scope: AiContextScope, kind: AiContextKind = 'core', id?: string): string {
  if (scope === 'hub') return AI_CONTEXT_KEY.HUB;
  if (kind === 'core') return `${scope}.core`;

  return `${scope}${TEMPLATE_SEPARATOR}${sanitizeAiContextId(id || 'untitled')}`;
}

export function getAiContextScope(key: string): AiContextScope | null {
  if (key === AI_CONTEXT_KEY.HUB) return 'hub';
  if (key === AI_CONTEXT_KEY.ICT_CORE || key.startsWith(`ict${TEMPLATE_SEPARATOR}`)) return 'ict';
  return null;
}

export function isAiContextKeyForView(key: string, view: string): boolean {
  if (view === 'hub') return false;

  if (view === 'ict') {
    return key === AI_CONTEXT_KEY.ICT_CORE || key.startsWith(`ict${TEMPLATE_SEPARATOR}`);
  }

  return false;
}

export function getAiContextDisplayName(key: string): string {
  if (key === AI_CONTEXT_KEY.ICT_CORE) return 'ICT Lifecycle';
  if (key.startsWith(`ict${TEMPLATE_SEPARATOR}`)) {
    return `ICT Template: ${key.slice(`ict${TEMPLATE_SEPARATOR}`.length)}`;
  }
  return key;
}

export function migrateLegacyAiContextKey(key: string): string | null {
  if (key === 'ict') return AI_CONTEXT_KEY.ICT_CORE;
  if (key === 'benefit' || key === 'benefit.core' || key.startsWith(`benefit${TEMPLATE_SEPARATOR}`)) return null;
  if (key === 'docfill' || key === 'docfill.core' || key.startsWith(`docfill${TEMPLATE_SEPARATOR}`)) return null;
  if (key.startsWith('template_')) return null;
  if (getAiContextScope(key)) return key;
  return null;
}

export function normalizeAiActiveModule(module: unknown): string {
  if (typeof module !== 'string') return AI_CONTEXT_KEY.HUB;
  if (module.startsWith('template_')) return AI_CONTEXT_KEY.HUB;
  if (module === 'ict') return AI_CONTEXT_KEY.ICT_CORE;
  return getAiContextScope(module) ? module : AI_CONTEXT_KEY.HUB;
}
