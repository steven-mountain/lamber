export type AiContextScope = 'hub' | 'benefit' | 'docfill' | 'ict';
export type AiContextKind = 'core' | 'template';

export const AI_CONTEXT_KEY = {
  HUB: 'hub',
  BENEFIT_CORE: 'benefit.core',
  DOCFILL_CORE: 'docfill.core',
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
  if (key === AI_CONTEXT_KEY.BENEFIT_CORE || key.startsWith(`benefit${TEMPLATE_SEPARATOR}`)) return 'benefit';
  if (key === AI_CONTEXT_KEY.DOCFILL_CORE || key.startsWith(`docfill${TEMPLATE_SEPARATOR}`)) return 'docfill';
  if (key === AI_CONTEXT_KEY.ICT_CORE || key.startsWith(`ict${TEMPLATE_SEPARATOR}`)) return 'ict';
  return null;
}

export function isAiContextKeyForView(key: string, view: string): boolean {
  if (view === 'hub') return false;

  if (view === 'benefit') {
    return key === AI_CONTEXT_KEY.BENEFIT_CORE;
  }

  if (view === 'docfill') {
    return key === AI_CONTEXT_KEY.DOCFILL_CORE || key.startsWith(`docfill${TEMPLATE_SEPARATOR}`);
  }

  if (view === 'ict') {
    return key === AI_CONTEXT_KEY.ICT_CORE || key.startsWith(`ict${TEMPLATE_SEPARATOR}`);
  }

  return false;
}

export function getAiContextDisplayName(key: string): string {
  if (key === AI_CONTEXT_KEY.BENEFIT_CORE) return 'Benefit Analysis';
  if (key === AI_CONTEXT_KEY.DOCFILL_CORE) return 'Docfill';
  if (key === AI_CONTEXT_KEY.ICT_CORE) return 'ICT Lifecycle';
  if (key.startsWith(`docfill${TEMPLATE_SEPARATOR}`)) {
    return `Docfill Template: ${key.slice(`docfill${TEMPLATE_SEPARATOR}`.length)}`;
  }
  if (key.startsWith(`ict${TEMPLATE_SEPARATOR}`)) {
    return `ICT Template: ${key.slice(`ict${TEMPLATE_SEPARATOR}`.length)}`;
  }
  return key;
}

export function migrateLegacyAiContextKey(key: string): string | null {
  if (key === 'ict') return AI_CONTEXT_KEY.ICT_CORE;
  if (key === 'docfill') return AI_CONTEXT_KEY.DOCFILL_CORE;
  if (key.startsWith('template_')) return null;
  if (getAiContextScope(key)) return key;
  return null;
}

export function normalizeAiActiveModule(module: unknown): string {
  if (typeof module !== 'string') return AI_CONTEXT_KEY.HUB;
  if (module.startsWith('template_')) return AI_CONTEXT_KEY.HUB;
  if (module === 'ict') return AI_CONTEXT_KEY.ICT_CORE;
  if (module === 'benefit') return AI_CONTEXT_KEY.BENEFIT_CORE;
  if (module === 'docfill') return AI_CONTEXT_KEY.DOCFILL_CORE;
  return getAiContextScope(module) ? module : AI_CONTEXT_KEY.HUB;
}
