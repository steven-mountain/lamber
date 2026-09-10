import catalogData from './catalog.json';

export type RequiredCondition = string | { field: string; equals: string | boolean | number };
export interface CompletionSource { field: string; source?: 'formData'; defaultValue?: string }
export interface CatalogField {
  key: string; label: string; kind: 'text' | 'derived' | 'check' | 'list' | 'image';
  defaultValue?: string; dynamicDefault?: boolean; requiredWhen?: RequiredCondition; stateKey?: string;
  reason?: string; columns?: string[]; completionAlways?: boolean;
  completionGroup?: string;
  completionSources?: CompletionSource[];
  validRow?: { nonEmpty: string[]; positive: string[] };
  listType?: 'editable' | 'generated';
}
export interface CatalogTemplate {
  id: string; name: string; suffix: string; excludedReason?: string; fields: CatalogField[];
}
export const templateCatalog = catalogData.templates as CatalogTemplate[];
export interface CompletionAsset { fieldKey?: string | null; exists?: boolean | null }
export interface CatalogCompletionItem {
  key: string; label: string; filled: boolean; evaluated: boolean;
  kind: 'field' | 'image'; templateName: string; usage?: 'attach1' | 'attach2';
}
const object = (value: unknown): Record<string, unknown> =>
  value && typeof value === 'object' && !Array.isArray(value) ? value as Record<string, unknown> : {};
export function getCatalogTemplate(name: string): CatalogTemplate | undefined {
  if (name.length > 256 || /[/\\\0:]/.test(name)) return undefined;
  const matches = templateCatalog.filter(t => !t.excludedReason && (name === t.name || (name.includes(t.name) && name.endsWith(t.suffix))));
  return matches.length === 1 ? matches[0] : undefined;
}
export function catalogTextValues(templateName: string, state: Record<string, unknown>): Record<string, string> {
  return Object.fromEntries((getCatalogTemplate(templateName)?.fields ?? []).filter(f => f.kind === 'text').flatMap(f => {
    const value = f.stateKey ? state[f.stateKey] : object(state.formData)[f.key];
    return typeof value === 'string' ? [[f.key, value]] : [];
  }));
}
/** Single completion algorithm for UI and plugin. Missing external business facts stay unknown. */
export function getCatalogCompletion(templateName: string, state: Record<string, unknown>, assets?: CompletionAsset[]): CatalogCompletionItem[] {
  const rules = getCatalogTemplate(templateName)?.fields ?? [];
  const items = rules.map((rule): CatalogCompletionItem => {
    const condition = rule.requiredWhen;
    const conditionKnown = !condition || typeof condition === 'string' || state[condition.field] != null;
    const required = !condition || (typeof condition === 'string' ? Boolean(state[condition]) : state[condition.field] === condition.equals);
    const supplied = object(state.completionValues)[rule.key];
    const raw = rule.stateKey ? state[rule.stateKey] : object(state.formData)[rule.key];
    let evaluated = true;
    let filled: boolean;
    if (!conditionKnown) { evaluated = false; filled = false; }
    else if (rule.completionSources) {
      const values = rule.completionSources.map(ref => (ref.source === 'formData' ? object(state.formData) : state)[ref.field] ?? ref.defaultValue);
      evaluated = !required || values.every(value => value != null);
      filled = !required || (evaluated && values.every(value => String(value).trim().length > 0));
    }
    else if (typeof supplied === 'boolean') filled = supplied;
    else if (rule.completionAlways) filled = true;
    else if (rule.kind === 'derived' || rule.kind === 'check') { evaluated = false; filled = false; }
    else if (rule.kind === 'image') {
      const legacy = state[`${rule.key}Images`];
      const present = assets ? assets.some(a => a.fieldKey === rule.key && a.exists === true)
        : Array.isArray(legacy) && legacy.some(img => !object(img).error);
      filled = !required || present;
    } else if (rule.kind === 'list') {
      const rows = state[rule.key];
      filled = !required || (Array.isArray(rows) && rows.some(row => !rule.validRow
        || rule.validRow.nonEmpty.some(key => String(object(row)[key] ?? '').trim().length > 0)
        || rule.validRow.positive.some(key => Number(object(row)[key] ?? 0) > 0)));
    }
    else {
      evaluated = !required || raw != null || !rule.dynamicDefault;
      filled = !required || String(raw ?? rule.defaultValue ?? '').trim().length > 0;
    }
    return {key:rule.key,label:rule.label,filled,evaluated,kind:rule.kind === 'image' ? 'image' : 'field',templateName,
      ...(rule.kind === 'image' ? {usage:rule.key as 'attach1' | 'attach2'} : {})};
  });
  // Several writable keys can jointly occupy one completion slot. First declaration owns order/label.
  const grouped = new Map<string, CatalogCompletionItem>();
  return items.filter((item, index) => {
    const group = rules[index]?.completionGroup;
    if (!group) return true;
    const first = grouped.get(group);
    if (!first) { grouped.set(group, item); return true; }
    first.filled = first.filled && item.filled;
    first.evaluated = first.evaluated && item.evaluated;
    return false;
  });
}
