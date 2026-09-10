import { TEMPLATE_FIELD_DEFAULTS } from './templateGenerationState';
/** Three-way merge: approved text can replace only an unchanged local value. */
export function mergeApprovedText(base: Record<string, unknown>, local: Record<string, string>, incoming: Record<string,string>) {
  const merged = {...local}; const conflicts: string[]=[];
  for (const [key,value] of Object.entries(incoming)) {
    const before = base[key] ?? TEMPLATE_FIELD_DEFAULTS[key] ?? '';
    const current = local[key] ?? TEMPLATE_FIELD_DEFAULTS[key] ?? '';
    if (current !== before && current !== value) conflicts.push(key);
    else merged[key]=value;
  }
  return {merged,conflicts};
}
