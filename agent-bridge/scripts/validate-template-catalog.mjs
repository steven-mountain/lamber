const safeKey = value => typeof value === 'string' && /^[a-zA-Z][a-zA-Z0-9_]*$/.test(value) && !['constructor','prototype','__proto__'].includes(value);
const onlyKeys = (value, keys) => value && typeof value === 'object' && !Array.isArray(value) && Object.keys(value).every(key => keys.includes(key));
export function validateCatalog(catalog) {
  for (const template of catalog.templates) {
    const keys = new Set(template.fields.map(f => f.key));
    const groups = new Map();
    for (const field of template.fields) {
      if (!['text','derived','check','list','image'].includes(field.kind)) throw new Error('Invalid kind');
      const condition = field.requiredWhen;
      if (condition !== undefined && !(safeKey(condition) || (onlyKeys(condition, ['field','equals']) && safeKey(condition.field)
        && ['string','boolean','number'].includes(typeof condition.equals) && (typeof condition.equals !== 'number' || Number.isFinite(condition.equals))))) throw new Error('Invalid requiredWhen');
      if (field.completionGroup !== undefined) {
        if (!safeKey(field.completionGroup) || field.kind !== 'text') throw new Error('Invalid completion group');
        groups.set(field.completionGroup, (groups.get(field.completionGroup) || 0) + 1);
      }
      if (field.completionSources !== undefined && (!Array.isArray(field.completionSources) || !field.completionSources.length || field.completionSources.length > 8
        || field.completionSources.some(ref => !onlyKeys(ref,['field','source','defaultValue']) || !safeKey(ref.field)
          || (ref.source !== undefined && ref.source !== 'formData') || (ref.defaultValue !== undefined && typeof ref.defaultValue !== 'string')))) throw new Error('Invalid completion sources');
      if (field.listType !== undefined && (field.kind !== 'list' || !['editable','generated'].includes(field.listType) || !field.reason)) throw new Error('Invalid list type');
      if (field.validRow !== undefined) {
        const rule = field.validRow;
        if (field.kind !== 'list' || !onlyKeys(rule,['nonEmpty','positive']) || !Array.isArray(rule.nonEmpty) || !Array.isArray(rule.positive)
          || ![...rule.nonEmpty,...rule.positive].length || [...rule.nonEmpty,...rule.positive].length > 8
          || [...rule.nonEmpty,...rule.positive].some(key => !safeKey(key) || !field.columns?.includes(key))) throw new Error('Invalid valid row');
      }
      if (!safeKey(field.key) || keys.size !== template.fields.length) throw new Error('Invalid field keys');
    }
    if ([...groups.values()].some(count => count < 2)) throw new Error('Completion groups must contain multiple fields');
  }
}
