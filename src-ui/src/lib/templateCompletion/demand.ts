import { getCatalogCompletion } from './catalog';
export type { CompletionAsset, CatalogCompletionItem as DemandCompletionItem } from './catalog';
/** Compatibility entry point; all templates share the catalog algorithm. */
export const getDemandCompletion = getCatalogCompletion;
