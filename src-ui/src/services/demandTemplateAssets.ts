import { convertFileSrc } from '@tauri-apps/api/core';
import { emit } from '@tauri-apps/api/event';
import { listenBusinessEvent as listen } from './businessEvents';
import { domainSaveService } from './domainSaveService';
import { projectService } from '../utils/projectService';

export const DEMAND_ASSETS_CHANGED = 'lamber-demand-assets-changed';
export interface DemandAssetTarget { workspaceId: string; projectId: string; templateName: string }
export interface DemandImage { assetId?: string; data?: string; width?: number; height?: number; error?: boolean }

export async function loadDemandImages(projectId: string, templateName: string) {
  const assets = await domainSaveService.loadTemplateAssets(projectId, templateName);
  const result: Record<'attach1' | 'attach2', DemandImage[]> = { attach1: [], attach2: [] };
  // Stable oldest-first order matches template uploads and generated documents.
  for (const asset of assets.slice().reverse()) {
    if (asset.usage !== 'attach1' && asset.usage !== 'attach2') continue;
    let data = ''; let error = false;
    try { data = convertFileSrc(await projectService.getTemplateAssetPath(asset.id)); }
    catch { error = true; }
    result[asset.usage as 'attach1' | 'attach2'].push({ assetId: asset.id, data, width: asset.width, height: asset.height, error });
  }
  return result;
}

export function mergeDemandImages(legacy: DemandImage[], assets: DemandImage[]) {
  // Migrated images belong to the asset table; retain only unmigrated legacy data.
  return [...legacy.filter(img => !img.assetId), ...assets];
}

export async function publishDemandAssetsChanged(target: DemandAssetTarget) {
  window.dispatchEvent(new CustomEvent(DEMAND_ASSETS_CHANGED, { detail: target }));
  await emit(DEMAND_ASSETS_CHANGED, target).catch(console.warn);
}

export function subscribeDemandAssetsChanged(callback: (target: DemandAssetTarget) => void) {
  let disposed = false;
  const local = (event: Event) => callback((event as CustomEvent<DemandAssetTarget>).detail);
  window.addEventListener(DEMAND_ASSETS_CHANGED, local);
  const subscription = listen<DemandAssetTarget>(DEMAND_ASSETS_CHANGED, event => { if (!disposed) callback(event.payload); }).catch(error => { console.warn(error); return () => {}; });
  return () => {
    disposed = true;
    window.removeEventListener(DEMAND_ASSETS_CHANGED, local);
    void subscription.then(unlisten => unlisten()).catch(console.warn);
  };
}
